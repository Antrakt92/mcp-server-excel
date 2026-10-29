"""Bounded Windows stdio smoke of a published MCP executable; never opens Excel."""

import argparse
import ctypes
import hashlib
import json
import os
from pathlib import Path
import queue
import subprocess
import sys
import tempfile
import threading
import time


def require(condition, message):
    if not condition:
        raise RuntimeError(message)


class Reader:
    def __init__(self, stream, *, protocol=False, stop_id=None):
        self.stream = stream
        self.protocol = protocol
        self.stop_id = stop_id
        self.items = queue.Queue(maxsize=128)
        self.failure = None
        self.thread = threading.Thread(target=self.read, daemon=True)
        self.thread.start()

    def read(self):
        try:
            for line in self.stream:
                value = json.loads(line) if self.protocol else None
                if self.protocol:
                    require(isinstance(value, dict) and value.get("jsonrpc") == "2.0", "Non-protocol stdout")
                    self.items.put_nowait(value)
                    if self.stop_id is not None and value.get("id") == self.stop_id:
                        return
                # Stderr is drained but never copied into the public report.
        except BaseException as error:
            self.failure = error

    def response(self, request_id, *, timeout_seconds=8):
        deadline = time.monotonic() + timeout_seconds
        while time.monotonic() < deadline:
            if self.failure:
                raise RuntimeError("Protocol reader failed") from self.failure
            try:
                item = self.items.get(timeout=min(0.1, max(0.01, deadline - time.monotonic())))
            except queue.Empty:
                continue
            if item.get("id") == request_id:
                require("error" not in item, "Unexpected JSON-RPC error")
                return item["result"]
        raise TimeoutError("MCP response deadline expired")


def start(exe, *, stop_id=None):
    process = subprocess.Popen(
        [str(exe)], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True, encoding="utf-8", env=dict(os.environ, EXCELMCP_TELEMETRY_MODE="off"),
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    return process, Reader(process.stdout, protocol=True, stop_id=stop_id), Reader(process.stderr)


def send(process, request_id, method, params=None):
    message = {"jsonrpc": "2.0", "method": method}
    if request_id is not None:
        message["id"] = request_id
    if params is not None:
        message["params"] = params
    process.stdin.write(json.dumps(message) + "\n")
    process.stdin.flush()


def initialize(process, reader):
    send(process, 1, "initialize", {
        "protocolVersion": "2024-11-05", "capabilities": {},
        "clientInfo": {"name": "fork-published-protocol-smoke", "version": "1"},
    })
    result = reader.response(1)
    require(bool(result.get("protocolVersion")), "Missing negotiated protocol version")
    send(process, None, "notifications/initialized")
    send(process, 2, "tools/list", {})
    tools = reader.response(2)["tools"]
    require(bool(tools) and "file_read" in {tool["name"] for tool in tools}, "Missing file_read tool")
    return len(tools)


def tool(process, reader, request_id, arguments, *, expected_error=False, timeout_seconds=8):
    send(process, request_id, "tools/call", {"name": "file_read", "arguments": arguments})
    result = reader.response(request_id, timeout_seconds=timeout_seconds)
    text = [entry["text"] for entry in result["content"] if entry.get("type") == "text"]
    require(len(text) == 1, "Expected one JSON text result")
    structured = result.get("structuredContent")
    require(isinstance(structured, dict) and json.loads(text[0]) == structured, "Text/structured result mismatch")
    require((result.get("isError") is True) == expected_error, "Unexpected diagnostic error classification")
    return structured


def cleanup(process, readers, receipt=None):
    primary = sys.exception()
    receipt = receipt if receipt is not None else {}
    receipt.update(owned_process_id=process.pid, forced_termination=False,
                   exit_confirmed=False, no_lingering_owned_host=False)
    failures = []
    details = receipt.setdefault("failures", [])

    def record(phase, error):
        failures.append(error)
        details.append({"phase": phase, "exception_type": type(error).__name__})

    try:
        if process.poll() is None:
            receipt["forced_termination"] = True
            process.kill()  # Popen retains the exact Windows process handle.
    except BaseException as error:
        record("terminate_owned_host", error)
    try:
        process.wait(timeout=3)
        receipt["exit_confirmed"] = process.poll() is not None
        receipt["no_lingering_owned_host"] = receipt["exit_confirmed"]
        receipt["exit_code"] = process.returncode
    except BaseException as error:
        record("wait_owned_host", error)
    for reader in readers:
        try:
            reader.thread.join(timeout=2)
            if reader.thread.is_alive():
                record("join_owned_reader", RuntimeError("Owned pipe reader did not stop"))
            if reader.failure:
                record("owned_reader", reader.failure)
        except BaseException as error:
            record("join_owned_reader", error)
    for stream in (process.stdin, process.stdout, process.stderr):
        try:
            # A stuck reader must not turn close() into an unbounded lock wait.
            if any(reader.stream is stream and reader.thread.is_alive() for reader in readers):
                continue
            if stream and not stream.closed:
                stream.close()
        except BaseException as error:
            record("close_owned_stream", error)
    receipt["passed"] = not failures and receipt["exit_confirmed"]
    if failures:
        if primary:
            primary.add_note("Owned process cleanup failed: " + ", ".join(type(error).__name__ for error in failures))
        else:
            raise ExceptionGroup("Owned process cleanup failed", failures)
    return receipt


def process_metrics(process):
    """Read counters through the exact Popen handle; never enumerate other processes."""
    from ctypes import wintypes

    class MemoryCounters(ctypes.Structure):
        _fields_ = [("cb", wintypes.DWORD), ("faults", wintypes.DWORD)] + [
            (name, ctypes.c_size_t) for name in (
                "peak_working_set", "working_set", "peak_paged", "paged",
                "peak_nonpaged", "nonpaged", "pagefile", "peak_pagefile", "private")]

    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    query_memory = kernel.K32GetProcessMemoryInfo
    query_memory.argtypes = [wintypes.HANDLE, ctypes.POINTER(MemoryCounters), wintypes.DWORD]
    query_memory.restype = wintypes.BOOL
    query_handles = kernel.GetProcessHandleCount
    query_handles.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
    query_handles.restype = wintypes.BOOL
    memory, handles = MemoryCounters(), wintypes.DWORD()
    memory.cb = ctypes.sizeof(memory)
    if not query_memory(int(process._handle), ctypes.byref(memory), memory.cb):
        raise ctypes.WinError(ctypes.get_last_error())
    if not query_handles(int(process._handle), ctypes.byref(handles)):
        raise ctypes.WinError(ctypes.get_last_error())
    return {"handles": handles.value, "privateBytes": getattr(memory, "private"),
            "workingSetBytes": memory.working_set}


def cleanup_case(process, readers, receipt):
    primary = sys.exception()
    if primary is not None:
        receipt["passed"] = False
        receipt["error"] = {"exception_type": type(primary).__name__}
    try:
        cleanup(process, readers, receipt=receipt.setdefault("cleanup", {}))
    except BaseException as error:
        receipt["passed"] = False
        receipt.setdefault("error", {"exception_type": type(error).__name__})
        raise


def protocol_case(exe, idle_seconds, receipt=None):
    receipt = receipt if receipt is not None else {"case": "initialize-list-negative-parity-eof", "passed": False}
    process, out, err = start(exe)
    try:
        count = initialize(process, out)
        sessions = tool(process, out, 3, {"action": "list"})
        require(sessions.get("success") is True and sessions.get("sessions") == [], "Fresh host sessions are not empty")
        with tempfile.TemporaryDirectory(prefix="excelmcp-negative-") as folder:
            # Missing files fail before the Service's temporary native workbook-open path.
            # The private worker permits 15s startup, plus snapshot and release waits.
            negative = tool(process, out, 4, {"action": "test", "file_path": str(Path(folder) / "absent.xlsx")},
                            expected_error=True, timeout_seconds=30)
            require(negative.get("success") is False and negative.get("exists") is False,
                    "Missing-file diagnostic unexpectedly succeeded")
        samples = [process_metrics(process)]
        if idle_seconds:
            deadline = time.monotonic() + idle_seconds
            while time.monotonic() < deadline:
                require(process.poll() is None, "Host exited during idle")
                time.sleep(min(0.25, max(0, deadline - time.monotonic())))
                samples.append(process_metrics(process))
        process.stdin.close()
        require(process.wait(timeout=8) == 0, "EOF shutdown returned a failure")
        receipt.update({"case": "initialize-list-negative-parity-eof", "passed": True, "toolCount": count, "idleSeconds": idle_seconds,
                "idleResources": {"scope": "Owned host after protocol warmup; observational, not an all-day leak guarantee",
                                  "growthAssertions": False, "sampleCount": len(samples),
                                  "baseline": samples[0], "final": samples[-1],
                                  "peak": {key: max(sample[key] for sample in samples) for key in samples[0]},
                                  "delta": {key: samples[-1][key] - samples[0][key] for key in samples[0]}}})
        return receipt
    finally:
        cleanup_case(process, [out, err], receipt)


def stdout_case(exe, receipt=None):
    receipt = receipt if receipt is not None else {"case": "broken-stdout-stdin-held-open", "passed": False}
    process, out, err = start(exe, stop_id=2)
    try:
        initialize(process, out)
        out.thread.join(timeout=2)
        require(not out.thread.is_alive(), "Reader did not release stdout before close")
        process.stdout.close()
        send(process, 3, "tools/list", {})
        # The stdin writer remains open throughout this exit check.
        code = process.wait(timeout=8)
        require(code == 0, "Broken-stdout shutdown returned a failure")
        receipt.update(case="broken-stdout-stdin-held-open", passed=True, exitCode=code)
        return receipt
    finally:
        cleanup_case(process, [out, err], receipt)


PARENT_HELPER = r'''
import ctypes,json,os,pathlib,subprocess,sys,time
k=ctypes.WinDLL('kernel32',use_last_error=True)
k.OpenProcess.argtypes=[ctypes.c_uint,ctypes.c_int,ctypes.c_uint];k.OpenProcess.restype=ctypes.c_void_p
k.GetCurrentProcess.restype=ctypes.c_void_p
k.DuplicateHandle.argtypes=[ctypes.c_void_p,ctypes.c_void_p,ctypes.c_void_p,ctypes.POINTER(ctypes.c_void_p),ctypes.c_uint,ctypes.c_int,ctypes.c_uint]
k.CloseHandle.argtypes=[ctypes.c_void_p]
owner=k.OpenProcess(0x40,False,int(sys.argv[3]))
if not owner:raise ctypes.WinError(ctypes.get_last_error())
with open(sys.argv[4],'wb') as out,open(sys.argv[5],'wb') as err:
 p=subprocess.Popen([sys.argv[1]],stdin=sys.stdin.buffer,stdout=out,stderr=err,env=dict(os.environ,EXCELMCP_TELEMETRY_MODE='off'),creationflags=subprocess.CREATE_NO_WINDOW)
 handed=False
 try:
  remote=ctypes.c_void_p()
  if not k.DuplicateHandle(k.GetCurrentProcess(),int(p._handle),owner,ctypes.byref(remote),0,False,2):raise ctypes.WinError(ctypes.get_last_error())
  stage=pathlib.Path(sys.argv[2]+'.stage');stage.write_text(json.dumps({'handle':remote.value}));stage.replace(sys.argv[2])
  deadline=time.monotonic()+8
  while not pathlib.Path(sys.argv[6]).exists() and time.monotonic()<deadline:time.sleep(.02)
  handed=pathlib.Path(sys.argv[6]).exists()
  if not handed:raise RuntimeError('Owner did not acknowledge process handle')
 finally:
  if not handed and p.poll() is None:p.kill();p.wait(timeout=3)
  k.CloseHandle(owner)
'''


def parent_case(exe, receipt=None):
    receipt = receipt if receipt is not None else {"case": "parent-exit-stdin-held-open", "passed": False}
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.WaitForSingleObject.argtypes = [ctypes.c_void_p, ctypes.c_uint]
    kernel.TerminateProcess.argtypes = [ctypes.c_void_p, ctypes.c_uint]
    kernel.CloseHandle.argtypes = [ctypes.c_void_p]
    with tempfile.TemporaryDirectory(prefix="excelmcp-parent-") as folder:
        root = Path(folder)
        record, ack = root / "owned.json", root / "ack"
        read_fd, write_fd = os.pipe()
        helper = None
        owned = None
        try:
            helper = subprocess.Popen(
                [sys.executable, "-c", PARENT_HELPER, str(exe), str(record), str(os.getpid()), str(root / "stdout"), str(root / "stderr"), str(ack)],
                stdin=read_fd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW,
            )
            os.close(read_fd)
            read_fd = None
            deadline = time.monotonic() + 5
            while not record.exists() and time.monotonic() < deadline:
                if helper.poll() is not None:
                    break
                time.sleep(0.02)
            owned = int(json.loads(record.read_text())["handle"])
            ack.write_text("owned")
            require(helper.wait(timeout=5) == 0, "Parent helper failed")
            require(kernel.WaitForSingleObject(owned, 8000) == 0, "Parent-exit shutdown deadline expired")
            code = ctypes.c_ulong()
            kernel.GetExitCodeProcess.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_ulong)]
            require(kernel.GetExitCodeProcess(owned, ctypes.byref(code)), "Cannot read child exit status")
            require(code.value == 0, "Parent-exit shutdown returned a failure")
            receipt.update(case="parent-exit-stdin-held-open", passed=True, exitCode=code.value)
            return receipt
        finally:
            primary = sys.exception()
            if primary is not None:
                receipt["passed"] = False
                receipt["error"] = {"exception_type": type(primary).__name__}
            finalizer = receipt.setdefault("cleanup", {})
            finalizer.update(helper_process_id=helper.pid if helper is not None else None,
                             owned_child_handle_received=owned is not None, exit_confirmed=False, forced_termination=False)
            failures = []
            if helper is not None:
                try:
                    if helper.poll() is None:
                        helper.kill()
                    helper.wait(timeout=3)
                except BaseException as error:
                    failures.append(error)
            if owned is None and record.exists():
                try:
                    owned = int(json.loads(record.read_text())["handle"])
                except BaseException as error:
                    failures.append(error)
            if owned:
                try:
                    state = kernel.WaitForSingleObject(owned, 0)
                    require(state in (0, 258), "Owned process wait failed")
                    if state == 258:
                        finalizer["forced_termination"] = True
                        require(kernel.TerminateProcess(owned, 1), "Owned process termination failed")
                        require(kernel.WaitForSingleObject(owned, 3000) == 0, "Owned process cleanup deadline expired")
                    finalizer["exit_confirmed"] = True
                except BaseException as error:
                    failures.append(error)
                finally:
                    if not kernel.CloseHandle(owned):
                        failures.append(RuntimeError("Owned process handle close failed"))
            for descriptor in (write_fd, read_fd):
                if descriptor is not None:
                    try:
                        os.close(descriptor)
                    except BaseException as error:
                        failures.append(error)
            finalizer["owned_child_handle_received"] = owned is not None
            finalizer["failures"] = [{"exception_type": type(error).__name__} for error in failures]
            finalizer["passed"] = not failures and finalizer["exit_confirmed"]
            if failures:
                receipt["passed"] = False
                receipt.setdefault("error", {"exception_type": type(primary).__name__ if primary else "ExceptionGroup"})
                if primary:
                    primary.add_note("Parent-case cleanup failed: " + ", ".join(type(error).__name__ for error in failures))
                else:
                    raise ExceptionGroup("Parent-case cleanup failed", failures)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--exe", required=True, type=Path)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--idle-seconds", type=int, default=0, choices=range(61), metavar="0..60")
    args = parser.parse_args()
    require(sys.version_info >= (3, 11), "Python 3.11 or later is required")
    require(os.name == "nt", "This smoke requires Windows")
    exe = args.exe.resolve(strict=True)
    digest = hashlib.sha256(exe.read_bytes()).hexdigest()
    report = {"sha256": digest, "allPassed": False, "cases": []}
    try:
        for name, run, extra in (
            ("initialize-list-negative-parity-eof", protocol_case, (args.idle_seconds,)),
            ("broken-stdout-stdin-held-open", stdout_case, ()),
            ("parent-exit-stdin-held-open", parent_case, ())):
            case = {"case": name, "passed": False}
            report["cases"].append(case)
            try:
                run(exe, *extra, receipt=case)
            except BaseException as error:
                case["passed"] = False
                case.setdefault("error", {"exception_type": type(error).__name__})
                raise
        require(hashlib.sha256(exe.read_bytes()).hexdigest() == digest, "Executable changed during verification")
        report["allPassed"] = True
    except BaseException as error:
        report["error"] = {"exception_type": type(error).__name__}
        raise
    finally:
        primary = sys.exception()
        try:
            if args.report:
                args.report.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
            print(json.dumps(report, indent=2))
        except BaseException as error:
            if primary:
                primary.add_note("Report output failed: " + type(error).__name__)
            else:
                raise


if __name__ == "__main__":
    main()
