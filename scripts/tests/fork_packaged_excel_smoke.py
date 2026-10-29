"""Round-trip disposable XLSM fixtures through packaged MCP and CLI entry points.

Uses Python's standard library and the reviewed sibling protocol/benchmark
helpers. Never discovers or attaches to preexisting Excel workbooks. No native
work occurs with --self-test. Reports contain hashes and verification outcomes,
not paths, workbook values, tool payloads, or captured diagnostic text.
"""

from __future__ import annotations

import argparse
import ctypes
from ctypes import wintypes
import importlib.util
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
import unittest
import uuid

sys.dont_write_bytecode = True


class SmokeError(RuntimeError):
    pass


def require(condition, message):
    if not condition:
        raise SmokeError(message)


def helper(name, path):
    require(path.is_file(), "A required reviewed smoke helper is missing.")
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def tool_payload(result, expected_error=False):
    text = [item.get("text") for item in result.get("content", []) if item.get("type") == "text"]
    require(len(text) == 1 and isinstance(text[0], str), "MCP result did not contain one JSON text object.")
    try:
        value = json.loads(text[0])
    except ValueError as error:
        raise SmokeError("MCP result text was not JSON.") from error
    require(isinstance(value, dict) and value == result.get("structuredContent"), "MCP text and structured content differed.")
    require((result.get("isError") is True) == expected_error, "MCP error classification differed from the expected outcome.")
    require(value.get("success") is (not expected_error), "MCP business outcome differed from the expected outcome.")
    return value


def owned_fixture_sessions(entries, path, id_key):
    expected = os.path.normcase(os.path.abspath(path))
    selected = []
    for entry in entries:
        location = entry.get("filePath")
        if not isinstance(location, str) or os.path.normcase(os.path.abspath(location)) != expected:
            continue
        require(entry.get("isAttached") is False, "Fixture cleanup refused an attached workbook.")
        identifier = entry.get(id_key)
        require(isinstance(identifier, str) and identifier, "Owned fixture session lacks its public identifier.")
        selected.append((identifier, entry.get("canClose") is True))
    return selected


DATA = [["Item", "Count", "Approved"], ["synthetic", 7, True], ["other", 11, False]]


def verify_page(page):
    require(page.get("rangeAddress") == "$A$1:$C$3" and page.get("scopeAddress") == "$A$1:$C$3", "Bounded page scope changed.")
    require((page.get("rowCount"), page.get("columnCount")) == (3, 3), "Bounded page dimensions changed.")
    require(page.get("hasMore") is False, "Complete bounded page reported an incomplete read.")
    require(page.get("nextRowOffset") is None and page.get("nextColumnOffset") is None, "Complete page retained continuation offsets.")
    values = page.get("values")
    require(isinstance(values, list) and len(values) == 3, "Round-trip returned an invalid row count.")
    for actual, expected in zip(values, DATA):
        require(isinstance(actual, list) and len(actual) == 3, "Round-trip returned an invalid column count.")
        for value, target in zip(actual, expected):
            if type(target) is int:
                require(type(value) in (int, float) and value == target, "Round-trip changed a numeric cell.")
            else:
                require(type(value) is type(target) and value == target, "Round-trip changed a text or Boolean cell.")


def verify_bounds_and_format(bounds, summary):
    require(bounds.get("rangeAddress") == "$A$1:$C$3", "Used bounds changed.")
    require((bounds.get("rowCount"), bounds.get("columnCount"), bounds.get("cellCount")) == (3, 3, 9), "Used bounds dimensions changed.")
    require(summary.get("cellCount") == 9 and summary.get("rangeAddress") == "$A$1:$C$3", "Format summary scope changed.")
    stored = summary.get("stored")
    require(isinstance(stored, dict) and isinstance(stored.get("font"), dict), "Native font inspection was missing.")
    require(type(stored["font"].get("bold")) is bool and len(stored.get("borders", [])) == 8, "Native format inspection was incomplete.")


def process_identity(value):
    require(isinstance(value, dict), "Native process identity is missing.")
    pid, created = value.get("processId"), value.get("startedAtUtcFileTime")
    require(type(pid) is int and pid > 0 and type(created) is int and created > 0, "Native process identity is invalid.")
    return (pid, created)


def selected_owned_records(records, controller):
    expected = (controller["process_id"], controller["created_at_utc_filetime"])
    selected = []
    for record in records:
        if not isinstance(record, dict) or record.get("controller") != {"processId": expected[0], "startedAtUtcFileTime": expected[1]}:
            continue
        process_identity(record["controller"])
        worker = process_identity(record.get("worker"))
        excel = process_identity(record["excel"]) if record.get("excel") is not None else None
        require(excel is None or record.get("ownsExcel") is True, "Smoke refused an attached native process.")
        selected.append((worker, excel))
    return selected


class NativeProcessHandle:
    """Query/synchronize access only; PID reuse never transfers ownership."""

    def __init__(self, identity):
        require(os.name == "nt", "Native process proof requires Windows.")
        self.kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        self.kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        self.kernel.OpenProcess.restype = wintypes.HANDLE
        self.kernel.GetProcessTimes.argtypes = [wintypes.HANDLE] + [ctypes.POINTER(wintypes.FILETIME)] * 4
        self.kernel.GetProcessTimes.restype = wintypes.BOOL
        self.kernel.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
        self.kernel.WaitForSingleObject.restype = wintypes.DWORD
        self.kernel.CloseHandle.argtypes = [wintypes.HANDLE]
        self.kernel.CloseHandle.restype = wintypes.BOOL
        self.kernel.ProcessIdToSessionId.argtypes = [wintypes.DWORD, ctypes.POINTER(wintypes.DWORD)]
        self.kernel.ProcessIdToSessionId.restype = wintypes.BOOL
        self.handle = self.kernel.OpenProcess(0x100000 | 0x1000, False, identity[0])
        require(bool(self.handle), "An owned native process could not be opened for exit proof.")
        try:
            times = [wintypes.FILETIME() for _ in range(4)]
            require(self.kernel.GetProcessTimes(self.handle, *(ctypes.byref(value) for value in times)), "Native process creation time could not be read.")
            created = (times[0].dwHighDateTime << 32) | times[0].dwLowDateTime
            require(created == identity[1], "Native process identity changed; PID reuse was refused.")
            own_session, target_session = wintypes.DWORD(), wintypes.DWORD()
            require(self.kernel.ProcessIdToSessionId(os.getpid(), ctypes.byref(own_session))
                    and self.kernel.ProcessIdToSessionId(identity[0], ctypes.byref(target_session))
                    and own_session.value == target_session.value, "Native process session did not match the smoke owner.")
        except BaseException:
            self.close()
            raise

    def wait(self, seconds):
        result = self.kernel.WaitForSingleObject(self.handle, max(0, min(0xfffffffe, int(seconds * 1000))))
        require(result in (0, 258), "Native process exit query failed.")
        return result == 0

    def close(self):
        if self.handle:
            handle, self.handle = self.handle, None
            require(self.kernel.CloseHandle(handle), "Native process proof handle could not be released.")


class NativeOwners:
    def __init__(self, controller, report, registry=None, opener=NativeProcessHandle):
        self.controller, self.report, self.opener = controller, report, opener
        self.registry = registry if registry is not None else Path(os.environ["LOCALAPPDATA"]) / "ExcelMcp" / "workers"
        self.handles = {}
        self.entries = report.setdefault("native_processes", [])
        self.failures = report.setdefault("native_ownership_failures", [])

    def records(self):
        if not self.registry.exists():
            return []
        paths = list(self.registry.glob("*.json"))
        require(len(paths) <= 4096, "Native registry inventory exceeded its diagnostic limit.")
        records = []
        for path in paths:
            try:
                require(not path.is_symlink(), "Native registry link was refused.")
                with path.open("rb") as stream:
                    data = stream.read(65537)
                require(len(data) <= 65536, "Native registry record exceeded its diagnostic limit.")
                records.append(json.loads(data.decode("utf-8-sig")))
            except FileNotFoundError:
                continue
        return records

    def capture(self, required=False):
        matched = selected_owned_records(self.records(), self.controller)
        observed_excel = False
        for worker, excel in matched:
            for role, identity in (("worker", worker), ("excel", excel)):
                if identity is None:
                    continue
                observed_excel |= role == "excel"
                key = (role,) + identity
                if key in self.handles:
                    continue
                entry = {"role": role, "process_id": identity[0], "created_at_utc_filetime": identity[1], "exit_confirmed": False}
                self.entries.append(entry)
                try:
                    self.handles[key] = (self.opener(identity), entry)
                except Exception as error:
                    self.failures.append({"phase": "capture", "exception_type": type(error).__name__, "process_id": identity[0], "created_at_utc_filetime": identity[1]})
                    raise
        require(not required or observed_excel, "Owned Excel identity was not observed before fixture work.")

    def finish(self, seconds=15):
        try:
            self.capture()
        except Exception as error:
            self.failures.append({"phase": "final_capture", "exception_type": type(error).__name__})
        deadline = time.monotonic() + seconds
        for handle, entry in self.handles.values():
            try:
                entry["exit_confirmed"] = handle.wait(max(0, deadline - time.monotonic()))
                if not entry["exit_confirmed"]:
                    self.failures.append({"phase": "wait", "process_id": entry["process_id"], "created_at_utc_filetime": entry["created_at_utc_filetime"]})
            except Exception as error:
                self.failures.append({"phase": "wait", "exception_type": type(error).__name__, "process_id": entry["process_id"]})
            finally:
                try:
                    handle.close()
                except Exception as error:
                    self.failures.append({"phase": "close_handle", "exception_type": type(error).__name__, "process_id": entry["process_id"]})
        self.handles.clear()
        confirmed = any(entry["role"] == "excel" for entry in self.entries) and bool(self.entries) and all(entry["exit_confirmed"] for entry in self.entries) and not self.failures
        self.report["native_exit_confirmed"] = confirmed
        require(confirmed, "Owned worker/Excel exit was not confirmed; exact identities and fixture were retained in the receipt.")


def retained_cli_state(item, report, phase):
    data = item.get("data", {})
    retained = data.get("retainedExcelProcesses")
    state = {"phase": phase, "processes": [], "retained_error_present": bool(data.get("retainedExcelError")),
             "shutdown_warning_present": bool(data.get("shutdownWarning")), "response_success": item.get("success") is True}
    report.setdefault("cli_native_retention", []).append(state)
    require(isinstance(retained, list), "CLI native retention state was missing.")
    for entry in retained:
        require(isinstance(entry, dict) and type(entry.get("processId")) is int and entry["processId"] > 0
                and isinstance(entry.get("startTime"), str), "CLI retained native identity was invalid.")
        state["processes"].append({"process_id": entry["processId"], "start_time": entry["startTime"]})
    require(type(data.get("retainedExcelProcessCount")) is int and data["retainedExcelProcessCount"] == len(retained) and not retained
            and not data.get("retainedExcelError") and not data.get("shutdownWarning") and item.get("success") is True,
            "CLI retained native state prevents smoke acceptance.")


class McpHost:
    id_key = "workbook_session_id"

    def __init__(self, exe, protocol, report):
        self.exe, self.protocol, self.report = exe, protocol, report
        self.process = None
        self.readers = []
        self.next_id = 10
        self.schemas = {}

    def response(self, request_id):
        deadline = time.monotonic() + 180
        while True:
            try:
                return self.readers[0].response(request_id)
            except TimeoutError:
                # Keep waiting for the same request; never send a mutation again.
                require(time.monotonic() < deadline, "Native MCP response deadline expired.")

    def start(self):
        self.process, out, err = self.protocol.start(self.exe)
        self.readers = [out, err]
        self.report["tool_count"] = self.protocol.initialize(self.process, out)
        self.protocol.send(self.process, 3, "tools/list", {})
        self.schemas = {tool["name"]: tool["inputSchema"] for tool in self.response(3)["tools"]}
        required = {"file": {"read_only", "workbook_session_id", "file_path"},
                    "file_read": {"action"}, "range": {"values", "range_address"},
                    "range_read": {"page_rows", "page_columns"}, "range_format_read": {"view"},
                    "workbook_read": {"action"}, "vba": {"arguments", "expected_source_hash"},
                    "vba_read": {"module_name"}}
        for name, fields in required.items():
            require(name in self.schemas and fields <= self.schemas[name].get("properties", {}).keys(), "Packaged MCP schema lacks a required smoke capability.")
        require(self.sessions() == [], "Fresh smoke MCP host already owns workbook sessions.")

    def call(self, name, arguments, expected_error=False):
        self.next_id += 1
        begin = time.perf_counter()
        self.protocol.send(self.process, self.next_id, "tools/call", {"name": name, "arguments": arguments})
        value = tool_payload(self.response(self.next_id), expected_error)
        self.report.setdefault("calls", []).append({"tool": name, "action": arguments["action"],
                                                   "wall_seconds": time.perf_counter() - begin, "expected_error": expected_error})
        return value

    def sessions(self):
        return self.call("file_read", {"action": "list"})["sessions"]

    def close(self, identifier):
        self.call("file", {"action": "close", "workbook_session_id": identifier, "save": False})

    def shutdown(self):
        if self.process is None:
            return
        try:
            if self.process.poll() is None:
                self.process.stdin.close()
            require(self.process.wait(timeout=30) == 0, "Packaged MCP host did not exit successfully.")
            self.report["host_exit_confirmed"] = True
        finally:
            primary = sys.exception()
            try:
                self.protocol.cleanup(self.process, self.readers, receipt=self.report.setdefault("host_cleanup", {}))
            except Exception as error:
                self.report.setdefault("host_cleanup_failures", []).append({"phase": "protocol_cleanup", "exception_type": type(error).__name__})
                self.report["passed"] = False
                if primary is None:
                    raise SmokeError("Packaged MCP host cleanup was not confirmed; see receipt.") from None


class CliHost:
    id_key = "sessionId"

    def __init__(self, exe, benchmark, report):
        self.exe, self.benchmark, self.report = exe, benchmark, report
        self.pipe = "excelmcp-packaged-smoke-" + uuid.uuid4().hex
        self.environment = dict(os.environ, EXCELMCP_CLI_PIPE=self.pipe, EXCELMCP_TELEMETRY_MODE="off")
        self.runner = benchmark.Runner(exe, self.environment)
        self.process = None
        self.stdout = self.stderr = None

    def start(self):
        self.benchmark.validate_status(self.runner.call(["service", "status"], "preflight", "service.status")[0], False, 0)
        self.stdout, self.stderr = tempfile.TemporaryFile(), tempfile.TemporaryFile()
        self.process = subprocess.Popen([str(self.exe), "--quiet", "service", "run", "--pipe-name", self.pipe],
                                        stdin=subprocess.DEVNULL, stdout=self.stdout, stderr=self.stderr,
                                        env=self.environment, creationflags=subprocess.CREATE_NO_WINDOW)
        self.report["owned_host"] = self.benchmark.owned_process_identity(self.process)
        deadline = time.monotonic() + 20
        while True:
            require(self.process.poll() is None, "Private packaged CLI daemon exited during startup.")
            status = self.runner.call(["service", "status"], "readiness", "service.status", require_success=False)[0]
            if status["success"] and status.get("data", {}).get("running") is True:
                self.benchmark.validate_status(status, True, self.process.pid)
                break
            require(time.monotonic() < deadline, "Private packaged CLI daemon was not ready.")
            time.sleep(0.1)
        require(self.sessions() == [], "Fresh private CLI daemon already owns sessions.")

    def direct(self, arguments, command):
        item = self.runner.call(arguments, command, command, timeout=180)[0]
        return item.get("data", {})

    def call(self, command, arguments=None, identifier=None, expected_error=False):
        request = {"command": command}
        if arguments is not None:
            request["args"] = arguments
        if identifier is not None:
            request["sessionId"] = identifier
        item = self.runner.call(["batch", "--input", "-", "--stop-on-error"], command, command,
                                input_text=json.dumps([request]), indexed=True,
                                require_success=not expected_error, timeout=180)[0]
        require(item["success"] is (not expected_error), "Packaged CLI outcome differed from the expected outcome.")
        return item if expected_error else item.get("data", {})

    def sessions(self):
        return self.direct(["session", "list"], "session.list")["sessions"]

    def close(self, identifier):
        self.call("session.close", {"save": False}, identifier)

    def shutdown(self):
        try:
            if self.process is not None:
                if self.process.poll() is None:
                    stopped = self.runner.call(["service", "stop"], "cleanup.stop", "service.stop", timeout=45, require_success=False)[0]
                    retained_cli_state(stopped, self.report, "stop")
                    require(stopped["success"] is True, "CLI shutdown was not accepted.")
                require(self.process.wait(timeout=15) == 0, "Private CLI daemon did not exit successfully.")
                status = self.runner.call(["service", "status"], "cleanup.status", "service.status")[0]
                retained_cli_state(status, self.report, "status")
                self.benchmark.validate_status(status, False, 0)
                self.report["host_exit_confirmed"] = True
        finally:
            primary = sys.exception()
            cleanup = self.report.setdefault("host_cleanup", {})
            failure = self.benchmark.finish_owned_child(self.process, cleanup, primary)
            if cleanup.get("forced_termination"):
                self.report["forced_owned_host_termination"] = True
            try:
                final_status = self.runner.call(["service", "status"], "cleanup.finalstatus", "service.status", timeout=10, require_success=False)[0]
                retained_cli_state(final_status, self.report, "final_status")
                self.benchmark.validate_status(final_status, False, 0)
            except Exception as error:
                cleanup.setdefault("failures", []).append({"phase": "final_status", "exception_type": type(error).__name__})
                if failure is None:
                    failure = SmokeError("CLI native retention after shutdown was not confirmed; see receipt.")
            for label, stream in (("stdout", self.stdout), ("stderr", self.stderr)):
                if stream is not None:
                    try:
                        stream.close()
                    except Exception as error:
                        cleanup.setdefault("failures", []).append({"phase": f"close_{label}", "exception_type": type(error).__name__})
                        if failure is None:
                            failure = SmokeError("Packaged CLI log cleanup was not confirmed; see receipt.")
            self.report["subprocesses"] = self.runner.records
            if failure is not None:
                self.report["passed"] = False
                if primary is None:
                    raise SmokeError("Packaged CLI host cleanup was not confirmed; see receipt.") from None


def cleanup_sessions(host, path, report):
    deadline = time.monotonic() + 90
    while True:
        selected = owned_fixture_sessions(host.sessions(), path, host.id_key)
        if not selected:
            report["fixture_sessions_closed"] = True
            return
        for identifier, can_close in selected:
            if can_close:
                host.close(identifier)
        require(time.monotonic() < deadline, "Owned fixture sessions could not be closed; native cleanup remains unconfirmed.")
        time.sleep(0.2)


def mcp_case(host, path, identities):
    created = host.call("file", {"action": "create", "file_path": str(path)})
    identifier = created[host.id_key]
    host.native.capture(required=True)
    host.call("range", {"action": "set-values", "workbook_session_id": identifier,
                        "sheet_name": "Sheet1", "range_address": "A1:C3", "values": DATA})
    code = 'Public Function PackagedValue(ByVal token As Variant, ByVal amount As Variant) As String\nPackagedValue = CStr(token) & ":" & CStr(amount)\nEnd Function'
    host.call("vba", {"action": "import", "workbook_session_id": identifier, "module_name": "PackagedSmoke", "vba_code": code})
    source = host.call("vba_read", {"action": "view", "workbook_session_id": identifier, "module_name": "PackagedSmoke"})
    replacement = code.replace("CStr(token)", '"guarded:" & CStr(token)')
    host.call("vba", {"action": "replace-module", "workbook_session_id": identifier,
                      "module_name": "PackagedSmoke", "expected_source_hash": source["sourceHash"], "vba_code": replacement})
    invoked = host.call("vba", {"action": "invoke", "workbook_session_id": identifier,
                               "procedure_name": "PackagedSmoke.PackagedValue",
                               "arguments": [{"type": "string", "value": "synthetic"}, {"type": "number", "value": 7}]})
    require(invoked.get("returnKind") == "string" and invoked.get("value") == "guarded:synthetic:7", "Packaged typed VBA result changed.")
    host.report["guarded_vba_and_typed_invoke"] = True
    host.call("file", {"action": "save", "workbook_session_id": identifier})
    host.close(identifier)
    before = identities.file_identity(path)["sha256"]
    audit = host.call("file", {"action": "open", "file_path": str(path), "read_only": True})
    identifier = audit[host.id_key]
    host.native.capture(required=True)
    require(audit.get("auditOnly") is True, "Packaged MCP did not open an audit session.")
    info = host.call("workbook_read", {"action": "get-info", "workbook_session_id": identifier})
    require(info.get("readOnly") is True, "Native MCP workbook was not read-only.")
    common = {"workbook_session_id": identifier, "sheet_name": "Sheet1"}
    page_args = dict(common, action="get-values-page", range_address="A1:C3", page_rows=3, page_columns=3)
    verify_page(host.call("range_read", page_args))
    bounds = host.call("range_read", dict(common, action="get-used-range-info"))
    summary = host.call("range_format_read", dict(common, action="get-format-summary", range_address="A1:C3", view="stored"))
    verify_bounds_and_format(bounds, summary)
    refused = host.call("range", dict(common, action="set-values", range_address="A2", values=[["refused"]], overwrite_policy="allow"), True)
    require(refused.get("errorCategory") == "Permissions", "Audit mutation was not refused by the MCP permission guard.")
    verify_page(host.call("range_read", page_args))
    host.close(identifier)
    require(identities.file_identity(path)["sha256"] == before, "MCP audit altered the saved package.")
    host.report.update(roundtrip_passed=True, audit_hash_unchanged=True, saved_fixture_sha256=before)


def cli_case(host, path, identities):
    identifier = host.direct(["session", "create", str(path)], "session.create")["sessionId"]
    host.native.capture(required=True)
    host.call("range.set-values", {"sheetName": "Sheet1", "rangeAddress": "A1:C3", "values": DATA}, identifier)
    host.call("session.save", identifier=identifier)
    host.close(identifier)
    before = identities.file_identity(path)["sha256"]
    audit = host.direct(["session", "open", str(path), "--read-only"], "session.open")
    identifier = audit["sessionId"]
    host.native.capture(required=True)
    require(audit.get("auditOnly") is True, "Packaged CLI did not open an audit session.")
    require(host.call("workbook.get-info", identifier=identifier).get("readOnly") is True, "Native CLI workbook was not read-only.")
    args = {"sheetName": "Sheet1", "rangeAddress": "A1:C3", "pageRows": 3, "pageColumns": 3}
    verify_page(host.call("range.get-values-page", args, identifier))
    bounds = host.call("range.get-used-range-info", {"sheetName": "Sheet1"}, identifier)
    summary = host.call("rangeformat.get-format-summary", {"sheetName": "Sheet1", "rangeAddress": "A1:C3", "view": "stored"}, identifier)
    verify_bounds_and_format(bounds, summary)
    refused = host.call("range.set-values", {"sheetName": "Sheet1", "rangeAddress": "A2", "values": [["refused"]], "overwritePolicy": "allow"}, identifier, True)
    require(refused.get("error", {}).get("category") == "Permissions", "Audit mutation was not refused by the CLI permission guard.")
    verify_page(host.call("range.get-values-page", args, identifier))
    host.close(identifier)
    require(identities.file_identity(path)["sha256"] == before, "CLI audit altered the saved package.")
    host.report.update(roundtrip_passed=True, audit_hash_unchanged=True, saved_fixture_sha256=before)


def cleanup_fixtures(root, report, remove=shutil.rmtree):
    if root is None:
        return
    if report["success"]:
        try:
            require(root.resolve().parent == Path(tempfile.gettempdir()).resolve()
                    and root.name.startswith("excelmcp-packaged-native-") and not root.is_symlink(), "Disposable fixture cleanup scope changed.")
            remove(root)
            report["disposable_fixtures_removed"] = True
            return
        except Exception as error:
            report["success"] = False
            report["fixture_cleanup_failure_type"] = type(error).__name__
            report.setdefault("error", "Disposable fixture cleanup was not confirmed; see receipt.")
    report["retained_fixture_directory_name"] = root.name


class ParserTests(unittest.TestCase):
    def test_mcp_result_requires_matching_structured_and_text_outcomes(self):
        value = {"success": True, "values": [[False, 0]]}
        result = {"content": [{"type": "text", "text": json.dumps(value)}], "structuredContent": value}
        self.assertEqual(value, tool_payload(result))
        with self.assertRaises(SmokeError):
            tool_payload(dict(result, structuredContent={"success": True}))
        with self.assertRaises(SmokeError):
            tool_payload(dict(result, isError=True))

    def test_page_rejects_stringified_numbers_and_booleans(self):
        page = {"rangeAddress": "$A$1:$C$3", "scopeAddress": "$A$1:$C$3", "rowCount": 3,
                "columnCount": 3, "hasMore": False, "values": DATA}
        verify_page(page)
        for column, replacement in ((1, "7"), (2, "True")):
            altered = json.loads(json.dumps(page))
            altered["values"][1][column] = replacement
            with self.assertRaises(SmokeError):
                verify_page(altered)

    def test_cli_cleanup_does_not_replace_stop_failure(self):
        benchmark = helper("benchmark_cleanup_test", Path(__file__).with_name("fork_cli_batch_benchmark.py"))
        primary = SmokeError("Original stop failed.")

        class Child:
            pid = 321

            def __init__(self, denied):
                self.denied = denied

            def poll(self):
                return None

            def kill(self):
                if self.denied:
                    raise PermissionError("private-marker")

            def wait(self, timeout):
                raise subprocess.TimeoutExpired("private-marker", timeout)

        class Runner:
            records = []

            def call(self, *args, **kwargs):
                raise primary

        class Stream:
            def close(self):
                raise OSError("private-marker")

        for denied in (True, False):
            with self.subTest(kill_denied=denied):
                host = CliHost.__new__(CliHost)
                host.process, host.benchmark, host.runner = Child(denied), benchmark, Runner()
                host.stdout = host.stderr = Stream()
                host.report = {"passed": False, "owned_host": {"pid": 321, "creation_filetime": 456}}
                with self.assertRaises(SmokeError) as caught:
                    host.shutdown()
                self.assertIs(primary, caught.exception)
                self.assertFalse(host.report["passed"])
                self.assertEqual({"pid": 321, "creation_filetime": 456}, host.report["owned_host"])
                self.assertFalse(host.report["host_cleanup"]["no_lingering_owned_daemon"])
                self.assertNotIn("private-marker", json.dumps(host.report))

    def test_mcp_cleanup_does_not_replace_exit_failure(self):
        primary = subprocess.TimeoutExpired("private-marker", 30)

        class Child:
            def poll(self):
                return 0

            def wait(self, timeout):
                raise primary

        class Protocol:
            def cleanup(self, *args, **kwargs):
                raise PermissionError("private-marker")

        host = McpHost(Path("synthetic.exe"), Protocol(), {"passed": False})
        host.process = Child()
        with self.assertRaises(subprocess.TimeoutExpired) as caught:
            host.shutdown()
        self.assertIs(primary, caught.exception)
        self.assertEqual("PermissionError", host.report["host_cleanup_failures"][0]["exception_type"])
        self.assertNotIn("private-marker", json.dumps(host.report))

    def test_fixture_cleanup_failure_keeps_receipt_and_primary(self):
        path = Path(tempfile.gettempdir()) / "excelmcp-packaged-native-synthetic"

        def denied(root):
            raise PermissionError("private-marker")

        report = {"success": True, "error": "Original failure."}
        cleanup_fixtures(path, report, denied)
        self.assertFalse(report["success"])
        self.assertEqual("Original failure.", report["error"])
        self.assertEqual("PermissionError", report["fixture_cleanup_failure_type"])
        self.assertEqual(path.name, report["retained_fixture_directory_name"])
        self.assertNotIn("private-marker", json.dumps(report))

    def test_native_registry_requires_exact_controller_and_rejects_attached(self):
        controller = {"process_id": 321, "created_at_utc_filetime": 456}
        record = {"controller": {"processId": 321, "startedAtUtcFileTime": 456},
                  "worker": {"processId": 322, "startedAtUtcFileTime": 457},
                  "excel": {"processId": 323, "startedAtUtcFileTime": 458}, "ownsExcel": True}
        self.assertEqual([((322, 457), (323, 458))], selected_owned_records([record], controller))
        for changed in ({"processId": 999, "startedAtUtcFileTime": 456}, {"processId": 321, "startedAtUtcFileTime": 999}):
            self.assertEqual([], selected_owned_records([dict(record, controller=changed)], controller))
        with self.assertRaises(SmokeError):
            selected_owned_records([dict(record, ownsExcel=False)], controller)

    def test_native_owner_requires_observation_and_retains_unresolved_identity(self):
        class Handle:
            def __init__(self, identity):
                self.identity, self.closed = identity, False

            def wait(self, seconds):
                return self.identity[0] == 322

            def close(self):
                self.closed = True

        report = {}
        owner = NativeOwners({"process_id": 321, "created_at_utc_filetime": 456}, report, registry=Path("unused"), opener=Handle)
        owner.records = lambda: []
        with self.assertRaises(SmokeError):
            owner.capture(required=True)
        with self.assertRaises(SmokeError):
            owner.finish(seconds=0)
        self.assertFalse(report["native_exit_confirmed"])
        owner.records = lambda: [{"controller": {"processId": 321, "startedAtUtcFileTime": 456},
            "worker": {"processId": 322, "startedAtUtcFileTime": 457},
            "excel": {"processId": 323, "startedAtUtcFileTime": 458}, "ownsExcel": True}]
        owner.capture(required=True)
        handles = [handle for handle, _ in owner.handles.values()]
        owner.records = lambda: []  # Disappearing registry record does not discard retained handles.
        with self.assertRaises(SmokeError):
            owner.finish(seconds=0)
        self.assertTrue(all(handle.closed for handle in handles))
        self.assertEqual([322, 323], [entry["process_id"] for entry in report["native_processes"]])
        self.assertFalse(report["native_exit_confirmed"])
        self.assertFalse(report["native_processes"][1]["exit_confirmed"])

    def test_native_exit_proof_survives_removed_registry_record(self):
        opened = []

        class Handle:
            def __init__(self, identity):
                self.closed = False
                opened.append(self)

            def wait(self, seconds):
                return True

            def close(self):
                self.closed = True

        report = {}
        owner = NativeOwners({"process_id": 321, "created_at_utc_filetime": 456}, report, registry=Path("unused"), opener=Handle)
        owner.records = lambda: [{"controller": {"processId": 321, "startedAtUtcFileTime": 456},
            "worker": {"processId": 322, "startedAtUtcFileTime": 457},
            "excel": {"processId": 323, "startedAtUtcFileTime": 458}, "ownsExcel": True}]
        owner.capture(required=True)
        owner.records = lambda: []
        owner.finish(seconds=0)
        self.assertTrue(report["native_exit_confirmed"])
        self.assertEqual(2, len(opened))
        self.assertTrue(all(handle.closed for handle in opened))
        self.assertEqual({}, owner.handles)

    def test_native_capture_failure_keeps_exact_identity_without_adopting_pid(self):
        def changed(identity):
            raise SmokeError("Native process identity changed; PID reuse was refused.")

        report = {}
        owner = NativeOwners({"process_id": 321, "created_at_utc_filetime": 456}, report, registry=Path("unused"), opener=changed)
        owner.records = lambda: [{"controller": {"processId": 321, "startedAtUtcFileTime": 456},
            "worker": {"processId": 322, "startedAtUtcFileTime": 457}, "excel": None}]
        with self.assertRaises(SmokeError):
            owner.capture()
        self.assertEqual({}, owner.handles)
        self.assertEqual(322, report["native_ownership_failures"][0]["process_id"])
        self.assertEqual(457, report["native_ownership_failures"][0]["created_at_utc_filetime"])

    def test_cli_retained_identity_is_recorded_even_on_failed_response(self):
        report = {}
        item = {"success": False, "data": {"retainedExcelProcesses": [{"processId": 321, "startTime": "2026-10-29T12:00:00Z"}],
            "retainedExcelProcessCount": 1, "retainedExcelError": "private-marker", "shutdownWarning": "private-marker"}}
        with self.assertRaises(SmokeError):
            retained_cli_state(item, report, "stop")
        state = report["cli_native_retention"][0]
        self.assertEqual(321, state["processes"][0]["process_id"])
        self.assertTrue(state["retained_error_present"])
        self.assertNotIn("private-marker", json.dumps(report))

    @unittest.skipUnless(os.name == "nt", "Windows query/synchronize handles")
    def test_retained_process_handle_survives_exit_and_refuses_pid_reuse(self):
        benchmark = helper("benchmark_handle_test", Path(__file__).with_name("fork_cli_batch_benchmark.py"))
        child = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(0.5)"],
                                 stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                 creationflags=subprocess.CREATE_NO_WINDOW)
        retained = None
        try:
            identity = benchmark.owned_process_identity(child)
            pair = (identity["process_id"], identity["created_at_utc_filetime"])
            with self.assertRaises(SmokeError):
                NativeProcessHandle((pair[0], pair[1] + 1))
            retained = NativeProcessHandle(pair)
            self.assertFalse(retained.wait(0))
            child.wait(timeout=5)
            self.assertTrue(retained.wait(0))
        finally:
            if retained is not None:
                retained.close()
            if child.poll() is None:
                child.kill()
            child.wait(timeout=5)

    def test_cleanup_selects_only_exact_nonattached_fixture(self):
        path = os.path.abspath("synthetic.xlsm")
        entries = [{"filePath": path, "isAttached": False, "id": "owned", "canClose": True},
                   {"filePath": os.path.abspath("unrelated.xlsm"), "isAttached": False, "id": "foreign", "canClose": True}]
        self.assertEqual([("owned", True)], owned_fixture_sessions(entries, path, "id"))
        entries[0]["isAttached"] = True
        with self.assertRaises(SmokeError):
            owned_fixture_sessions(entries, path, "id")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--mcp-exe", type=Path)
    parser.add_argument("--cli-exe", type=Path)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--protocol-helper", type=Path, default=Path(__file__).with_name("fork_protocol_smoke.py"))
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        return 0 if unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(ParserTests)).wasSuccessful() else 1
    if args.mcp_exe is None or args.cli_exe is None:
        parser.error("--mcp-exe and --cli-exe are required for the native smoke.")
    report = {"success": False, "scope": "Disposable synthetic XLSM; packaged entry points and embedded interop", "cases": []}
    root = None
    try:
        require(os.name == "nt", "Packaged native smoke requires Windows.")
        benchmark = helper("fork_cli_batch_benchmark", Path(__file__).with_name("fork_cli_batch_benchmark.py"))
        protocol = helper("fork_protocol_smoke", args.protocol_helper)
        executables = {"mcp": args.mcp_exe.resolve(strict=True), "cli": args.cli_exe.resolve(strict=True)}
        before = {name: benchmark.binary_identity(path) for name, path in executables.items()}
        report["provenance"] = {"binaries": before, "script_sha256": benchmark.file_identity(Path(__file__))["sha256"],
                                "protocol_helper_sha256": benchmark.file_identity(args.protocol_helper)["sha256"],
                                "benchmark_helper_sha256": benchmark.file_identity(Path(__file__).with_name("fork_cli_batch_benchmark.py"))["sha256"]}
        root = Path(tempfile.mkdtemp(prefix="excelmcp-packaged-native-"))
        for name, factory, run in (("mcp", lambda case: McpHost(executables["mcp"], protocol, case), mcp_case),
                                   ("cli", lambda case: CliHost(executables["cli"], benchmark, case), cli_case)):
            case = {"entry_point": name, "passed": False}
            report["cases"].append(case)
            host = factory(case)
            host.native = None
            path = root / f"{name}.xlsm"
            try:
                host.start()
                if name == "mcp":
                    case["owned_host"] = benchmark.owned_process_identity(host.process)
                host.native = NativeOwners(case["owned_host"], case)
                run(host, path, benchmark)
                case["passed"] = True
            finally:
                primary = sys.exception()
                failures = []
                try:
                    if host.native is None and host.process is not None:
                        case.setdefault("owned_host", benchmark.owned_process_identity(host.process))
                        host.native = NativeOwners(case["owned_host"], case)
                    if host.native is not None:
                        host.native.capture()
                except Exception as error:
                    failures.append(type(error).__name__)
                    case.setdefault("native_ownership_failures", []).append({"phase": "before_cleanup", "exception_type": type(error).__name__})
                try:
                    if host.process is not None and host.process.poll() is None:
                        cleanup_sessions(host, path, case)
                except Exception as error:
                    failures.append(type(error).__name__)
                try:
                    host.shutdown()
                except Exception as error:
                    failures.append(type(error).__name__)
                try:
                    require(host.native is not None, "Native ownership capture was not initialized.")
                    host.native.finish()
                except Exception as error:
                    failures.append(type(error).__name__)
                if failures:
                    case["cleanup_failure_types"] = failures
                    case["passed"] = False
                    if primary is None:
                        raise SmokeError("Owned fixture or host cleanup was not confirmed.")
        require(all(before[name] == benchmark.binary_identity(path) for name, path in executables.items()), "Packaged binaries changed during the smoke.")
        report["binary_hashes_unchanged"] = True
        report["success"] = all(case["passed"] and case.get("fixture_sessions_closed") and case.get("host_exit_confirmed") and case.get("native_exit_confirmed") for case in report["cases"])
    except Exception as error:
        report["error"] = str(error) if isinstance(error, SmokeError) else f"Packaged smoke failed ({type(error).__name__}); raw details withheld."
    finally:
        cleanup_fixtures(root, report)
        encoded = json.dumps(report, indent=2) + "\n"
        if args.report:
            args.report.write_text(encoded, encoding="utf-8")
        print(encoded, end="")
    return 0 if report["success"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
