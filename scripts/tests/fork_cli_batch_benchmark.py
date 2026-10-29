"""Benchmark session.list launch overhead on a private CLI daemon; never opens Excel.

Python standard library only. Run --self-test without a CLI, or supply --cli
and optionally --output for a local JSON receipt. Defaults: three rounds,
eight session.list operations per mode, with both modes warmed first.
"""

from __future__ import annotations

import argparse
import ctypes
from ctypes import wintypes
from datetime import datetime, timezone
import hashlib
import io
import json
import math
import os
from pathlib import Path
import platform
import re
import statistics
import subprocess
import sys
import tempfile
import time
import unittest
import uuid


class BenchmarkError(RuntimeError):
    """A diagnostic safe to include in a receipt, without raw CLI output."""


def decode_envelopes(text, command, count=1, indexed=False, require_success=True):
    """Validate v2.6 output and every batch item, rather than legacy result fields."""
    lines = [line for line in text.splitlines() if line.strip()]
    if len(lines) != count:
        raise BenchmarkError("CLI returned an unexpected JSON item count.")
    results = []
    for index, line in enumerate(lines):
        try:
            item = json.loads(line)
        except (ValueError, TypeError) as error:
            raise BenchmarkError("CLI stdout was not JSON-only NDJSON.") from error
        if not isinstance(item, dict) or item.get("command") != command:
            raise BenchmarkError("CLI returned an unexpected command envelope.")
        duration = item.get("durationMs")
        if (type(item.get("success")) is not bool or type(duration) not in (int, float)
                or not math.isfinite(duration) or duration < 0):
            raise BenchmarkError("CLI envelope has invalid outcome or duration metadata.")
        if indexed and (type(item.get("index")) is not int or item["index"] != index):
            raise BenchmarkError("Batch indices were missing, duplicated, or reordered.")
        if not item["success"]:
            if not isinstance(item.get("error"), dict):
                raise BenchmarkError("Failed CLI envelope lacks its structured error.")
            if require_success:
                category = item["error"].get("category", "unspecified")
                if not isinstance(category, str) or not re.fullmatch(r"[\w.-]{1,80}", category):
                    category = "unspecified"
                raise BenchmarkError(f"CLI reported failure ({category}); raw output withheld.")
        elif item.get("error") is not None:
            raise BenchmarkError("Successful CLI envelope also reports an error.")
        results.append(item)
    return results


def empty_sessions(item):
    data = item.get("data")
    if (not isinstance(data, dict) or data.get("sessions") != []
            or type(data.get("count")) is not int or data["count"] != 0):
        raise BenchmarkError("The private daemon did not report exactly zero workbook sessions.")
    if "daemonState" in data and data["daemonState"] != "running":
        raise BenchmarkError("A session list did not reach the running private daemon.")


def validate_status(item, running, process_id=None):
    data = item.get("data")
    if not isinstance(data, dict):
        raise BenchmarkError("Service status did not contain a data object.")
    expected = "running" if running else "stopped"
    if data.get("running") is not running or data.get("daemonState") != expected:
        raise BenchmarkError("Private service state was not confirmed.")
    for key in ("sessionCount", "retainedExcelProcessCount"):
        if type(data.get(key)) is not int or data[key] != 0:
            raise BenchmarkError("Private service reported sessions or retained Excel processes.")
    if data.get("retainedExcelProcesses") != []:
        raise BenchmarkError("Retained Excel process list was not confirmed empty.")
    if process_id is not None and (type(data.get("processId")) is not int
                                   or data["processId"] != process_id):
        raise BenchmarkError("Service PID does not match the exact child created by this benchmark.")
    return data


def file_identity(path):
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return {"name": path.name, "bytes": path.stat().st_size, "sha256": digest.hexdigest()}


def binary_identity(cli):
    # Source-built apphosts load adjacent assemblies; hashing the .exe alone is insufficient.
    files = {cli}
    if cli.with_suffix(".dll").is_file():
        files.update(cli.parent.glob("*.dll"))
        files.update(cli.parent.glob("*.deps.json"))
        files.update(cli.parent.glob("*.runtimeconfig.json"))
    return [file_identity(path) for path in sorted(files, key=lambda path: path.name.lower())]


def stream_identity(stream):
    stream.seek(0)
    digest = hashlib.sha256()
    length = 0
    for block in iter(lambda: stream.read(1024 * 1024), b""):
        length += len(block)
        digest.update(block)
    return {"bytes": length, "sha256": digest.hexdigest()}


def owned_process_identity(process):
    """Capture creation time through the retained Windows Popen process handle."""
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.GetProcessTimes.argtypes = [wintypes.HANDLE] + [ctypes.POINTER(wintypes.FILETIME)] * 4
    kernel.GetProcessTimes.restype = wintypes.BOOL
    created, exited, kernel_time, user_time = (wintypes.FILETIME() for _ in range(4))
    if not kernel.GetProcessTimes(int(process._handle), ctypes.byref(created), ctypes.byref(exited),
                                  ctypes.byref(kernel_time), ctypes.byref(user_time)):
        raise BenchmarkError("Could not capture the owned daemon creation identity.")
    return {"process_id": process.pid,
            "created_at_utc_filetime": (created.dwHighDateTime << 32) | created.dwLowDateTime}


class Runner:
    def __init__(self, cli, environment):
        self.cli = cli
        self.environment = environment
        self.records = []

    def call(self, arguments, label, command, *, count=1, indexed=False, input_text=None,
             require_success=True, timeout=20):
        started = time.perf_counter()
        try:
            completed = subprocess.run([str(self.cli), "--quiet", *arguments], input=input_text,
                                       capture_output=True, text=True, encoding="utf-8",
                                       env=self.environment, timeout=timeout,
                                       creationflags=subprocess.CREATE_NO_WINDOW)
        except subprocess.TimeoutExpired as error:
            self.records.append({"phase": label, "wall_seconds": time.perf_counter() - started,
                                 "timed_out": True})
            raise BenchmarkError("A CLI subprocess exceeded its bounded wait.") from error
        except OSError as error:
            raise BenchmarkError("A CLI subprocess could not be launched.") from error
        self.records.append({"phase": label, "wall_seconds": time.perf_counter() - started,
                             "exit_code": completed.returncode,
                             "stdout_bytes": len(completed.stdout.encode("utf-8")),
                             "stderr_bytes": len(completed.stderr.encode("utf-8"))})
        items = decode_envelopes(completed.stdout, command, count, indexed, require_success)
        self.records[-1]["cli_duration_ms"] = [item["durationMs"] for item in items]
        if completed.returncode != 0 and (require_success or all(item["success"] for item in items)):
            raise BenchmarkError("CLI subprocess exit code did not confirm success.")
        return items


def finish_owned_child(process, cleanup, primary=None):
    """Finalize only the retained Popen handle; keep the first failure and receipt."""
    if process is None:
        return primary
    cleanup["owned_child_pid"] = process.pid
    cleanup["no_lingering_owned_daemon"] = False
    cleanup["owned_child_wait_confirmed"] = False
    failures = cleanup.setdefault("failures", [])
    try:
        alive = process.poll() is None
    except Exception as error:
        failures.append({"phase": "poll", "exception_type": type(error).__name__})
        alive = True
    if alive:
        cleanup["forced_termination"] = True
        try:
            # Popen retains the exact created process handle; never adopt by PID/name.
            process.kill()
        except Exception as error:
            failures.append({"phase": "kill", "exception_type": type(error).__name__})
    try:
        exit_code = process.wait(timeout=5)
        cleanup["owned_child_exit_code"] = exit_code
        cleanup["owned_child_wait_confirmed"] = True
        cleanup["no_lingering_owned_daemon"] = True
    except Exception as error:
        failures.append({"phase": "wait", "exception_type": type(error).__name__})
    if primary is None and (failures or cleanup.get("forced_termination")):
        return BenchmarkError("Owned child cleanup required forced termination or could not be confirmed; see cleanup receipt.")
    return primary


def benchmark(cli, rounds=3, commands=8, source_revision=None):
    if os.name != "nt":
        raise BenchmarkError("The native CLI benchmark requires Windows.")
    if not cli.is_file() or cli.suffix.lower() != ".exe":
        raise BenchmarkError("--cli must select an existing Windows executable.")
    started = time.perf_counter()
    pipe = "excelmcp-benchmark-" + uuid.uuid4().hex
    environment = dict(os.environ, EXCELMCP_CLI_PIPE=pipe, EXCELMCP_TELEMETRY_MODE="off",
                       DOTNET_CLI_TELEMETRY_OPTOUT="1")
    runner = Runner(cli, environment)
    identity_before = binary_identity(cli)
    receipt = {
        "schema": 1, "success": False, "created_utc": datetime.now(timezone.utc).isoformat(),
        "operation": "Zero-session session.list: separate CLI launches versus one ordered batch",
        "scope": "Process/connection overhead; excludes Excel calculation, workbook work and printing",
        "environment": {"platform": platform.platform(), "machine": platform.machine(),
                        "python": platform.python_version(), "cpu_count": os.cpu_count(),
                        "dotnet_root_configured": bool(environment.get("DOTNET_ROOT")),
                        "telemetry": "off"},
        "provenance": {"script_sha256": file_identity(Path(__file__))["sha256"],
                       "binary_files": identity_before, "supplied_source_revision": source_revision},
        "private_pipe": True, "rounds": rounds, "operations_per_mode": commands,
        "workbook_open_or_create_commands": 0, "validated_empty_session_lists": 0,
        "warmup": "One single session.list and one full batch before measured rounds",
        "subprocesses": runner.records, "runs": [], "cleanup": {},
    }
    daemon = None
    failure = None
    batch_input = json.dumps([{"command": "session.list"}] * commands)

    def run_mode(mode, label):
        begin = time.perf_counter()
        first_record = len(runner.records)
        if mode == "singles":
            for index in range(commands):
                empty_sessions(runner.call(["session", "list"], f"{label}.{index}", "session.list")[0])
                receipt["validated_empty_session_lists"] += 1
        else:
            items = runner.call(["batch", "--input", "-", "--stop-on-error"], label,
                                "session.list", count=commands, indexed=True, input_text=batch_input)
            for item in items:
                empty_sessions(item)
                receipt["validated_empty_session_lists"] += 1
        return {"mode": mode, "wall_seconds": time.perf_counter() - begin,
                "subprocess_seconds": sum(item["wall_seconds"] for item in runner.records[first_record:]),
                "cli_duration_ms": [duration for item in runner.records[first_record:]
                                    for duration in item["cli_duration_ms"]]}

    with tempfile.TemporaryFile() as stdout, tempfile.TemporaryFile() as stderr:
        try:
            validate_status(runner.call(["service", "status"], "preflight", "service.status")[0], False, 0)
            launch = time.perf_counter()
            daemon = subprocess.Popen([str(cli), "--quiet", "service", "run", "--pipe-name", pipe],
                                      stdin=subprocess.DEVNULL, stdout=stdout, stderr=stderr,
                                      env=environment, creationflags=subprocess.CREATE_NO_WINDOW)
            receipt["daemon_launch_seconds"] = time.perf_counter() - launch
            receipt["owned_daemon"] = owned_process_identity(daemon)
            deadline = time.monotonic() + 20
            while True:
                if daemon.poll() is not None:
                    raise BenchmarkError("Owned private daemon exited during startup.")
                status = runner.call(["service", "status"], "readiness", "service.status", require_success=False)[0]
                if status["success"] and status.get("data", {}).get("running") is True:
                    validate_status(status, True, daemon.pid)
                    break
                if time.monotonic() >= deadline:
                    raise BenchmarkError("Owned private daemon readiness was not confirmed.")
                time.sleep(0.1)
            receipt["startup_ready_seconds"] = time.perf_counter() - launch
            empty_sessions(runner.call(["session", "list"], "warmup.single", "session.list")[0])
            receipt["validated_empty_session_lists"] += 1
            run_mode("batch", "warmup.batch")
            for number in range(rounds):
                order = ("singles", "batch") if number % 2 == 0 else ("batch", "singles")
                measurements = [run_mode(mode, f"round.{number}.{mode}") for mode in order]
                receipt["runs"].append({"round": number + 1, "order": list(order), "measurements": measurements})
                validate_status(runner.call(["service", "status"], f"round.{number}.status", "service.status")[0], True, daemon.pid)
        except Exception as error:
            failure = error if isinstance(error, BenchmarkError) else BenchmarkError(f"Benchmark failed ({type(error).__name__}); raw details withheld.")
        finally:
            cleanup_started = time.perf_counter()
            cleanup = receipt["cleanup"]
            cleanup["forced_termination"] = False
            try:
                if daemon is not None:
                    if daemon.poll() is None:
                        stop = runner.call(["service", "stop"], "cleanup.stop", "service.stop")[0]
                        stop_data = stop.get("data", {})
                        if stop_data.get("retainedExcelProcessCount") != 0 or stop_data.get("shutdownWarning"):
                            raise BenchmarkError("Private daemon shutdown reported retained work or a warning.")
                    daemon.wait(timeout=10)
                    cleanup["owned_child_exit_code"] = daemon.returncode
                    cleanup["owned_child_wait_confirmed"] = True
                    if daemon.returncode != 0:
                        raise BenchmarkError("Owned daemon exited with a nonzero code.")
                    validate_status(runner.call(["service", "status"], "cleanup.status", "service.status")[0], False, 0)
                    cleanup["observed_service_stopped"] = True
                    cleanup["no_lingering_owned_daemon"] = daemon.poll() is not None
            except Exception as error:
                if failure is None:
                    failure = error if isinstance(error, BenchmarkError) else BenchmarkError("Owned daemon shutdown was not confirmed.")
            finally:
                failure = finish_owned_child(daemon, cleanup, failure)
                cleanup["wall_seconds"] = time.perf_counter() - cleanup_started
                for label, stream in (("stdout", stdout), ("stderr", stderr)):
                    try:
                        identity = stream_identity(stream)
                        cleanup[f"daemon_{label}_bytes"] = identity["bytes"]
                        cleanup[f"daemon_{label}_sha256"] = identity["sha256"]
                    except Exception as error:
                        cleanup.setdefault("failures", []).append({"phase": f"hash_{label}", "exception_type": type(error).__name__})
                        if failure is None:
                            failure = BenchmarkError("Owned daemon log provenance could not be confirmed.")
    try:
        receipt["provenance"]["binary_unchanged"] = identity_before == binary_identity(cli)
    except Exception as error:
        receipt["provenance"]["binary_unchanged"] = False
        receipt["provenance"]["recheck_error_type"] = type(error).__name__
        if failure is None:
            failure = BenchmarkError("CLI binary provenance could not be rechecked.")
    if not receipt["provenance"]["binary_unchanged"] and failure is None:
        failure = BenchmarkError("CLI binary files changed during the benchmark; measurements are not comparable.")
    if failure is None and not receipt["cleanup"].get("forced_termination"):
        medians = {mode: statistics.median(measurement["wall_seconds"]
                   for run in receipt["runs"] for measurement in run["measurements"] if measurement["mode"] == mode)
                   for mode in ("singles", "batch")}
        receipt["median_seconds"] = medians
        receipt["batch_speedup_ratio"] = medians["singles"] / medians["batch"]
        receipt["success"] = True
    else:
        receipt["error"] = str(failure or BenchmarkError("Cleanup required forced termination; timings are not accepted."))
    receipt["total_wall_seconds"] = time.perf_counter() - started
    return receipt


class ParserTests(unittest.TestCase):
    def item(self, **overrides):
        value = {"command": "session.list", "success": True, "durationMs": 2,
                 "data": {"sessions": [], "count": 0}}
        value.update(overrides)
        return value

    def test_v26_single_and_indexed_batch(self):
        empty_sessions(decode_envelopes(json.dumps(self.item()), "session.list")[0])
        text = "\n".join(json.dumps(self.item(index=index)) for index in range(8))
        self.assertEqual(8, len(decode_envelopes(text, "session.list", 8, True)))

    def test_legacy_missing_data_is_rejected(self):
        with self.assertRaises(BenchmarkError):
            empty_sessions(decode_envelopes(json.dumps(self.item(data=None, sessions=[])), "session.list")[0])

    def test_batch_count_order_failure_and_bad_metadata_are_rejected(self):
        for text in ("", "banner\n{}", json.dumps(self.item(index=1)),
                     json.dumps(self.item(durationMs=True)), json.dumps(self.item(durationMs=-1)),
                     json.dumps(self.item(success=1))):
            with self.subTest(text=text), self.assertRaises(BenchmarkError):
                decode_envelopes(text, "session.list", indexed=True)

    def test_failed_envelope_never_exposes_raw_message(self):
        text = json.dumps(self.item(success=False, error={"category": "InvalidInput", "message": "private-marker"}))
        with self.assertRaises(BenchmarkError) as failure:
            decode_envelopes(text, "session.list")
        self.assertNotIn("private-marker", str(failure.exception))

    def test_nonempty_sessions_and_boolean_count_are_rejected(self):
        for data in ({"sessions": [{"filePath": "private-marker"}], "count": 1},
                     {"sessions": [], "count": False}):
            with self.subTest(data=data), self.assertRaises(BenchmarkError):
                empty_sessions(self.item(data=data))

    def test_status_requires_exact_owned_pid_and_no_excel(self):
        data = {"running": True, "daemonState": "running", "processId": 42,
                "sessionCount": 0, "retainedExcelProcessCount": 0, "retainedExcelProcesses": []}
        validate_status({"data": data}, True, 42)
        for change in ({"processId": 43}, {"sessionCount": 1}, {"retainedExcelProcessCount": 1}):
            with self.subTest(change=change), self.assertRaises(BenchmarkError):
                validate_status({"data": dict(data, **change)}, True, 42)

    def test_cleanup_preserves_primary_identity_and_failed_receipt(self):
        class Child:
            pid = 321
            returncode = None

            def __init__(self, denied):
                self.denied = denied

            def poll(self):
                return None

            def kill(self):
                if self.denied:
                    raise PermissionError("private-marker")

            def wait(self, timeout):
                raise subprocess.TimeoutExpired("private-marker", timeout)

        for denied in (True, False):
            with self.subTest(kill_denied=denied):
                primary = BenchmarkError("Original measurement failed.")
                receipt = {"success": False, "owned_daemon": {"pid": 321, "creation_filetime": 456}, "cleanup": {}}
                result = finish_owned_child(Child(denied), receipt["cleanup"], primary)
                self.assertIs(primary, result)
                self.assertFalse(receipt["success"])
                self.assertEqual({"pid": 321, "creation_filetime": 456}, receipt["owned_daemon"])
                self.assertFalse(receipt["cleanup"]["no_lingering_owned_daemon"])
                phases = {item["phase"] for item in receipt["cleanup"]["failures"]}
                self.assertIn("wait", phases)
                self.assertEqual(denied, "kill" in phases)
                self.assertNotIn("private-marker", json.dumps(receipt))

    def test_log_provenance_hashes_without_returning_contents(self):
        value = b"private-marker" * 100_000
        result = stream_identity(io.BytesIO(value))
        self.assertEqual({"bytes": len(value), "sha256": hashlib.sha256(value).hexdigest()}, result)
        self.assertNotIn("private-marker", json.dumps(result))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--cli", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--rounds", type=int, default=3)
    parser.add_argument("--commands", type=int, default=8)
    parser.add_argument("--source-revision", help="Optional reviewed source commit (40 or 64 hex characters); binary hashes remain authoritative.")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        return 0 if unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(ParserTests)).wasSuccessful() else 1
    if args.cli is None or not 1 <= args.rounds <= 100 or not 1 <= args.commands <= 64:
        parser.error("Supply --cli, 1-100 rounds and 1-64 commands per mode.")
    if args.source_revision is not None and not re.fullmatch(r"(?:[0-9a-fA-F]{40}|[0-9a-fA-F]{64})", args.source_revision):
        parser.error("--source-revision must be a full hexadecimal commit identifier.")
    try:
        result = benchmark(args.cli.resolve(), args.rounds, args.commands, args.source_revision)
    except Exception as error:
        result = {"success": False, "error": str(error) if isinstance(error, BenchmarkError)
                  else f"Benchmark setup failed ({type(error).__name__}); raw details withheld."}
    encoded = json.dumps(result, indent=2) + "\n"
    if args.output:
        args.output.write_text(encoded, encoding="utf-8")
    print(encoded, end="")
    return 0 if result["success"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
