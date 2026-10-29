"""Exercise the actual protocol finalizer under simultaneous primary/cleanup failures."""
import io
from contextlib import redirect_stdout
import json
import os
import queue
import subprocess
import sys
import tempfile
from pathlib import Path
import unittest
from types import SimpleNamespace
from unittest.mock import patch

import fork_protocol_smoke as protocol
from fork_protocol_smoke import cleanup, process_metrics


class FailedProcess:
    pid = 123
    returncode = None

    def __init__(self):
        self.stdin, self.stdout, self.stderr = io.StringIO(), io.StringIO(), io.StringIO()
        self.waited = False

    def poll(self):
        return None

    def kill(self):
        raise PermissionError("private process details")

    def wait(self, timeout):
        self.waited = True
        raise subprocess.TimeoutExpired("private command", timeout)


class ReaderDeadlineTests(unittest.TestCase):
    def response_after(self, delay, **options):
        payload = {"success": False, "exists": False, "message": "missing fixture"}
        reader = protocol.Reader(io.StringIO(json.dumps({
            "jsonrpc": "2.0", "id": 4, "result": payload}) + "\n"), protocol=True)
        reader.thread.join(timeout=1)
        self.assertFalse(reader.thread.is_alive())
        now = [0.0]
        take = reader.items.get

        def delayed_get(timeout):
            now[0] += timeout
            if now[0] < delay:
                raise queue.Empty
            return take(timeout=0)

        with patch.object(protocol.time, "monotonic", side_effect=lambda: now[0]), \
                patch.object(reader.items, "get", side_effect=delayed_get):
            try:
                return reader.response(4, **options), now[0]
            except TimeoutError:
                self.assertLessEqual(now[0], options.get("timeout_seconds", 8) + 0.02)
                raise

    def test_default_deadline_still_rejects_nine_second_response(self):
        with self.assertRaises(TimeoutError):
            self.response_after(9)

    def test_explicit_worker_budget_accepts_delayed_response_without_changing_payload(self):
        result, elapsed = self.response_after(9, timeout_seconds=30)
        self.assertEqual({"success": False, "exists": False, "message": "missing fixture"}, result)
        self.assertGreaterEqual(elapsed, 9)
        self.assertLess(elapsed, 9.2)

    def test_explicit_worker_budget_remains_bounded_without_response(self):
        with self.assertRaises(TimeoutError):
            self.response_after(float("inf"), timeout_seconds=30)

    def test_tool_forwards_budget_and_preserves_error_parity(self):
        structured = {"success": False, "exists": False}
        reader = SimpleNamespace(response=lambda request_id, **options: {
            "content": [{"type": "text", "text": json.dumps(structured)}],
            "structuredContent": structured, "isError": True})
        with patch.object(protocol, "send") as send, \
                patch.object(reader, "response", wraps=reader.response) as response:
            self.assertEqual(structured, protocol.tool(None, reader, 4,
                {"action": "test"}, expected_error=True, timeout_seconds=30))
        response.assert_called_once_with(4, timeout_seconds=30)
        send.assert_called_once_with(None, 4, "tools/call", {
            "name": "file_read", "arguments": {"action": "test"}})

    def test_only_missing_file_call_uses_worker_budget(self):
        process = SimpleNamespace(stdin=io.StringIO(), wait=lambda timeout: 0)
        out, err = object(), object()
        with patch.object(protocol, "start", return_value=(process, out, err)), \
                patch.object(protocol, "initialize", return_value=33), \
                patch.object(protocol, "tool", side_effect=[
                    {"success": True, "sessions": []}, {"success": False, "exists": False}]) as tool, \
                patch.object(protocol, "process_metrics", return_value={}), \
                patch.object(protocol, "cleanup_case"):
            self.assertTrue(protocol.protocol_case(None, 0)["passed"])
        self.assertEqual({}, tool.call_args_list[0].kwargs)
        self.assertEqual({"expected_error": True, "timeout_seconds": 30}, tool.call_args_list[1].kwargs)
        self.assertEqual("test", tool.call_args_list[1].args[3]["action"])


class CleanupTests(unittest.TestCase):
    @unittest.skipUnless(os.name == "nt", "Windows process counters")
    def test_resource_counters_use_the_owned_child_handle(self):
        process = subprocess.Popen([sys.executable, "-c", "import sys; sys.stdin.read()"],
                                   stdin=subprocess.PIPE, creationflags=subprocess.CREATE_NO_WINDOW)
        try:
            counters = process_metrics(process)
            self.assertGreater(counters["handles"], 0)
            self.assertGreater(counters["privateBytes"], 0)
            self.assertGreater(counters["workingSetBytes"], 0)
        finally:
            process.stdin.close()
            try:
                process.wait(timeout=3)
            finally:
                if process.poll() is None:
                    process.kill()
                    process.wait(timeout=3)

    def test_real_finalizer_keeps_primary_and_structured_cleanup(self):
        process, receipt = FailedProcess(), None
        primary = ValueError("original operation")
        observed = None
        try:
            try:
                raise primary
            finally:
                receipt = cleanup(process, [])
        except ValueError as error:
            observed = error
        self.assertIs(observed, primary)
        self.assertIsInstance(receipt, dict)
        self.assertTrue(process.waited)
        self.assertFalse(receipt["passed"])
        self.assertFalse(receipt["exit_confirmed"])
        self.assertEqual(["PermissionError", "TimeoutExpired"],
                         [item["exception_type"] for item in receipt["failures"]])
        self.assertNotIn("private", json.dumps(receipt))
        self.assertTrue(all(stream.closed for stream in (process.stdin, process.stdout, process.stderr)))

    def test_failed_protocol_case_keeps_sanitized_primary_and_cleanup_in_report(self):
        process = FailedProcess()
        primary = ValueError("private operation details")
        receipt = {"case": "initialize-list-negative-parity-eof", "passed": False}
        thread = SimpleNamespace(join=lambda timeout: None, is_alive=lambda: False)
        out = SimpleNamespace(thread=thread, failure=None, stream=process.stdout)
        err = SimpleNamespace(thread=thread, failure=None, stream=process.stderr)
        with patch.object(protocol, "start", return_value=(process, out, err)), \
                patch.object(protocol, "initialize", return_value=33), \
                patch.object(protocol, "tool", side_effect=primary):
            with self.assertRaises(ValueError) as caught:
                protocol.protocol_case(None, 0, receipt=receipt)
        self.assertIs(primary, caught.exception)
        self.assertFalse(receipt["passed"])
        self.assertEqual("ValueError", receipt["error"]["exception_type"])
        self.assertEqual(123, receipt["cleanup"]["owned_process_id"])
        self.assertEqual(["PermissionError", "TimeoutExpired"],
                         [entry["exception_type"] for entry in receipt["cleanup"]["failures"]])
        self.assertNotIn("private", json.dumps(receipt))

    def test_main_emits_failed_case_with_actual_cleanup_receipt(self):
        process = FailedProcess()
        primary = ValueError("private operation details")
        thread = SimpleNamespace(join=lambda timeout: None, is_alive=lambda: False)
        out = SimpleNamespace(thread=thread, failure=None, stream=process.stdout)
        err = SimpleNamespace(thread=thread, failure=None, stream=process.stderr)
        with tempfile.TemporaryDirectory(prefix="excelmcp-receipt-unit-") as folder:
            exe, report_path = Path(folder) / "synthetic.exe", Path(folder) / "report.json"
            exe.write_bytes(b"synthetic nonexecutable fixture")
            with patch.object(protocol, "os", SimpleNamespace(name="nt")), \
                    patch.object(protocol.sys, "argv", ["smoke", "--exe", str(exe), "--report", str(report_path)]), \
                    patch.object(protocol, "start", return_value=(process, out, err)), \
                    patch.object(protocol, "initialize", return_value=33), \
                    patch.object(protocol, "tool", side_effect=primary), redirect_stdout(io.StringIO()):
                with self.assertRaises(ValueError) as caught:
                    protocol.main()
            self.assertIs(primary, caught.exception)
            report = json.loads(report_path.read_text(encoding="utf-8"))
        self.assertFalse(report["allPassed"])
        self.assertEqual("ValueError", report["error"]["exception_type"])
        self.assertEqual(1, len(report["cases"]))
        case = report["cases"][0]
        self.assertFalse(case["passed"])
        self.assertEqual(123, case["cleanup"]["owned_process_id"])
        self.assertEqual(["PermissionError", "TimeoutExpired"], [entry["exception_type"] for entry in case["cleanup"]["failures"]])
        self.assertNotIn("private", json.dumps(report))

    def test_without_primary_reports_both_cleanup_failures(self):
        with self.assertRaises(ExceptionGroup) as caught:
            cleanup(FailedProcess(), [])
        self.assertEqual([PermissionError, subprocess.TimeoutExpired],
                         [type(error) for error in caught.exception.exceptions])


if __name__ == "__main__":
    unittest.main()
