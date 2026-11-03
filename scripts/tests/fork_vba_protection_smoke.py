"""Native synthetic VBA protection/authentication acceptance through MCP or CLI.

Requires Windows, desktop Excel, built product and the sibling native auth probe.
Passwords travel through stdin/public protocol only. Never adopts existing sessions
or terminates Excel. Every fixture/worker/Excel must close through the product.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path

sys.dont_write_bytecode = True
HERE = Path(__file__).resolve().parent


def load(name):
    import importlib.util
    spec = importlib.util.spec_from_file_location(name, HERE / (name + ".py"))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


smoke = load("fork_packaged_excel_smoke")
benchmark = load("fork_cli_batch_benchmark")
protocol = load("fork_protocol_smoke")
require = smoke.require

PROBE_CODE = '''Public Function Marker() As Long
Marker = 73
End Function
Public Function FixtureEnvironment() As String
FixtureEnvironment = Application.Version & "|" & CStr(Application.Build) & "|" & Application.OperatingSystem
End Function
Public Function PrepareCells() As Boolean
ThisWorkbook.Worksheets(1).Range("A1").Value2 = "Synthetic protection fixture"
ThisWorkbook.Worksheets(1).Range("B1").Value2 = 73
PrepareCells = True
End Function
Public Function CellsPreserved() As Boolean
CellsPreserved = (ThisWorkbook.Worksheets(1).Range("A1").Value2 = "Synthetic protection fixture") And _
 (ThisWorkbook.Worksheets(1).Range("B1").Value2 = 73)
End Function
Public Function OpenProbeProperties() As String
On Error GoTo Failed
Dim command As Object
Set Application.VBE.ActiveVBProject = ThisWorkbook.VBProject
Application.VBE.MainWindow.Visible = True
Set command = Application.VBE.CommandBars.FindControl(ID:=2578)
command.Execute
OpenProbeProperties = "PASS"
Exit Function
Failed:
OpenProbeProperties = "CALLER_ERROR:" & CStr(Err.Number)
End Function
Public Function HideProbeEditor() As Boolean
Application.VBE.MainWindow.Visible = False
HideProbeEditor = True
End Function'''


def invoke(host, identifier, procedure):
    return smoke.packaged_call(host, "vba.invoke", {
        "procedureName": "ProtectionProbe." + procedure, "arguments": []}, identifier)


def live_excel(native):
    native.capture(required=True)
    current = {excel for _, excel in smoke.selected_owned_records(native.records(), native.controller) if excel}
    live = [(key[1], key[2]) for key, (handle, _) in native.handles.items()
            if key[0] == "excel" and (key[1], key[2]) in current and not handle.wait(0)]
    require(len(live) == 1, "Authentication requires exactly one live owned synthetic Excel.")
    return live[0]


def authenticate(host, identifier, native, helper, directory, password, mode, label):
    pid, created = live_excel(native)
    identity, ready, result = (directory / (label + suffix) for suffix in ("-identity.json", "-ready.json", "-result.json"))
    identity.write_text(json.dumps({"processId": pid, "startedAt": created}), encoding="utf-8")
    with (directory / (label + "-auth.log")).open("xb") as log:
        child = subprocess.Popen([str(helper), "--identity", str(identity), "--mode", mode,
                                  "--ready", str(ready), "--result", str(result)],
                                 stdin=subprocess.PIPE, stdout=log, stderr=log,
                                 creationflags=subprocess.CREATE_NO_WINDOW)
        try:
            child.stdin.write(password.encode("utf-8"))
            child.stdin.close()
            deadline = time.monotonic() + 10
            while not ready.exists():
                require(child.poll() is None and time.monotonic() < deadline,
                        "The owned authentication helper did not become ready.")
                time.sleep(0.025)
            require(invoke(host, identifier, "OpenProbeProperties").get("value") == "PASS",
                    "The synthetic native dialog caller failed.")
            require(child.wait(timeout=25) == 0, "Native authentication or dialog cleanup failed.")
            receipt = json.loads(result.read_text(encoding="utf-8"))
            require(receipt.get("success") is True and receipt.get("expectedOutcomeObserved") is True
                    and receipt.get("cleanupConfirmed") is True and receipt.get("processIdentityConfirmed") is True,
                    "Native authentication did not prove the outcome and cleanup.")
        finally:
            # The helper owns a bounded finally-cleanup path. Never kill it before
            # that path can dismiss the exact introduced synthetic prompts.
            child.wait(timeout=30)


def run(options):
    directory = options.report.resolve().parent
    directory.mkdir(parents=True, exist_ok=False)
    report = {"scope": "Synthetic VBA protection through " + options.entrypoint, "success": False}
    password = ' A"b\\c Exact!7 '
    inputs = [Path(__file__), HERE / "fork_packaged_excel_smoke.py", HERE / "fork_cli_batch_benchmark.py",
              HERE / "fork_protocol_smoke.py", options.exe, options.auth_probe]
    for executable in (options.exe, options.auth_probe):
        inputs.extend(sorted(executable.parent.glob("*.dll")))
        inputs.extend(sorted(executable.parent.glob("*.json")))
    report["inputHashes"] = {str(path.resolve()): hashlib.sha256(path.read_bytes()).hexdigest() for path in inputs}
    snapshots = directory / "inputs"
    snapshots.mkdir()
    report["inputSnapshots"] = {}
    for index, (name, expected) in enumerate(report["inputHashes"].items()):
        snapshot = snapshots / (str(index).zfill(3) + "-" + Path(name).name)
        shutil.copyfile(name, snapshot)
        require(hashlib.sha256(snapshot.read_bytes()).hexdigest() == expected,
                "An input changed while preserving the native acceptance snapshot.")
        report["inputSnapshots"][name] = snapshot.relative_to(directory).as_posix()
    host = (smoke.McpHost(options.exe, protocol, report) if options.entrypoint == "mcp"
            else smoke.CliHost(options.exe, benchmark, report))
    is_mcp = options.entrypoint == "mcp"
    native, identifier, other = None, None, None
    path = directory / "ProtectionFixture.xlsm"

    def vba(action, args=None, session=None, expected_error=False):
        args = args or {}
        if is_mcp:
            tool = "vba_read" if action in ("status", "view") else "vba"
            keys = {"moduleName": "module_name"}
            return host.call(tool, dict(action=action, workbook_session_id=session or identifier,
                                       **{keys.get(key, key): value for key, value in args.items()}), expected_error)
        return host.call("vba." + action, args, session or identifier, expected_error)

    def create(file):
        result = (host.call("file", {"action": "create", "file_path": str(file)}) if is_mcp
                  else host.call("session.create", {"filePath": str(file)}))
        native.capture(required=True)
        return result[host.id_key]

    def reopen_protected():
        nonlocal identifier
        if is_mcp:
            host.call("file", {"action": "save", "workbook_session_id": identifier})
        else:
            host.call("session.save", {}, identifier)
        host.close(identifier)
        identifier = None
        opened = (host.call("file", {"action": "open", "file_path": str(path)}) if is_mcp
                  else host.direct(["session", "open", str(path)], "session.open"))
        identifier = opened[host.id_key]
        native.capture(required=True)
        require(vba("status")["protection"] == "Locked", "Save/fresh open did not retain project protection.")
        if is_mcp:
            host.call("file", {"action": "enable-macros", "workbook_session_id": identifier,
                               "approved": True, "allow_reopen": True})
        else:
            host.call("session.enable-macros", {"approved": True, "allowReopen": True}, identifier)
        native.capture(required=True)
        selected = smoke.owned_fixture_sessions(host.sessions(), path, host.id_key)
        require(len(selected) == 1, "The exact owned reopened fixture was not found.")
        identifier = selected[0][0]

    try:
        host.start()
        native = smoke.NativeOwners(benchmark.owned_process_identity(host.process), report)
        identifier = create(path)
        smoke.packaged_call(host, "vba.import", {"moduleName": "ProtectionProbe", "vbaCode": PROBE_CODE}, identifier)
        environment = invoke(host, identifier, "FixtureEnvironment").get("value")
        require(isinstance(environment, str) and len(environment.split("|")) == 3,
                "Native Excel version/build/OS metadata was not confirmed.")
        report["nativeEnvironment"] = environment
        require(invoke(host, identifier, "PrepareCells").get("value") is True, "Fixture cells were not prepared.")
        source_before = vba("view", {"moduleName": "ProtectionProbe"})["sourceHash"]
        rejected = vba("protect-project", {"password": "Synthetic\u041f\u0440\u0438\u043c\u0435\u0440!7"}, expected_error=True)
        require("printable ASCII" in json.dumps(rejected), "Lossy Unicode input was not explicitly rejected.")
        require(vba("status")["protection"] == "None", "Rejected Unicode input changed protection.")
        require(vba("view", {"moduleName": "ProtectionProbe"})["sourceHash"] == source_before,
                "Rejected Unicode input changed source.")
        report["lossyUnicodeRejectedBeforeMutation"] = True
        other = create(directory / "Untouched.xlsm")
        configured = vba("protect-project", {"password": password})
        require(configured.get("configurationApplied") is True and configured.get("saveRequired") is True
                and configured.get("reopenRequired") is True, "Pending protection state was not reported.")
        require(password not in json.dumps(configured, ensure_ascii=False), "Response echoed a synthetic password.")
        report["pendingConfiguration"] = True
        require(vba("status", session=other)["protection"] == "None", "Another project was changed.")
        report["otherProjectUnchanged"] = True
        host.close(other)
        other = None
        reopen_protected()
        report["persistedLock"] = True
        attempts = (("wrong-character", password.replace("Exact", "Other")),
                    ("wrong-trimmed", password.strip()),
                    ("wrong-case", password.replace("Exact", "eXact")),
                    ("wrong-escape", password.replace("\\", "/")),
                    ("correct", password))
        for label, submitted in attempts:
            mode = "correct" if label == "correct" else "wrong"
            expected = "None" if mode == "correct" else "Locked"
            require(mode == "correct" or submitted != password, "The negative password must be distinct.")
            authenticate(host, identifier, native, options.auth_probe, directory, submitted, mode, label)
            require(vba("status")["protection"] == expected, "Authentication changed unexpected project state.")
            report[label + "PasswordVerified"] = True
        require(invoke(host, identifier, "HideProbeEditor").get("value") is True, "The synthetic editor did not close.")
        require(vba("view", {"moduleName": "ProtectionProbe"})["sourceHash"] == source_before,
                "VBA source changed during protection/authentication.")
        require(invoke(host, identifier, "CellsPreserved").get("value") is True, "Workbook cells changed.")
        report["sourcePreserved"] = report["cellsPreserved"] = True
        replacement = password.replace("Exact", "Fresh")
        replaced = vba("protect-project", {"password": replacement})
        require(replaced.get("configurationApplied") is True and replaced.get("saveRequired") is True
                and replaced.get("reopenRequired") is True, "Accessible-project password replacement was not configured.")
        reopen_protected()
        authenticate(host, identifier, native, options.auth_probe, directory, password, "wrong", "replaced-old")
        require(vba("status")["protection"] == "Locked", "The old password still authenticated after replacement.")
        authenticate(host, identifier, native, options.auth_probe, directory, replacement, "correct", "replaced-correct")
        require(vba("status")["protection"] == "None", "The new password did not authenticate after replacement.")
        require(invoke(host, identifier, "HideProbeEditor").get("value") is True, "The replacement editor did not close.")
        require(vba("view", {"moduleName": "ProtectionProbe"})["sourceHash"] == source_before
                and invoke(host, identifier, "CellsPreserved").get("value") is True,
                "Password replacement changed VBA source or workbook cells.")
        report["passwordReplacementPersisted"] = True
        report["success"] = True
    except Exception as error:
        report["failureType"] = type(error).__name__
        raise
    finally:
        failures = []
        for session in (other, identifier):
            if session:
                try:
                    host.close(session)
                except Exception as error:
                    failures.append({"phase": "close_session", "type": type(error).__name__})
        try:
            host.shutdown()
        except Exception as error:
            failures.append({"phase": "shutdown", "type": type(error).__name__})
        try:
            if native:
                native.finish()
        except Exception as error:
            failures.append({"phase": "native_exit", "type": type(error).__name__})
        report["cleanupFailures"] = failures
        report["inputDrift"] = [name for name, expected in report["inputHashes"].items()
                                if not Path(name).is_file() or hashlib.sha256(Path(name).read_bytes()).hexdigest() != expected]
        report["success"] = report["success"] and not failures and not report["inputDrift"]
        options.report.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    require(report["success"], "Protection acceptance or owned cleanup failed; see the private receipt.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--entrypoint", choices=("mcp", "cli"), required=True)
    parser.add_argument("--exe", type=Path, required=True)
    parser.add_argument("--auth-probe", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True, help="Receipt under a NEW private directory.")
    run(parser.parse_args())
    print("PASS: persisted protection, exact authentication, source/cell preservation and owned exit.")
