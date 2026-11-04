"""Verify known-password VBA authentication through public MCP and CLI entry points.

Windows/desktop Excel only. Creates owned synthetic XLSM fixtures; never attaches,
executes VBA, changes macro policy, or uses a test-only authentication driver.
Passwords exist only in memory and public MCP input / CLI batch stdin. Receipts
contain binary identities and assertions, never passwords, source or cell values.
Run only after reviewing this script and freezing the matching product closure.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import secrets
import shutil
import sys
import tempfile

sys.dont_write_bytecode = True

MODULE = "AuthenticationSmoke"
SOURCE = "Option Explicit\nPublic Const AuthenticationMarker As Long = 7\n"
EDITED_SOURCE = "Option Explicit\nPublic Const AuthenticationMarker As Long = 11\n"
CELLS = [["synthetic", 7], [True, "original"]]
EDITED_CELLS = [["synthetic", 11], [False, "persisted"]]


class SmokeError(RuntimeError):
    pass


def require(condition, message):
    if not condition:
        raise SmokeError(message)


def load_helper(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, "Reviewed smoke helper was not found.")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def no_echo(value, passwords):
    encoded = (json.dumps(value, ensure_ascii=False), json.dumps(value, ensure_ascii=True))
    for password in passwords:
        for candidate in (password, password.strip()):
            for text in encoded:
                require(candidate not in text and json.dumps(candidate)[1:-1] not in text,
                        "A public response exposed synthetic password text; raw response withheld.")


def call(host, entry_point, command, arguments=None, identifier=None, expected_error=False, passwords=()):
    """Only native public spellings cross the boundary; CLI credentials always use CliHost.call stdin."""
    arguments = arguments or {}
    if entry_point == "cli":
        value = host.call(command, arguments, identifier, expected_error)
    else:
        category, action = command.split(".", 1)
        if category == "session":
            tool = "file_read" if action == "list" else "file"
        elif category == "vba":
            tool = "vba_read" if action in ("view", "status") else "vba"
        elif category == "range":
            tool = "range_read" if action == "get-values" else "range"
        else:
            raise SmokeError("The smoke requested an unsupported public command mapping.")
        names = {"filePath": "file_path", "moduleName": "module_name", "vbaCode": "vba_code",
                 "expectedSourceHash": "expected_source_hash", "sheetName": "sheet_name",
                 "rangeAddress": "range_address", "overwritePolicy": "overwrite_policy"}
        payload = {names.get(key, key): item for key, item in arguments.items()}
        payload["action"] = action
        if identifier is not None:
            payload["workbook_session_id"] = identifier
        value = host.call(tool, payload, expected_error)
    no_echo(value, passwords)
    return value


def auth_result(value, project, verified, already_accessible=False):
    expected = {"projectName": project, "protectionBefore": "None" if already_accessible else "Locked",
                "protectionAfter": "Locked" if verified is False else "None",
                "authenticationPerformed": not already_accessible, "alreadyAccessible": already_accessible,
                "passwordVerified": verified, "accessUnlocked": verified is True,
                "protectionSettingsChanged": False, "saved": False, "reopened": False}
    for name, expected_value in expected.items():
        require(name in value and type(value[name]) is type(expected_value) and value[name] == expected_value,
                "Authentication result field did not match the exact expected state.")


def auth_error(value, category, performed, verified, accessible, protection):
    actual_category = value.get("errorCategory")
    if isinstance(value.get("error"), dict):
        actual_category = value["error"].get("category")
    require(actual_category == category, "Authentication error category differed from the expected state.")
    state = value.get("partialState")
    require(isinstance(state, dict) and isinstance(state.get("phase"), str) and state["phase"],
            "Authentication failure omitted its observed partial state.")
    expected = {"challengeObserved": not accessible, "authenticationPerformed": performed,
                "passwordVerified": verified, "accessUnlocked": False, "alreadyAccessible": accessible,
                "currentProtection": protection, "dialogCleanupConfirmed": True,
                "activeProjectRestored": True, "visibilityRestored": True,
                "protectionSettingsChanged": False, "saved": False, "reopened": False,
                "rollbackAttempted": False}
    for name, expected_value in expected.items():
        require(name in state and type(state[name]) is type(expected_value) and state[name] == expected_value,
                "Authentication error lost exact outcome or cleanup state.")


def normalized_code(value):
    require(isinstance(value, str), "The module read omitted source text.")
    return value.replace("\r\n", "\n").replace("\r", "\n").strip("\n")


def source_read(host, entry_point, identifier, expected, passwords):
    value = call(host, entry_point, "vba.view", {"moduleName": MODULE}, identifier, passwords=passwords)
    require(value.get("moduleName") == MODULE and normalized_code(value.get("code")) == normalized_code(expected),
            "Public module inspection did not return the exact expected synthetic source.")
    fingerprint = value.get("sourceHash")
    require(isinstance(fingerprint, str) and re.fullmatch(r"[0-9a-fA-F]{64}", fingerprint) is not None,
            "Public module inspection omitted its complete source fingerprint.")
    return fingerprint


def cell_read(host, entry_point, identifier, expected, passwords):
    value = call(host, entry_point, "range.get-values",
                 {"sheetName": "Sheet1", "rangeAddress": "A1:B2"}, identifier, passwords=passwords)
    actual = value.get("values")
    require(isinstance(actual, list) and len(actual) == len(expected), "The cell read changed its shape.")
    for row, target in zip(actual, expected):
        require(isinstance(row, list) and len(row) == len(target), "The cell read changed its shape.")
        for item, expected_item in zip(row, target):
            same_type = type(item) in (int, float) if type(expected_item) in (int, float) else type(item) is type(expected_item)
            require(same_type and item == expected_item,
                    "The public cell read did not preserve the exact expected synthetic values.")


def close_for_fresh_open(host, identifier):
    host.native.capture(required=True)
    host.close(identifier)
    # Retained handles prove actual Excel exit; a new session identifier alone is insufficient.
    excel = [handle for (role, _, _), (handle, _) in host.native.handles.items() if role == "excel"]
    require(excel and all(handle.wait(20) for handle in excel),
            "The previous owned Excel instance did not exit before the fresh-lock check.")


def run_case(host, entry_point, path, benchmark):
    password = ' A"' + secrets.token_hex(4) + '\\!7 '
    wrong = password.replace(' A"', ' a"', 1)
    passwords = (password, wrong)
    host.report["phase"] = "create-synthetic-fixture"
    created = call(host, entry_point, "session.create", {"filePath": str(path)}, passwords=passwords)
    identifier = created[host.id_key]
    host.native.capture(required=True)
    call(host, entry_point, "vba.import", {"moduleName": MODULE, "vbaCode": SOURCE}, identifier, passwords=passwords)
    call(host, entry_point, "range.set-values", {"sheetName": "Sheet1", "rangeAddress": "A1:B2", "values": CELLS},
         identifier, passwords=passwords)
    status = call(host, entry_point, "vba.status", identifier=identifier, passwords=passwords)
    project = status.get("projectName")
    require(status.get("projectAccess") is True and status.get("protection") == "None"
            and isinstance(project, str) and project, "Synthetic project access was not established.")
    original_hash = source_read(host, entry_point, identifier, SOURCE, passwords)
    cell_read(host, entry_point, identifier, CELLS, passwords)

    host.report["phase"] = "configure-and-persist-lock"
    configured = call(host, entry_point, "vba.protect-project", {"password": password}, identifier, passwords=passwords)
    require(configured.get("projectName") == project and configured.get("configurationApplied") is True
            and configured.get("currentProtection") == "None" and configured.get("saveRequired") is True
            and configured.get("reopenRequired") is True, "Project protection configuration was not confirmed.")
    call(host, entry_point, "session.save", identifier=identifier, passwords=passwords)
    close_for_fresh_open(host, identifier)
    saved_before_auth = benchmark.file_identity(path)["sha256"]
    opened = call(host, entry_point, "session.open", {"filePath": str(path)}, passwords=passwords)
    identifier = opened[host.id_key]
    host.native.capture(required=True)
    status = call(host, entry_point, "vba.status", identifier=identifier, passwords=passwords)
    require(status.get("projectName") == project and status.get("protection") == "Locked",
            "The freshly opened project did not retain its saved viewing lock.")
    host.report["persisted_lock_confirmed"] = True

    host.report["phase"] = "wrong-password-public-results"
    rejected = call(host, entry_point, "vba.verify-password", {"password": wrong}, identifier, passwords=passwords)
    auth_result(rejected, project, False)
    failure = call(host, entry_point, "vba.unlock-project", {"password": wrong}, identifier, True, passwords)
    auth_error(failure, "Permissions", True, False, False, "Locked")
    require(call(host, entry_point, "vba.status", identifier=identifier, passwords=passwords).get("protection") == "Locked",
            "A wrong password changed project access.")
    host.report["wrong_password_verification_and_unlock_confirmed"] = True

    # A nonempty all-space value must reach the native challenge unchanged,
    # rather than being rejected as a missing argument by an adapter.
    whitespace_rejected = call(host, entry_point, "vba.verify-password", {"password": "   "},
                               identifier, passwords=passwords)
    auth_result(whitespace_rejected, project, False)
    whitespace_failure = call(host, entry_point, "vba.unlock-project", {"password": "   "},
                              identifier, True, passwords)
    auth_error(whitespace_failure, "Permissions", True, False, False, "Locked")
    host.report["whitespace_password_rejection_confirmed"] = True

    host.report["phase"] = "verify-correct-and-already-accessible"
    accepted = call(host, entry_point, "vba.verify-password", {"password": password}, identifier, passwords=passwords)
    auth_result(accepted, project, True)
    failure = call(host, entry_point, "vba.verify-password", {"password": wrong}, identifier, True, passwords)
    auth_error(failure, "Prerequisite", False, None, True, "None")
    accessible = call(host, entry_point, "vba.unlock-project", {"password": wrong}, identifier, passwords=passwords)
    auth_result(accessible, project, None, already_accessible=True)
    require(benchmark.file_identity(path)["sha256"] == saved_before_auth,
            "Authentication changed the saved workbook without an explicit save.")
    require(source_read(host, entry_point, identifier, SOURCE, passwords) == original_hash,
            "Authentication changed the original module source fingerprint.")
    cell_read(host, entry_point, identifier, CELLS, passwords)
    host.report["correct_verify_and_accessible_semantics_confirmed"] = True

    host.report["phase"] = "guarded-source-and-cell-edit"
    call(host, entry_point, "vba.replace-module", {"moduleName": MODULE, "expectedSourceHash": original_hash,
         "vbaCode": EDITED_SOURCE}, identifier, passwords=passwords)
    edited_hash = source_read(host, entry_point, identifier, EDITED_SOURCE, passwords)
    require(edited_hash != original_hash, "Guarded source replacement did not change its fingerprint.")
    call(host, entry_point, "range.set-values", {"sheetName": "Sheet1", "rangeAddress": "A1:B2",
         "values": EDITED_CELLS, "overwritePolicy": "allow"}, identifier, passwords=passwords)
    cell_read(host, entry_point, identifier, EDITED_CELLS, passwords)
    call(host, entry_point, "session.save", identifier=identifier, passwords=passwords)
    close_for_fresh_open(host, identifier)
    persisted_hash = benchmark.file_identity(path)["sha256"]
    identifier = call(host, entry_point, "session.open", {"filePath": str(path)}, passwords=passwords)[host.id_key]
    host.native.capture(required=True)
    require(call(host, entry_point, "vba.status", identifier=identifier, passwords=passwords).get("protection") == "Locked",
            "Source/cell edits removed the saved viewing lock.")
    host.report["phase"] = "fresh-unlock-and-persisted-readback"
    auth_result(call(host, entry_point, "vba.unlock-project", {"password": password}, identifier, passwords=passwords),
                project, True)
    require(source_read(host, entry_point, identifier, EDITED_SOURCE, passwords) == edited_hash,
            "Freshly reopened module fingerprint differed from the saved edit.")
    cell_read(host, entry_point, identifier, EDITED_CELLS, passwords)
    require(benchmark.file_identity(path)["sha256"] == persisted_hash,
            "Fresh authentication/readback changed the saved workbook.")
    host.report.update(persisted_source_and_cells_confirmed=True, protection_retained_after_edit=True,
                       response_no_password_echo=True, macros_executed=False, macro_policy_changed=False,
                       saved_fixture_sha256=persisted_hash, phase="complete")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--mcp-exe", type=Path, required=True)
    parser.add_argument("--cli-exe", type=Path, required=True)
    parser.add_argument("--source-revision", required=True)
    parser.add_argument("--report-directory", type=Path, required=True,
                        help="New, private, absolute directory; existing directories are refused.")
    args = parser.parse_args()
    require(args.report_directory.is_absolute() and not args.report_directory.exists(),
            "Use a new absolute private report directory.")
    args.report_directory.mkdir(parents=True, exist_ok=False)
    report_path = args.report_directory / "verification.json"
    report = {"success": False, "scope": "Owned synthetic XLSM; public MCP and CLI known-password authentication",
              "cases": [], "stderr_no_echo_scope": "Product CLI regression gate; raw stderr is not retained by this harness."}
    root = None
    benchmark = None
    before = None
    executables = {}
    helpers = {}
    try:
        require(os.name == "nt", "Native public-entrypoint acceptance requires Windows.")
        require(re.fullmatch(r"[0-9a-fA-F]{40}", args.source_revision) is not None,
                "Supply the exact frozen source revision.")
        directory = Path(__file__).resolve().parent
        helpers = {"packaged": directory / "fork_packaged_excel_smoke.py",
                   "benchmark": directory / "fork_cli_batch_benchmark.py", "protocol": directory / "fork_protocol_smoke.py",
                   "authentication": Path(__file__).resolve()}
        benchmark = load_helper("authentication_benchmark", helpers["benchmark"])
        packaged = load_helper("authentication_packaged", helpers["packaged"])
        protocol = load_helper("authentication_protocol", helpers["protocol"])
        executables = {"mcp": args.mcp_exe.resolve(strict=True), "cli": args.cli_exe.resolve(strict=True)}
        before = {"binaries": {name: benchmark.binary_identity(path) for name, path in executables.items()},
                  "helpers": {name: benchmark.file_identity(path) for name, path in helpers.items()}}
        report["provenance"] = {"supplied_source_revision": args.source_revision, "before": before}
        inputs = set(helpers.values()) | set(executables.values())
        for executable in executables.values():
            inputs.update(executable.parent.glob("*.dll"))
            inputs.update(executable.parent.glob("*.json"))
        report["inputHashes"] = {str(path.resolve()): benchmark.file_identity(path)["sha256"]
                                 for path in sorted(inputs, key=lambda item: str(item).lower())}
        snapshots = args.report_directory / "inputs"
        snapshots.mkdir()
        report["inputSnapshots"] = {}
        for index, (name, expected) in enumerate(report["inputHashes"].items()):
            snapshot = snapshots / (str(index).zfill(3) + "-" + Path(name).name)
            shutil.copyfile(name, snapshot)
            require(benchmark.file_identity(snapshot)["sha256"] == expected,
                    "An input changed while preserving its native acceptance snapshot.")
            snapshot.chmod(0o444)
            report["inputSnapshots"][name] = snapshot.relative_to(args.report_directory).as_posix()
        (args.report_directory / "inputs-before.json").write_text(json.dumps(before, indent=2) + "\n", encoding="utf-8")
        report_path.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
        root = Path(tempfile.mkdtemp(prefix="excelmcp-packaged-native-"))
        for name in ("mcp", "cli"):
            case = {"entry_point": name, "passed": False}
            report["cases"].append(case)
            host = packaged.McpHost(executables[name], protocol, case) if name == "mcp" else packaged.CliHost(executables[name], benchmark, case)
            host.native = None
            path = root / (name + ".xlsm")
            try:
                host.start()
                if name == "mcp":
                    case["owned_host"] = benchmark.owned_process_identity(host.process)
                host.native = packaged.NativeOwners(case["owned_host"], case)
                run_case(host, name, path, benchmark)
                case["passed"] = True
            except Exception as error:
                case["failure_type"] = type(error).__name__
            finally:
                failures = []
                try:
                    if host.native is None and host.process is not None:
                        case.setdefault("owned_host", benchmark.owned_process_identity(host.process))
                        host.native = packaged.NativeOwners(case["owned_host"], case)
                    if host.native is not None:
                        host.native.capture()
                except Exception as error:
                    failures.append({"phase": "capture", "type": type(error).__name__})
                for phase, operation in (
                    ("close_fixture", lambda: packaged.cleanup_sessions(host, path, case)
                     if host.process is not None and host.process.poll() is None else None),
                    ("shutdown_host", host.shutdown),
                    ("native_exit", lambda: host.native.finish() if host.native is not None else
                     require(False, "Native ownership was not initialized."))):
                    try:
                        operation()
                    except Exception as error:
                        failures.append({"phase": phase, "type": type(error).__name__})
                forced = case.get("forced_owned_host_termination") or case.get("host_cleanup", {}).get("forced_termination")
                if failures or forced:
                    case["cleanup_failures"] = failures
                    case["passed"] = False
            report_path.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
            if not (case.get("fixture_sessions_closed") and case.get("host_exit_confirmed")
                    and case.get("native_exit_confirmed")):
                break
        report["success"] = len(report["cases"]) == 2 and all(
            case["passed"] and case.get("fixture_sessions_closed") and case.get("host_exit_confirmed")
            and case.get("native_exit_confirmed") for case in report["cases"])
    except Exception as error:
        report["failure_type"] = type(error).__name__
    finally:
        if before is not None and benchmark is not None:
            try:
                after = {"binaries": {name: benchmark.binary_identity(path) for name, path in executables.items()},
                         "helpers": {name: benchmark.file_identity(path) for name, path in helpers.items()}}
                report["provenance"]["after"] = after
                report["inputDrift"] = [name for name, expected in report.get("inputHashes", {}).items()
                                        if not Path(name).is_file() or benchmark.file_identity(Path(name))["sha256"] != expected]
                report["snapshotDrift"] = [relative for name, relative in report.get("inputSnapshots", {}).items()
                                           if not (args.report_directory / relative).is_file()
                                           or hashlib.sha256((args.report_directory / relative).read_bytes()).hexdigest()
                                           != report["inputHashes"][name]]
                report["inputs_unchanged"] = before == after and not report["inputDrift"] and not report["snapshotDrift"]
                report["success"] = report["success"] and report["inputs_unchanged"]
                (args.report_directory / "inputs-after.json").write_text(json.dumps(after, indent=2) + "\n", encoding="utf-8")
            except Exception as error:
                report["success"] = False
                report["drift_check_failure_type"] = type(error).__name__
        if root is not None:
            packaged.cleanup_fixtures(root, report)
        report_path.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"success": report["success"], "cases": len(report["cases"])}))
    return 0 if report["success"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
