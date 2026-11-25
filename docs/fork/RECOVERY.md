# Verify, upgrade and recover

Use packages from [this fork's releases](https://github.com/Antrakt92/mcp-server-excel/releases). Download the MCP or CLI ZIP together with `SHA256SUMS`, `RELEASE-MANIFEST.json` and `VERIFICATION-RESULTS.json` from the **same release**.

## Verify a download

Run `Get-FileHash <downloaded-file> -Algorithm SHA256` for the ZIP, manifest and verification receipt, and compare each full value with its matching filename in `SHA256SUMS`. Check the manifest's version, upstream base and source revision before extracting. Match the verification receipt's source revision and executable hashes to the manifest and the extracted executables; see [release verification](OVERVIEW.md#distribution). Hashes verify that your download matches the publisher's release; they do not independently establish publisher trust or workbook correctness.

Extract into a new version-specific directory. Do not overwrite a running executable. Keep the previous verified package until the new version has passed the checks relevant to your work.

## Switch versions

1. Inspect active work and the intended Excel windows. Decide whether to save changes explicitly; shutdown is not permission to save or discard them.
2. Stop the relevant MCP connection using its client, or use the CLI's documented service controls. Do not kill all Excel or MCP processes.
3. Point the client configuration or command to the new extracted executable. Restart or reconnect that connection.
4. Check the new executable's version/help and MCP initialization through the intended client. A changed path or hash alone does not prove the client activated the new process.
5. Discover/list the intended workbook and inspect its state before continuing. Avoid opening another copy of a workbook already available to attach.

To roll back, repeat these steps with the previous verified release. Preserve workbook state first. Switching binaries cannot undo workbook edits.

### CLI daemons from earlier versions

The default CLI daemon pipe now includes the Windows session. A daemon already
running on the earlier per-user default keeps its pipe and workbook sessions.
Use the previous CLI to finish that work, or temporarily point the new CLI to
the old pipe from the same Windows session.

```powershell
$sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$env:EXCELMCP_CLI_PIPE = "excelmcp-cli-$sid"
excelcli service status
```

After finishing or handing over that daemon's work, clear
`EXCELMCP_CLI_PIPE` to use the new default. Explicit pipe names still require
the client and daemon to share a Windows session. Use distinct override names
when running custom daemons in separate sessions.

## Migrate automation from fork 2.5.4

Fork 3.0.0 uses upstream 2.6.0 and requires CLI consumer changes:

- Read returned values from `data` and failure details from `error`; the common envelope also includes `command`, `success` and `durationMs`. Batch entries retain their `index`.
- Replace `--sheet` and `--range` with `--sheet-name` and `--range-address`. Remove obsolete result-field selection from saved requests; current action help is authoritative.
- Check formula preflight `validationScope` and `unverifiedChecks`; `isValid` does not prove Excel compilation, UDF availability or safe overwriting.
- Rediscover MCP tools after reconnecting. Tool results retain JSON text and structured content; do not infer success from transport delivery alone.
- Normal opens disable VBA macros. The explicit enable-macros approval and saved-workbook reopening gate applies to both `vba run` and `vba invoke`; audit sessions cannot enable or execute macros.

Finish or explicitly hand over current workbook work with the old version before switching. Test migrated automation on synthetic data before using operational workbooks. See [agent workflows](AGENT-WORKFLOWS.md) for current inspection and VBA limits. Older verification records describe their own release only.

## When an MCP connection closes

`Transport closed` means the client lost its connection to that server.
The message alone gives no outcome for the last operation. If the client can
reconnect the server, use that control. Otherwise start a fresh client session
with the configured server. Verify that `file_read list` answers on the new
connection. A successful response confirms that connection, and its session
list describes only the current controller.

For a local desktop workbook, rediscover the open book before opening its
saved file again. Use MCP `file_read discover` on a working connection, or the
CLI if that is an appropriate entry point for the workbook:

```powershell
excelcli -q session discover
```

For the CLI, inspect `success`, `data.complete`, and `data.errors`. A successful
command can return incomplete discovery. Identify the intended workbook by
location and Excel process, and check `saved` for unsaved edits. If discovery
reports an active owner, a different build, or release still in progress, use
the owning connection or wait for release and rediscover.

Attach only when the task authorizes working on the open workbook. Use the
`workbookId` from the same controller's discovery result, as
`file attach` with `open_workbook_id` in MCP or
`excelcli session attach --workbook-id <ID>` in the CLI. Use the newly returned
session ID for subsequent commands (`data.sessionId` in CLI JSON). Attachment
keeps the workbook's unsaved edits and visibility.

Operation receipts belong to the controller that handled the original
request. A fresh MCP connection or the CLI cannot recover a receipt from a
different controller. If the original receipt is unavailable, inspect the
book's current state and any pending work before retrying a write, save,
or macro. Follow the recovery checks below while the operation is uncertain.

## After cancellation, timeout or a lost connection

Cancellation and timeout stop waiting; Excel's work may continue. Check the Excel window and discover/list sessions before retrying. Do not assume the operation failed, undo partial changes automatically, or close a workbook as cleanup. If the host exits, unsaved workbooks may remain open and visible and attached workbooks stay untouched, following upstream 2.6.0 ownership rules. Pending work can remain in a retained worker; inspect its reported state using the matching build before retrying. Changing binaries does not finish, reverse or authorize cleanup of that work. AutoSave may already have saved changes independently.

After a VBA timeout, look up the receipt if `vba.invoke` was sent with a batch
`operationId`. Check `operationPending` in `file_read list`. Once Excel is ready
again, read back the expected workbook effects before deciding whether to rerun.
Forced cleanup fails a native test run. Use the macro result and workbook
read-back to verify successful execution.

If recovery requires a process-level action, establish the exact owned process identity and current workbook state first. A process name alone does not establish ownership. Unrelated Excel instances must remain untouched.

When a visibility snapshot fails, `partialState.diagnostics.execution` separates
its exception type and HRESULT from the operation's failure. `callbackReturned`
means the callback returned normally; it does not prove saving or verify workbook
contents. `complete:false` means the available chain cannot establish one
unambiguous boundary. Original error fields and observed VBA state remain in
the response. A retained `ExcelDisconnected` session stops later batch commands
even if a partial state reports that an individual callback completed.

Disconnect errors also include `partialState.diagnostics.nativeProcess` when
the execution owner retained a native observation. Check `state` (`Alive`,
`Exited` or `Indeterminate`), `exitCode` and `nativeError` separately. An unknown
exit code is null, not zero. `startedAtUtcFileTime` is an exact decimal string;
keep those digits intact when comparing process identities.

`session.list` includes `recentExcelExits`, capped at 16 exact PID/creation-time
identities. This is diagnostic history held by that controller; it does not
survive a controller restart, recreate sessions or grant ownership of recovery
processes. Exit code zero confirms normal native exit, not saved workbook edits
or the cause of an earlier automation error. An unconfirmed exit never releases
ownership merely because a COM call or Quit returned.

When a close disconnect reports `Alive`, Excel was still running when the error
was handled. If the session then disappears, compare both `processId` and
`startedAtUtcFileTime` with `recentExcelExits` from the same controller. Use
discovery to check for an open workbook and verify the saved file before retrying.

If a known worker dies before acknowledging release, the controller can retain
its newer persisted exit observation after validating both process identities,
controller generation and build. This recovers diagnostic evidence for an
already-known Excel identity; it does not adopt an unseen process or confirm an
unacknowledged workbook operation.

If the session disappears after Excel exits with a nonzero or unavailable exit
code, commands to that session report its last outcome as unknown. An explicit
successful close reply can still confirm the close, including when the reply is
retained inside a status recording error. Check the saved file before repeating
any edits.

`file_read list` and `excelcli session list` also return `recentCloseAttempts`.
Each entry covers an explicitly requested close on a known session and pairs
its ID with the Excel PID and exact creation time when available. The history
keeps up to 16 session/process pairs in the current controller.

- `closeConfirmed` records a successful close response from the session's worker.
- `sessionReleased` records that the controller removed the session and its path claims.
- `saveConfirmed` records an explicit successful save-with-close response.
- `initialSaveConfirmed` means Excel completed the first save requested by
  `close(save=true)`. A later failure can still leave newer edits unsaved and
  the workbook open. A close error returned by the worker includes the same field
  when the first save completed. A lost worker reply can leave this unknown.
- `nativeExit` adds a matching confirmed process exit when available.
- `lastFailure` preserves the latest close error's category, type and HRESULT.
  A later success can add confirmation without erasing that failure.

- `recordingFailures` reports errors publishing the worker or controller result separately.
- `priorStatusDiagnostic` preserves a worker status failure reported before dispatch, even if the close request succeeds.
- `priorSaveBlocker` records the first confirmed dialog observed during the preceding save of that session. It reports the dialog kind (`officeModal` or `vbeModal`) and time, without dialog text. The same observation is available as `operation.firstBlocker` while the save status is retained. A VBE dialog observation does not identify the underlying VBA error.

If a session has `closeReadiness: "unconfirmed"`, a worker status refresh
failed. The row can retain `canClose: false` while the owning worker is still
able to recheck an explicitly requested close. `closeReadiness: "blocked"`
indicates pending work, worker release, or a blocker from the last session status. When available, `statusDiagnostic`
reports the stage, exception type, COM HRESULT and last matching native
process observation. Keep the existing close safeguards and confirm the
original session ID before sending a targeted close.

A false confirmation field means that evidence is unavailable. An observed exit
can appear after a failed close, including a normal exit with code zero. Check the
file on disk before retrying a save. This history is not a recovery handle and
expires when the controller restarts.
### Tracked commands after a lost reply

To retain phase evidence for `vba.invoke`, `session.save` or `table.append`,
send the command as a single batch entry with an `operationId`. Look up the
receipt by the original session ID and UUID on the same controller. The
`operationEvidence` object reports the last confirmed boundary. A failed
command can also return it in `partialState.diagnostics.operationEvidence`.

Use `excelcli -q session operation-result --session <ID> --operation-id <UUID>`
or MCP `file_read` action `operation-result` with `workbook_session_id` and `operation_id`.

- `controllerAccepted` means a receipt exists. The controller may still be
  validating or queuing the request.
- `controllerDispatchStarted` means the controller entered the worker request
  path. It does not confirm delivery or admission.
- `workerAdmitted` means the worker admitted that exact command.
- `serviceResponseRecorded` means the worker recorded a Service response.
  A timeout can reach this stage while native work is pending.

`initialFailure` keeps the first recorded command failure's category, type and
HRESULT when available. The separately bounded `result` retains the initial
reply if the late response is lost.
`workerFailure` contains the recorded failure category, exception type and
HRESULT when available. `nativeExit` can appear later, matched by exact PID and
creation time. An exit can follow later operations or normal cleanup. Its code
alone gives no evidence about which command caused it or what was saved.
An absent `nativeExit` means the controller has no matching observation yet.
Read `session list` or `file_read list`, then check the receipt again after the
controller observes the exit. Use `recentCloseAttempts` for a subsequent close
error. A close request after the controller has released a lost session can
report the process exit while both `saveConfirmed` and `closeConfirmed` stay false.

For a tracked `vba.invoke`, `requestedProcedure` is the unqualified selector
provided by the caller, when it can be represented as a bounded VBA identifier
or dotted identifier. It identifies the request. It does not show which VBA
statement or nested procedure was running when COM disconnected.

The selector is kept in the controller receipt, outside the worker checkpoint.
The worker checkpoint carries no workbook paths, VBA source, inputs or result
payloads. Each worker keeps its latest checkpoint, so an earlier phase may be
unavailable if the checkpoint was replaced before the controller observed it.
The controller receipt remains bounded and process-local. Use separate VBA
phase marks for individual operations inside a macro when debugging a fixture.

### Match a Windows crash report

For a confirmed Excel process exit, inspect the local Windows **Application**
log around the incident for source **Application Error**, event **1000**. Match
`AppName` (`EXCEL.EXE`), `ProcessId`, and `ProcessCreationTime` in the event XML
to the receipt's process identity. The event's PID and creation time may be in
hexadecimal. Compare them numerically with the decimal `processId` and
`startedAtUtcFileTime`. PID alone can match a recycled process.

If there is an exact match, record the event time, `ExceptionCode`, and
`IntegratorReportId` when present. A missing or delayed event leaves the report
unconfirmed. Keep the event details local and omit raw messages and paths from
shared diagnostics. See [Microsoft's Application Error guidance](https://learn.microsoft.com/en-us/troubleshoot/windows-server/performance/troubleshoot-application-service-crashing-behavior).

## Rebuild or report a problem

The separately labeled [historical 2.4.1 patch](LEGACY.md) preserves the earlier
development line. It is archival source and must not be applied to current releases.

Use the tagged source archive or Git checkout matching the release manifest. Follow [the build and verification boundary](OVERVIEW.md#build-and-verify) and `scripts/Build-ForkRelease.ps1`. Preserve newly generated checksums and manifest; a rebuild can differ from the published binary and needs its own verification.

For a fork issue, include the version, package checksum, Windows/Excel versions, entry point and minimal reproduction. Describe observed state and which checks ran. Remove workbook contents, credentials, user-specific paths and private logs before publishing.
