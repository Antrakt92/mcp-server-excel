# Automation & Advanced Features

Extend workbook workflows with code, live window control, What-If Analysis,
and mapped XML data.

[Back to the feature overview](../../FEATURES.md)

Use CLI help or MCP tool descriptions for current actions and inputs. The
summaries below focus on capabilities and their important requirements.

---

## VBA Macros (16 operations)

- **Inspect code:** List procedures with their kinds and line numbers, read a selected procedure or up to 500 lines, or view the full module.
- **Maintain code:** Import standard modules, update whole modules, replace one procedure only when its source has not changed, or remove components.
- **Run procedures:** Execute existing VBA procedures with supported parameters.
- **Typed invocation:** Invoke with typed scalar arguments and read scalar or rectangular-array function results.
- **Guarded module replacement:** Replace source only when its current fingerprint matches; failures retain observed partial state.
- **Find source:** Search literal text across modules with limited matches, line numbers, and short excerpts.
- **Diagnose project access:** Inspect actual access, password protection, and execution mode without changing settings.
- **Inspect dependencies:** List library references and flag broken ones without automatic repair.
- **Configure project protection:** Set a new password on an accessible, unlocked VBA project through its own Project Properties dialog. Saving and reopening remain explicit operations.
- **Authenticate a known password:** Unlock current-session project access, or check a password against a real locked-project challenge. Both actions leave protection settings untouched.

Listing or editing the VBA project requires Excel's manually configured trust
setting. Running an existing macro does not require project-inspection trust,
but Excel's macro security still applies. ExcelMcp does not enable trust or
bypass security automatically. Retain a macro-enabled file format when saving code.
Normal opens disable VBA macros, including `.xlsm` files. Enabling execution
requires separately approving a saved-only reopen; opening event code can run.
Replacing procedure source confirms only that Excel stored the text; it does not
confirm that the code compiles or runs correctly.
Procedure replacement preserves introductory comments and surrounding blank lines.
The status action reports blocked project access rather than guessing protection
or execution mode when access is denied.

### Project password protection

`vba.protect-project` configures locking for the selected, currently accessible
VBA project. It rejects an already locked project; it does not authenticate an
existing password, unlock a project, or remove protection. Workbook structure
protection is a separate operation.
This is the VBA editor's viewing lock, not workbook encryption. The workbook
must contain VBA code or a custom form for Excel to retain the password.

Passwords must use printable ASCII characters. VBA hashes passwords using the
project's legacy code page, which Excel's public object model does not expose.
Distinct Unicode passwords can therefore be accepted as the same password.
ExcelMcp rejects non-ASCII input before changing the project, rather than
silently substituting characters. See Microsoft's
[VBA password hash format](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/4370bcd9-1664-4428-a4c6-fca56aa7d7f1).

The action does not save or reopen the workbook. A successful dialog interaction
is not proof of persisted protection: explicitly save, close and reopen, then
check `vba.status` for `Locked`. The current editor can remain unlocked until that
reopen. Preserve unsaved work when choosing when to close.

For CLI use, the password actions require `excelcli batch --input -` with
`args.password` in stdin. Direct password options and batch input files are
refused. Migrate earlier `vba protect-project --password ...` calls to stdin.
Never include passwords in shell history, logs, session metadata or published examples.

The driver requires an interactive Windows desktop and Excel's standard VBA
Project Properties dialog. It briefly shows the editor, selects and verifies
the requested project, then restores the previous project and editor visibility
after confirmed dialog cleanup. Existing dialogs must be closed before calling.
It does not change Trust Center settings or enter an existing project's password.

Failures report the observed phase, whether acceptance was requested, and which
cleanup/restoration steps completed. An unknown outcome is not a rollback or
permission to repeat the request. Inspect that state before further changes.

### Known-password authentication

`vba.unlock-project` submits the supplied password to the selected locked
project's native challenge. `vba.verify-password` checks the same challenge;
successful verification also makes the project accessible in the current VBE
session. Both require a normal writable session opened by this host. Audit,
read-only and attached sessions are refused before dialog input. Neither action
enables macros, changes Trust Center settings or tries passwords other than the
one supplied. Printable ASCII input is transmitted exactly, including spaces,
quotes and backslashes. Unsupported characters and native length overflow are
refused without substituting or truncating the password.

For a recognized incorrect password, verification completes with
`passwordVerified: false`; unlocking fails with `Permissions` and reports the
observed locked state. If UI, cancellation or a timeout prevents observing an
outcome, the error retains `passwordVerified: null`, never a guessed mismatch. If the project is
already accessible, unlocking returns `alreadyAccessible: true`,
`authenticationPerformed: false` and `passwordVerified: null`. Verification
instead requires a locked project and returns `Prerequisite`; it does not
automatically close or reopen a workbook to create a challenge.

Authentication cancels the target Properties dialog without changing its lock
checkbox or password settings. It does not save, close, reopen or remove the
saved viewing lock. A later fresh open can therefore be locked again. Cleanup
and restoration failures preserve observed authentication state; cancellation
does not undo access already granted by Excel.

Known-password authentication is included in fork 4.0.0. See the release verification receipt for the tested native scope.

[VBA walkthrough](../guides/RUN-VBA-MACROS.md)

---

## Python in Excel (2 operations)

- **Add Python calculations:** Write Python-in-Excel formulas that can reference live worksheet data.
- **Read results:** Wait for cloud execution and read supported calculated values, distinguishing pending work from unavailable Python support.

This is Microsoft's licensed Python in Excel service, not local Python
execution. It requires a supported Microsoft 365 account and internet access.
Rich Python objects are not necessarily readable as ordinary values through
Excel automation; choose worksheet-value output when that is what the task needs.

A successful formula write does not mean the cloud calculation has finished.
An unavailable service and pending calculation need different recovery choices.

[Cloud calculation and failure guidance](../reference/behavioral-rules.md#python-in-excel)

---

## Window Management (17 operations)

- **Watch work live:** Show, hide, position, and arrange the session's Excel window.
- **Inspect context:** Read its active worksheet, selection, chart, window state, and view.
- **Improve navigation:** Adjust zoom, display options, frozen panes, and movable splits.
- **Prepare the opening view:** Set a sheet's scroll position and optional selection,
  check the actual view, then explicitly save the workbook.
- **Show progress:** Set and clear Excel status-bar text.

Preserve the user's existing visibility preference. Keeping a workbook open
does not mean showing its window. Window changes affect the selected session,
not unrelated Excel processes, and visible interaction needs a suitable desktop.

[Window workflow guidance](../reference/window.md)

---

## What-If Analysis (8 operations)

- **Find an input:** Use Goal Seek to adjust one input until a formula reaches a numeric target.
- **Compare assumptions:** Maintain named scenarios, apply their input values, and produce scenario summaries.
- **Explore sensitivity:** Build one- or two-variable Excel data tables.

Goal Seek and applying scenarios change live cells; they are not read-only
inspection. Read the resulting inputs and outputs. Data tables can be
calculation-intensive and need the worksheet layout prepared first.

Solver is not exposed here. It is an optional VBA add-in that requires separate
user configuration and macro-security decisions.

[What-If workflow guidance](../reference/analysis.md)

---

## XML Maps (6 operations)

- **Map structured data:** Manage workbook XML schemas and bind worksheet cells or columns to XML paths.
- **Exchange data:** Import and export mapped XML without dialogs.

Use existing mappings when they describe the intended data. Imports can write
worksheet cells or create a new mapped Table. Removing a map does not clear
previously imported cells.

External schema dependencies, DTDs, and schema-location attributes are rejected
so Excel cannot resolve unexpected external resources.

[XML workflow guidance](../reference/xmlmap.md)

---

## Batch jobs (1 operations)

- **Run known sequences:** Run commands in order in one call, with numbered results and optional stopping after ordinary errors; always stop when completion is uncertain.

Use MCP `batch` action `run` or
`excelcli batch`. Both accept the same command entries and continue after
errors by default, with an option to stop. Each result identifies its input
position.
Timeouts, cancellation, and reported unknown completion always stop further
commands because Excel work may still be running.

Batches do not roll back after failure or automatically undo or save changes.
Individual commands can save or discard edits.

[Batch inputs, results, and session handling](../reference/workflows.md#cli-batch-jobs)

---

## Related feature areas

- [Data & analytics](DATA-ANALYTICS.md) - combine code with queries and analytical models
- [Cells & workbooks](CELLS-WORKBOOKS.md) - manage the inputs, formulas, and files used by automation
- [Charts & visualization](CHARTS-VISUALS.md) - present automated results
- [Real Excel automation vs. file-parser libraries](../guides/EXCEL-COM-VS-FILE-PARSERS.md)
