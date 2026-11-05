# Fork 4.1.0 verification

This release adds failure diagnostics and controller exit history, and stops
dependent CLI/MCP batch work after an Excel disconnect. Existing 4.0.0 inputs,
the upstream 2.6.11 base and dependency versions are unchanged.

## Source and regression checks

Build the complete Release solution with SDK 10.0.403 and the reviewed ADO
importer. Run the affected Excel-free suites, generated contracts, packaging
checks and worker/MCP process-boundary tests. Keep source commit/tree and actual
results with the release evidence; a prior release's counts are historical.

Require the following regressions:

- CLI and MCP batches stop on `ExcelDisconnected`, even with ordinary
  continue-on-error behavior or a completed individual callback.
- Operation and visibility-snapshot errors remain distinguishable. Preserve
  cancellation, both causes of a combined failure and observed VBA state.
  Bounded exception traversal must report ambiguous/incomplete context honestly.
- Native observation requires a retained, identity-checked handle. A signaled
  handle and a successful exit-code read are separate facts; failed reads keep
  an explicit null code. Cover unsigned codes, concurrent observation/disposal,
  retryable teardown and preservation of a still-live neighboring workbook.
- Persist an exit observation before releasing ownership. An identical replay
  needs no second write; richer evidence still needs successful persistence.
  Recover a newer durable observation from a dead known worker before retirement,
  rejecting mismatched identities/build/generation without ownership callbacks.
- Keep at most 16 exact identities in `recentExcelExits`, including partial
  session-list responses. Preserve creation FILETIME digits as decimal strings.
- Raw MCP discovery and direct/batch invocation preserve native scalar types.
  The client presentation workaround is documentation, not a server-schema fix.
- Locked VBA source errors name the supported known-password unlock route.

Retain failing-before and passing-after evidence for behavior repairs. Run the
real-Excel lifecycle/diagnostic cases against the frozen matching build. Assert
actual retained cell edits and save state after errors, native exit separately
from COM cleanup, and no successful save claim after Excel has already exited.

## Workbook and synthetic checks

Keep operational workbooks private. Inspect and test a separately owned copy;
verify original and on-disk copy hashes after no-save cleanup. The real-copy
smoke covers known-password access, reviewed source inspection and a typed scalar
return. It does not validate every workbook workflow, calculation or print page.

The synthetic modeless/modal profile checks activation, closure, a timer callback,
confirmed cancellation and event draining while its module remains loaded. Bind
the timer to the creating process/UI thread. Require the matching timer ID and
`WM_TIMER`; validate any nonzero callback HWND against that same process/thread.
Do not assume a forwarded timer message keeps a null HWND.

Compare two calls in one Excel lifetime with one call in each of two fresh
lifetimes. Require typed Boolean results, exact owner/readiness checks, unchanged
files and normal Excel/worker/controller exit. This bounded comparison is not a
statistical stability estimate or a reproduction of the historical RPC failures.
Preserve failed diagnostic attempts separately; cleanup cannot turn them into
passing tests. Unknown callback cleanup keeps the test quarantined.

The historical RPC cause remains unresolved. Collect the operation/snapshot
diagnostics and retained native exit evidence if it recurs. Recovery-child
ownership still requires independent proof; exit history grants no cleanup rights.

## Packaged acceptance and publication

From the final clean merged source, build both portable Windows x64 packages:

```powershell
pwsh -NoProfile -File scripts/Build-ForkRelease.ps1 -Version 4.1.0
```

Verify manifest source identity, SHA-256 hashes, PE architecture and executable
versions. Against the extracted executables, run:

- `scripts/tests/fork_protocol_smoke.py` with `--idle-seconds 15`: protocol,
  negative parity, EOF, broken stdout and parent-exit boundaries.
- `scripts/tests/fork_packaged_excel_smoke.py`: workbook and lifecycle acceptance
  through both MCP and CLI.
- `scripts/tests/fork_vba_authentication_smoke.py`: public password actions,
  recognized rejection, accepted access, stdin-only CLI input, source/cell
  persistence after explicit save/reopen, and normal owned cleanup.

Publish two ZIPs, `RELEASE-MANIFEST.json`, `VERIFICATION-RESULTS.json` and
`SHA256SUMS` through the authenticated fork CLI procedure. Verification receipts
must describe the actual packaged bytes and distinguish current checks from
unchanged 4.0.0 screenshot coverage. Private books, paths and incident records
must stay outside release assets. Download and compare all five published assets.

After installation, verify the configured CLI and a fresh configured MCP process.
An existing client's activation is a separate reconnect check. Inspect pending
work before retiring superseded hosts and recycling the old installation.
