# Fork 4.0.0 verification

This major release adds known-password VBA authentication and fixes range
screenshots in normal, frozen and split views, including stale viewport images.
Existing CLI password workflows must migrate to batch JSON on stdin. Upstream remains 2.6.11 and MCP SDK 2.3.1;
no dependencies change.

## Source checks

Use clean committed source, SDK 10.0.403 and the reviewed ADO importer. Build
Release with zero warnings/errors. Run the affected Excel-free groups, generated
contract checks and fork packaging tests. Bind receipts to the tested source
commit/tree. Keep failed or interrupted runs separately; a later pass does not
erase them or prove an unexplained failure's cause.

Require focused generated-contract, CLI stdin and MCP protocol regressions for
one-space and several-space passwords on all three password actions. Assert
exact strings, one dispatch for valid inputs, and rejection of missing, null,
empty and non-string passwords before dispatch. Ordinary required string inputs
must continue to reject whitespace. Retain the failing-before and passing-after
receipts for this validation change.

Run the 45-case real-Excel Service VBA suite, including
`PersistentServiceVbaProtectionTests` and `PersistentServiceVbaAuthenticationTests`.
Require actual fresh-open locking, wrong-password rejection and exact-password
acceptance, already-accessible semantics, audit refusals and overlong-input
refusal. Verify source/cell preservation, unchanged sibling project and saved
bytes, restored editor context, and source edits that remain present after an
explicit save and fresh open with protection still locked. Require normal owned
Excel/worker/controller shutdown.

Run the normal-close exit-confirmation regression and existing late-exit lease
checks. An injected unconfirmed native exit must fail close, retain the exact
process identity and quarantine the session. After actual exit, another close
must reconcile tracking without attempting another workbook save or close.
Exercise affected lifecycle and ownership tests as well as VBA and screenshots.
Retain the earlier intermittent delayed-shutdown incident separately: successful
diagnostic runs do not explain its cause, and the ownership fix does not claim
to eliminate every delay in Excel shutdown.

Verify attached unsaved-workbook discovery and first Save As. Names such as
`Book1` must not enter file reservations. Check unsaved/saved location
transitions, unrelated reservations, saved-file readback after reopening, and
preserved neighboring workbook content and save state. Run the affected
incomplete-discovery cleanup and running attached-callback shutdown cases.

Run all 46 screenshot Service cases against the final clean merged release
source. Retain matched-scene evidence for stale viewport pixels after scrolling
or tiling. Every capture path uses a standard preparatory render followed by
the full-content final image. Verify ordinary single-pane captures, including
fractional zoom, as well as frozen and split panes. Check offscreen styled
cells, embedded charts and each newly scrolled tile against expected current
worksheet content at the same coordinates. Preserve existing pane, zoom,
selection, visibility, window-state and workbook Saved-state invariants.
Run the complete screenshot and VBA source suites once
against that final source, rather than repeating the complete suites at each
preparation step. Source changes invalidate affected evidence. Verify corner markers and complete worksheet pixels across asymmetric frozen panes and split views, varied row sizes, and
hidden/minimized/maximized windows. Check complete interior row-color bands as
well as corner markers, so an older viewport cannot pass through marker presence
alone. Include a frozen target worksheet that was not active before capture;
verify the target and previously active sheet independently, without a test
selection scrolling the view back into place. Require preserved pane layout, zoom,
selection, active sheet and workbook Saved state on success and supported failure
paths. Capture each pane's visible origin through `Pane.VisibleRange`
and compare it afterwards: native scroll getters alone can disagree with the
actual visible origin. Retain that diagnostic evidence rather than assuming
getter equality proves restoration. Confirm hidden/minimized visibility
restoration. Oversized cells, fully hidden ranges and incomplete crops must fail explicitly.

## Packaged acceptance

Build both self-contained Windows x64 packages from the final clean commit:

```powershell
pwsh -NoProfile -File scripts/Build-ForkRelease.ps1 -Version 4.0.0
```

Verify source identity, SHA-256 hashes, PE architecture and executable versions.
Against extracted packages, run `scripts/tests/fork_protocol_smoke.py` with a
15-second idle observation and `scripts/tests/fork_packaged_excel_smoke.py` for
both entry points. Require their protocol, workbook and lifecycle gates.

Run the new public authentication harness against both extracted executables:

```powershell
python scripts/tests/fork_vba_authentication_smoke.py --mcp-exe <EXTRACTED_MCP_EXE> --cli-exe <EXTRACTED_CLI_EXE> --source-revision <FINAL_40_CHARACTER_COMMIT> --report-directory <NEW_ABSOLUTE_PRIVATE_DIRECTORY>
```

This harness uses product MCP/CLI actions for authentication; it does not use
the older native password helper to substitute for product authentication.
Require persisted locking; recognized wrong verification and unlock results;
correct verification; already-accessible behavior; native recognized rejection
of nonempty all-space input through both public entry points; guarded source
and cell edits; explicit save/fresh-open locking; correct fresh unlock and exact persisted
readback; unchanged file bytes during authentication; and normal owned cleanup.
For all-space input against the synthetic non-space password, require
`passwordVerified:false` from verification and `Permissions` from unlocking,
with authentication performed and the project still locked. This proves exact
submission to the challenge, not creation or acceptance of an all-space
password. No-echo checks must compare actual generated credentials and their
encoded forms, without mistaking ordinary output whitespace for a secret.
CLI password values must travel through serialized batch stdin. Keep real
credentials and operational workbooks out of fixtures and receipts.

Snapshot runtime/harness inputs and confirm their hashes after execution. Any
manual/forced cleanup, leftover dialog or uncertain owned-process shutdown
fails acceptance. Record actual Excel version/build and Windows configuration.

## Publication and installation

Publish both ZIPs, `RELEASE-MANIFEST.json`, `VERIFICATION-RESULTS.json` and
`SHA256SUMS`. Independently download and compare every draft asset before
publication and every published asset afterwards. Hosted Actions stay disabled.

Install the exact published bytes, verify hashes, migrate the existing MCP
launcher and generation-specific CLI pipe, and check a fresh configured MCP
plus the normal CLI wrapper. Preserve active work before retiring superseded
owned hosts. Recycle old installations and redundant builds; retain unique
source and evidence. Existing-chat reconnection is a separate activation check.

## Limits and remaining evidence boundaries

Authentication requires an interactive desktop, a normal writable owned
session, printable ASCII and recognized English native dialog signatures.
Unknown dialogs fail closed. Recognized rejection is distinct from an unknown
outcome; cancellation is not rollback, and accepted access may remain open
after an error. Inspect `passwordVerified`, `accessUnlocked`, cleanup/restoration
flags and `rollbackAttempted` before another request.

Transport tests cover unknown-state serialization, cancellation and timeout
classification. They do not prove native cancellation at every dialog phase,
unknown-dialog handling under injected native failures, or recovery from every
cleanup failure. Require refusal of excessive input; this does not establish an exhaustive
exact-maximum password-length boundary across Office versions.

Screenshots photograph the live window, briefly bring Excel to the foreground,
temporarily show hidden/minimized Excel and require an interactive desktop. Frozen-pane divider lines may remain visible; existing
reported top-left truncation limits still apply to very large areas. Synthetic
checks do not establish every Office/language configuration, customer workbook
or print equivalence, long-duration stability, or historical RPC causes.
