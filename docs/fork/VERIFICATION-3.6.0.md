# Fork 3.6.0 verification

This feature release adds native VBA project password configuration to MCP and
CLI. The upstream baseline remains 2.6.10; no dependencies change.

## Source checks

Use clean committed source, SDK 10.0.403 and the reviewed ADO importer. Build
Release with zero warnings/errors. Run the Commit Excel-free groups, generated
contract checks and fork packaging checks. Retain the source commit/tree and
all failed or interrupted runs separately.

Run the real-Excel Service VBA regressions, including
`PersistentServiceVbaProtectionTests`. Require empty-project, overlong-password,
audit-session and already-locked refusals; preserved source; configured target
and unchanged sibling project in the same VBE; restored project/visibility;
save/fresh-open locking; and normal owned shutdown. Build and run the maintained
native authentication helper separately:

```powershell
dotnet build scripts/tests/native/VbaProjectProtectionAuthProbe/VbaProjectProtectionAuthProbe.csproj -c Release
```

## Packaged acceptance

Build both self-contained portable packages with
`scripts/Build-ForkRelease.ps1 -Version 3.6.0` from the final clean commit.
Verify source identity, SHA-256 hashes, PE architecture and executable versions.

Against extracted packages, run `scripts/tests/fork_protocol_smoke.py` with a
15-second idle observation and `scripts/tests/fork_packaged_excel_smoke.py` for
both entry points. Require the existing protocol, workbook and lifecycle checks.

Also run `scripts/tests/fork_vba_protection_smoke.py` sequentially with
`--entrypoint mcp` and `--entrypoint cli`. Supply the corresponding extracted
`--exe`, built `--auth-probe` and a `--report` under a new private directory.
The harness snapshots its runtime/helper inputs before opening Excel and checks
their hashes again afterwards. Require all of:

- Explicit non-ASCII rejection before source/protection changes.
- Pending save/reopen flags, then actual locking after fresh open.
- Rejection of distinct character, trimmed, case-changed and slash-changed
  passwords, followed by acceptance of the exact synthetic password.
- Password replacement on the authenticated accessible project: save/fresh open,
  rejection of the old password and acceptance of the new password.
- Unchanged VBA source, workbook cells and other project.
- Confirmed cleanup of every introduced dialog and normal exit of every owned
  Excel, worker and controller. Manual or forced cleanup keeps the run failed.

Record the actual Excel version/build/OS. Keep passwords in stdin/protocol only;
never put real credentials or operational workbooks into acceptance artifacts.

## Publication and installation

Publish both ZIPs, `RELEASE-MANIFEST.json`, `VERIFICATION-RESULTS.json` and
`SHA256SUMS`. Independently download and compare every draft asset before
publication and every published asset afterwards. Hosted Actions stay disabled.

Install the exact published bytes, verify hashes, migrate the existing MCP
launcher and generation-specific CLI pipe, and check a fresh configured MCP
plus the normal CLI wrapper. Preserve active work before retiring superseded
owned hosts. Recycle old installations and redundant builds; retain unique
source and evidence. Existing-chat reconnection is a separate activation check.

## Scope

The driver supports the standard VBE dialog in an interactive Windows desktop.
Accepted passwords use printable ASCII because the public object model does not
provide the project's password code page. Synthetic checks do not prove every
Office configuration, customer workbook calculations/printing, long-duration
stability or a cause/fix for historical RPC observations.
