# Fork 3.4.1 verification

This patch corrects numeric spill classification and integrates upstream
2.6.8 with maintenance through `edb5eada`. The fork retains 438 operations,
63 MCP tools and 33 feature areas. This guide defines the gates;
`VERIFICATION-RESULTS.json` records what actually ran.

## Source checks

Build the clean committed solution in Release using SDK 10.0.403 with zero
warnings/errors. Require 22 CLI discovery/typed-list regressions, 332 generated
Core/CLI/MCP contracts and 19 fork-builder/plugin-recovery checks. Run all 18
`PersistentServiceRangeErrorIdentityTests` against desktop Excel, with discovery
and execution reconciled. Integration also covers the existing value, formula,
trace and spill suites; keep those receipts distinct from final-source evidence.
Unchanged extension packaging evidence remains in the 3.4.0 integration record.
LLM evaluations are on-demand and are not part of this release gate.

Record commands, results, source commit/tree and the complete executable/test
support closure. Rebuild the matching closure after the final source commit.

## Portable package checks

Build both packages from the same frozen source:

```powershell
pwsh -NoProfile -File scripts/Build-ForkRelease.ps1 -Version 3.4.1
```

Run `scripts/tests/fork_protocol_smoke.py` against the extracted MCP executable,
then `scripts/tests/fork_packaged_excel_smoke.py` sequentially against both
extracted executables. Require passing protocol, workbook-state and cleanup
checks, with exact binary hashes before and after. Also verify that the packaged
CLI's `list-actions` output includes workbook and session actions without
connecting to Excel. These are synthetic checks, not private-workbook acceptance.

Also run five individual CLI JSON-flag cases: camelCase numeric and Boolean VBA
arguments, a PascalCase numeric argument, multi-column range sorting and
descending table sorting. Check returned scalar types and actual cell order,
with fixture and executable/support hashes unchanged through the checks.

Publish the MCP/CLI ZIPs, `RELEASE-MANIFEST.json`, `VERIFICATION-RESULTS.json`
and `SHA256SUMS` together. Download draft assets independently and verify hashes
before publishing. Keep inherited publication workflows disabled.

## Local activation and limits

Install the exact verified release through the maintained local activation
procedure, switch the CLI pipe and MCP command, and test the normal CLI wrapper
plus a fresh MCP connection. Inspect pending sessions before retiring owned old
hosts. Recycle superseded installations while preserving work and evidence.
Report existing-client reconnection separately from fresh-process activation.

Packages are unsigned Windows x64 distributions requiring installed desktop
Excel and an interactive session. Passing controls do not explain historical
COM/RPC failures, prove all-day stability or establish customer calculation and
print equivalence. No new benchmark is claimed.
