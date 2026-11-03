# Fork 3.5.0 verification

This minor release adds MCP batch and custom slicer styles, and integrates CLI
startup-lock cancellation. The catalogue has 442 operations, 64 MCP tools and
33 feature areas. This guide defines gates; `VERIFICATION-RESULTS.json` records
their actual execution and source identity.

## Final source

Build clean committed source with SDK 10.0.403 in Release with zero warnings
and errors. Freeze source and the matching test/product support closure before
tests, and verify both afterwards. Require the existing 22 CLI discovery/typed
argument regressions, generated Core/CLI/MCP contracts, and fork packaging and
plugin recovery tests with discovery reconciled to executed cases.

Run all 11 startup-cancellation cases and existing serialization, startup
deadline and abandoned-mutex controls. Require the batch protocol/execution
tests covering audit mode, bounded reads and default/nested session aliases.
Run the affected slicer-style option and CLI/MCP protocol tests.

Against desktop Excel, run the slicer-style and layout cases together with
existing table-style cases, the MCP batch save/reopen lifecycle smoke and all
18 native range-error identity cases. Record actual regional settings; en-US
coverage alone does not establish a second native regional configuration.
Earlier regional named-range results remain historical evidence.

## Portable packages

Build both packages from the same frozen source:

```powershell
pwsh -NoProfile -File scripts/Build-ForkRelease.ps1 -Version 3.5.0
```

Run `scripts/tests/fork_protocol_smoke.py` against the extracted MCP executable,
then `scripts/tests/fork_packaged_excel_smoke.py` sequentially against both
extracted executables. Verify the complete package/executable hashes before and
after, protocol parity, workbook state and normal owned cleanup. Confirm the
packaged MCP catalogue includes batch and slicer-style operations and the CLI
`list-actions` remains available without Excel.

Retain the five individual CLI JSON-flag controls: camelCase numeric/Boolean
VBA arguments, PascalCase numeric argument, multi-column range sorting and
descending table sorting. Check result types and actual persisted cell order.

Publish both ZIPs, `RELEASE-MANIFEST.json`, `VERIFICATION-RESULTS.json` and
`SHA256SUMS` together. Independently download draft assets and compare bytes
before publication. Inherited Actions publication workflows remain disabled.

## Installation and limits

Download the published assets, verify hashes, switch the configured MCP command
and generation-specific CLI pipe, and check the normal CLI wrapper plus a fresh
MCP connection. Inspect pending sessions before retiring superseded owned hosts.
Recycle old installations and redundant owned outputs; preserve work and evidence.
Existing-chat reconnection is a separate activation boundary.

Packages are unsigned Windows x64 distributions requiring desktop Excel and an
interactive session. Synthetic checks do not prove customer workbook calculation,
print equivalence, all-day stability, every Office configuration or the cause of
historical RPC and StatusBar observations. No new benchmark is claimed.
