# Fork 3.3.1 verification

This patch fixes the CLI parser for structured JSON lists. It retains upstream
2.6.7 and the same 438 operations, 63 MCP tools and 33 feature areas. This guide
defines the checks; `VERIFICATION-RESULTS.json` records what actually ran.

## Source and package identity

Use clean committed release source with the SDK in `global.json`. Build the
solution in Release with zero warnings/errors. Record its commit/tree and build
log, then keep the product/test support files unchanged during verification.

Run the eight CLI list/argument regressions, generated Core/CLI/MCP contract
checks and the 12 fork-builder cases. Reconcile actual test counts with discovery
and retain the TRX files. These are selected suites, not a full-suite claim.
The broader named-range and chart suites recorded for 3.3.0 are historical;
representative controls are checked again through the packaged smoke below.

Build both portable packages from the same frozen source:

```powershell
pwsh -NoProfile -File scripts/Build-ForkRelease.ps1 -Version 3.3.1
```

Download or retain these five assets together:

- `ExcelMcp-Antrakt-MCP-3.3.1-win-x64.zip`
- `ExcelMcp-Antrakt-CLI-3.3.1-win-x64.zip`
- `RELEASE-MANIFEST.json`
- `VERIFICATION-RESULTS.json`
- `SHA256SUMS`

Match archive, manifest and receipt hashes to `SHA256SUMS`. Match the receipt's
source and executable identities to the manifest and extracted binaries.

## Packaged acceptance

Run each native group sequentially using disposable synthetic books:

```powershell
python scripts/tests/fork_protocol_smoke.py --exe '<MCP_EXE>' --report protocol-results.json --idle-seconds 15
python scripts/tests/fork_packaged_excel_smoke.py --mcp-exe '<MCP_EXE>' --cli-exe '<CLI_EXE>' --report native-results.json
```

The protocol smoke covers discovery, result parity, EOF, broken stdout and parent
exit. The native smoke checks both entry points, audit refusals, bounded reads,
VBA, worksheet views, named-range protection and chart save/reopen behavior.
Require confirmed fixture/session/host/native cleanup and unchanged binaries.

Also exercise individual CLI flags against the extracted CLI: camelCase numeric
and Boolean VBA arguments, existing PascalCase arguments, multi-column range
sorting and descending table sorting. Verify returned types and actual cell
order, not just command success. Fingerprint the complete executable/support
closure before and after. Record cleanup separately from native exit codes.

## Publish and activate

Use the fork's local portable builder and authenticated GitHub CLI. Download the
uploaded draft assets independently and verify their bytes/digests before
publication. Do not run inherited upstream publication workflows.

Install the exact verified release, switch the intended MCP configuration and
CLI generation pipe, then check the normal wrapper and a fresh MCP process.
Inspect sessions/workers before retiring old hosts. Preserve unsaved work and
unique evidence; recycle superseded installations after migration.

Configuration changes and fresh-process checks do not prove an existing client
has reconnected. Record that activation separately. Client-transformed schemas
may still limit scalar arguments; use the verified CLI route when needed.

## Limits

Packages are unsigned Windows x64 distributions requiring installed Excel and an
interactive desktop. Synthetic tests do not establish customer calculation or
print equivalence, all Office versions, historical crash causes or all-day
stability. Record actual regional settings and missing coverage. No benchmark or
new performance claim is required for this patch.
