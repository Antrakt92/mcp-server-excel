# Fork 3.0.1 verification

Fork 3.0.1 is a maintenance release after 3.0.0. It integrates upstream 2.6.1
and the subsequent chart-label removal fix (`efd33740`), fork named-range
creation repairs, dependency updates and clearer UsedRange guidance. No new
public capability or incompatible contract is introduced.

The release's `VERIFICATION-RESULTS.json` records actual checks and omissions.
This page explains that evidence and its reproduction; it is not a substitute
for a completed receipt and does not claim unrecorded test results.

## Match evidence to packages

Download these assets together from the same fork release:

- `ExcelMcp-Antrakt-MCP-3.0.1-win-x64.zip`
- `ExcelMcp-Antrakt-CLI-3.0.1-win-x64.zip`
- `RELEASE-MANIFEST.json`
- `VERIFICATION-RESULTS.json`
- `SHA256SUMS`

Compare every payload hash with `SHA256SUMS`. Match the verification receipt's
source revision and executable hashes to `RELEASE-MANIFEST.json` and the actual
extracted executables. The manifest separately records the clean source commit
and tree, SDK, architecture and package metadata checks. Successful packaging
alone does not establish protocol, workbook or native cleanup acceptance.

## Reproduce local checks

Use the pinned SDK from `global.json`, Windows x64, PowerShell 7.4 or later and
installed desktop Excel. Follow [BUILD.md](BUILD.md), including its isolated ADO
importer environment when applicable. Build from clean committed source, then
extract each archive into its own directory before checking its binaries.

With Python 3.11 or later:

```powershell
python scripts/tests/fork_protocol_smoke.py --exe '<MCP_EXE>' --report protocol-results.json --idle-seconds 15
python scripts/tests/fork_packaged_excel_smoke.py --mcp-exe '<MCP_EXE>' --cli-exe '<CLI_EXE>' --report native-results.json
python scripts/tests/fork_cli_batch_benchmark.py --cli '<CLI_EXE>' --source-revision '<SOURCE_COMMIT>' --output cli-performance.json
```

The protocol check covers initialization, discovery, result parity, EOF, broken
stdout with stdin retained, parent exit with a retained writer and bounded idle
observations. It does not open Excel. The native check uses disposable synthetic
XLSM fixtures through both entry points: save/reopen, read-only mutation refusal,
hash preservation, bounded reads and guarded VBA invocation. Run native commands
serially and require confirmed fixture, host and Excel cleanup.

The CLI benchmark uses a private daemon and empty session lists. Its warmed
rounds compare separate CLI processes with a single batch process and record
startup, workload, shutdown, cleanup and unchanged binary hashes. These timings
measure process/connection overhead rather than Excel calculation or printing.

Source checks and final package checks are separate. Retain the actual local
build, affected regression, generated-contract and script/package results in the
verification receipt; do not reuse a source result as packaged-binary evidence.

## Limits and distribution boundary

Synthetic regressions do not establish customer workbook calculation, print
equivalence, every Office version or all-day resource stability. Busy recovery
coverage does not establish a fix for an unreproduced cloud AutoSave hang.
Bounded idle observations are not an all-day leak guarantee.

This portable fork release distributes unsigned Windows x64 MCP and CLI ZIPs.
It does not publish the upstream npm, NuGet, extension or MCP-registry identities.
Their source-channel metadata remains upstream 2.6.1; portable binaries explicitly
use fork 3.0.1. Earlier releases and their evidence remain immutable and available
for rollback; changing binaries does not undo workbook edits.
