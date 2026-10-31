# Fork 3.2.0 verification

Fork 3.2.0 integrates upstream 2.6.6 worksheet scroll/selection controls
and adds completed-session cancellation recovery. It exposes 437 operations through
63 MCP tools across 33 feature areas.

This guide describes the verification route. The release's
`VERIFICATION-RESULTS.json` records the checks actually run, source provenance,
executable hashes and omitted coverage. Listing a command here is not evidence
that it passed.

## Match evidence to the packages

Download these assets from the same release:

- `ExcelMcp-Antrakt-MCP-3.2.0-win-x64.zip`
- `ExcelMcp-Antrakt-CLI-3.2.0-win-x64.zip`
- `RELEASE-MANIFEST.json`
- `VERIFICATION-RESULTS.json`
- `SHA256SUMS`

Compare the archive, manifest and verification-receipt hashes with `SHA256SUMS`.
Match the receipt's source commit/tree and executable hashes to the manifest and
extracted binaries. Packaging checks establish PE architecture, version metadata
and archive contents; they do not establish workbook behavior or client activation.

## Freeze source and select the checks

Use clean, committed source with no other writers. Build the selected test and
product projects together from that revision before running tests. Capture source
and build-support identities before and after verification, including test
assemblies and the CLI/MCP worker closure. A matching test DLL alone does not prove
that its worker remained compatible.

The source matrix for this release covers the selected Excel-free CLI, Core, MCP
and Service checks, the named-pipe cancellation regressions, and generated action
contracts. This is selected coverage, not a claim that all eight test projects or
every test group ran. Record exact filters, discovered/executed cases, results and
source identities in the receipt.

Run the 12 release-builder admission/metadata regressions after changing packaging
defaults or guards. The native window matrix comprises the selected 99 Service
window/view/persistence cases and 8 MCP routing cases. Keep recording-only MCP
checks distinct from tests that actually start Excel; record the executed scope
rather than describing all 107 cases as native.

Historical evidence retains its original revision, commands and outcomes. Reuse
requires a reviewed comparison of runtime, generated contracts, dependencies, SDK
and relevant test/build inputs. Do not carry forward 3.1.0's full-suite totals as
fresh evidence for 3.2.0.

The test-report harness also has a focused 33-case regression suite. It checks
Unicode name reconciliation without relaxing case multiplicities or outcome gates.

## Check the final executables

Use Windows x64, installed desktop Excel, PowerShell 7.4 or later and the SDK pinned
in `global.json`. Follow [BUILD.md](BUILD.md), including the ADO importer
prerequisite. Build the committed source through `scripts/Build-ForkRelease.ps1`
and extract each archive into its own directory. With Python 3.11 or later:

```powershell
python scripts/tests/fork_protocol_smoke.py --exe '<MCP_EXE>' --report protocol-results.json --idle-seconds 15
python scripts/tests/fork_packaged_excel_smoke.py --mcp-exe '<MCP_EXE>' --cli-exe '<CLI_EXE>' --report native-results.json
python scripts/tests/fork_cli_batch_benchmark.py --cli '<CLI_EXE>' --source-revision '<SOURCE_COMMIT>' --output cli-performance.json
```

The protocol check covers initialization/discovery, matching response content,
EOF, broken stdout with stdin retained, parent exit and bounded idle observations
without opening Excel. Record its three cases and cleanup results.

The packaged native check uses disposable synthetic XLSM fixtures through both
MCP and CLI. It checks save/reopen, read-only mutation refusal, file-hash
preservation, bounded reads and guarded VBA invocation. Include the new
`window set-view`/`get-view` route in final-artifact acceptance, rather than relying
only on the unchanged count of 63 MCP tools. Run native groups sequentially and
require confirmed fixture, host and Excel cleanup for each entry point.

The CLI benchmark runs three rounds of eight commands per mode on a private
daemon, comparing separate CLI processes with one ordered batch. Validate empty
session lists, source/binary identity and daemon cleanup. Report startup, workload
and shutdown costs. These measurements cover client overhead; they do not measure
Excel calculations, printing or whole-workbook performance.

## Publication and local activation

Verify the uploaded assets independently against the accepted local files and
checksums. After installation, verify the version through the normal MCP/CLI
entry points. A new configured path does not prove an existing MCP connection
has reconnected. Keep publication, local installation, activation and retirement
of superseded installations as separate results.

## Limits

The release distributes unsigned, self-contained Windows x64 MCP and CLI ZIPs.
Upstream npm, NuGet, extension and registry metadata remains at 2.6.6; the portable
fork builder does not publish those channels. Synthetic fixtures do not prove
customer workbook calculations, print equivalence, every Office version or
all-day stability. Switching binaries does not undo workbook edits.
