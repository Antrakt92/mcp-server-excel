# Fork 3.1.0 verification

Fork 3.1.0 integrates upstream 2.6.4, including chart text formatting, native
PivotCache sharing and close-stage timings. It retains the fork's audit sessions,
bounded reads, guarded VBA and process ownership safeguards. The source exposes
436 operations through 63 MCP tools across 33 feature areas.

This page describes the verification route. The release's
`VERIFICATION-RESULTS.json` records actual results, source provenance and skipped
checks; this guide does not certify a run by listing its command.

## Match evidence to packages

Download these assets from the same release:

- `ExcelMcp-Antrakt-MCP-3.1.0-win-x64.zip`
- `ExcelMcp-Antrakt-CLI-3.1.0-win-x64.zip`
- `RELEASE-MANIFEST.json`
- `VERIFICATION-RESULTS.json`
- `SHA256SUMS`

Compare the ZIP, manifest and verification-receipt hashes with `SHA256SUMS`.
Match the receipt's source revision and executable hashes to the manifest and
the extracted executables. The manifest records the clean source commit and
tree, upstream base, SDK and packaging checks. Packaging alone does not prove
protocol behavior, workbook handling or native cleanup.

## Source evidence and release checks

Keep source tests separate from checks of the final packaged executables. If a
previous source test run is reused, retain its original revision, commands,
counts and results. Record a comparison showing that the runtime, generated
contracts, dependencies, SDK and relevant test inputs are unchanged between
that revision and the release source. Explain each excluded path. A matching
version or a small diff does not establish equivalence.

Run affected release-builder regressions separately when packaging defaults or
guards change. Record their exact source and results. Do not relabel earlier
tests as a fresh full-suite run at the release commit. Final package checks must
use the newly built executables, even when source-test evidence is reusable.

## Check the packaged executables

Use Windows x64, installed desktop Excel, PowerShell 7.4 or later and the SDK
pinned in `global.json`. Follow [BUILD.md](BUILD.md), including the ADO importer
requirements where applicable. Build clean committed source, then extract each
ZIP to its own directory. With Python 3.11 or later:

```powershell
python scripts/tests/fork_protocol_smoke.py --exe '<MCP_EXE>' --report protocol-results.json --idle-seconds 15
python scripts/tests/fork_packaged_excel_smoke.py --mcp-exe '<MCP_EXE>' --cli-exe '<CLI_EXE>' --report native-results.json
python scripts/tests/fork_cli_batch_benchmark.py --cli '<CLI_EXE>' --source-revision '<SOURCE_COMMIT>' --output cli-performance.json
```

The protocol check exercises initialization, discovery, matching result content,
EOF, broken stdout with stdin retained, owning-parent exit and bounded idle
observations without opening Excel. The native check uses synthetic XLSM
fixtures through MCP and CLI: save/reopen, audit mutation refusal, hash
preservation, bounded reads and guarded VBA invocation. Run native checks
serially and require confirmed fixture, host and Excel cleanup.

The CLI benchmark warms both modes and compares separate processes with one
ordered batch on a private daemon. It validates empty session lists and records
startup, workload, shutdown, cleanup and executable hashes. It measures CLI
launch and connection overhead; it does not measure Excel calculation or printing.

## Limits

Retain actual outcomes and omissions in the receipt. Synthetic fixtures do not
prove customer workbook calculations, print equivalence, every Office version
or all-day stability. A local package check does not prove remote asset integrity
or that a user's MCP client has switched to the new executable.

This release distributes unsigned, self-contained Windows x64 MCP and CLI ZIPs.
Upstream npm, NuGet, extension and registry metadata remains at 2.6.4; those
channels are not published by the portable fork builder. Earlier releases keep
their original source bases and receipts. Switching binaries does not undo
workbook edits.
