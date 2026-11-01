# Fork 3.3.0 verification

Fork 3.3.0 integrates upstream 2.6.7 chart controls and protects existing named
ranges during creation. It exposes 438 operations through 63 MCP tools across
33 feature areas. This guide defines the checks; the release's
`VERIFICATION-RESULTS.json` records what actually ran. A listed command is not a
passing result.

## Match evidence to the packages

Download these five assets from the same release:

- `ExcelMcp-Antrakt-MCP-3.3.0-win-x64.zip`
- `ExcelMcp-Antrakt-CLI-3.3.0-win-x64.zip`
- `RELEASE-MANIFEST.json`
- `VERIFICATION-RESULTS.json`
- `SHA256SUMS`

Check archive, manifest and verification-receipt hashes against `SHA256SUMS`.
Match the receipt's source commit/tree and executable hashes to the manifest and
extracted binaries. PE metadata and archive checks establish packaging integrity;
they do not establish workbook behavior or client activation.

## Freeze source and verify the build

Use committed, clean source with no other writers. Record source inputs before a
standard Release rebuild of the solution. Retain the successful build log and
its hash, source commit/tree and build interval. Capture the resulting test
assemblies and product/worker support closure, then prove those inputs remain
unchanged through verification. Vendor dependencies copied by MSBuild may retain
their original timestamps; compiled owned assemblies must belong to this build.

Run the canonical selected Excel-free CLI, Core, MCP and Service checks, with
discovery/execution reconciliation and exact expected case counts. Record the
filters and results. This is selected coverage, not a full-suite claim. Run the
12 release-builder regressions and 33 typed test-report regressions when verifying
the changed packaging defaults and release evidence.

Run the 55 native named-range cases and 186 selected chart/PivotChart cases
sequentially against the same frozen Service build. Require complete discovery
reconciliation and confirmed owned-process cleanup. The named-range group covers
duplicate identities, hidden/local names and save/reopen protection. The chart
group covers title/axis controls and ordinary/Data Model PivotChart persistence.

Record the actual desktop Excel/Windows regional settings. Caller `CultureInfo`
changes do not establish a second host configuration. Historical runs retain
their original revisions and settings; do not report them as fresh checks of
this release.

## Check the final executables

Build with `scripts/Build-ForkRelease.ps1 -Version 3.3.0` using the prerequisites
in [BUILD.md](BUILD.md). Extract each ZIP into its own directory. With Python
3.11 or later:

```powershell
python scripts/tests/fork_protocol_smoke.py --exe '<MCP_EXE>' --report protocol-results.json --idle-seconds 15
python scripts/tests/fork_packaged_excel_smoke.py --mcp-exe '<MCP_EXE>' --cli-exe '<CLI_EXE>' --report native-results.json
```

The protocol check covers initialization/discovery, response content, EOF,
broken stdout with stdin retained, parent exit and bounded idle observations.
It does not open Excel.

The packaged native check uses disposable synthetic XLSM fixtures through both
MCP and CLI. It verifies bounded reads, worksheet views, guarded VBA invocation,
read-only mutation refusal and unchanged file hashes during audit. It also checks
that rejected duplicate-name aliases preserve the original definition and
visibility, and that title placement/alignment and axis line settings survive
save/reopen. The receipt must show these checks for both entry points, followed
by confirmed fixture, host and owned Excel cleanup.

No new performance claim is made. A CLI benchmark is optional for this release;
if omitted, the receipt must explicitly mark it as not run. Prior measurements
remain evidence for their original versions and workloads.

## Publication and local activation

Download uploaded assets independently and verify their bytes against the
accepted local files, checksums and GitHub asset digests before publication.
After installation, verify the normal CLI entry point and a fresh MCP connection.
An updated configuration path does not prove an existing client reconnected.
Record publication, installation, activation and retirement of superseded
installations separately.

## Limits

The release distributes unsigned, self-contained Windows x64 MCP and CLI ZIPs.
It does not publish upstream npm, NuGet, extension or registry channels. Synthetic
fixtures do not establish customer calculation/print equivalence, every Office
version or all-day stability. Missing regional coverage remains an explicit
limit. Installing a new binary does not undo workbook edits.
