# Antrakt Fork changelog

Release history for [Antrakt92/mcp-server-excel](https://github.com/Antrakt92/mcp-server-excel).
Dates are the original GitHub publication dates, in UTC. Changes are grouped by
the first fork release that contains them. The [upstream changelog](../../CHANGELOG.md)
retains the original project's version numbers and contributor credits.

| Version | Published | Summary |
| --- | --- | --- |
| [3.2.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.2.0) | 2026-10-31 | Worksheet scroll/selection controls and completed-session cancellation recovery |
| [3.1.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.1.0) | 2026-10-30 | Chart text formatting, native PivotCache sharing, close timings and MCP response optimization |
| [3.0.1](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.0.1) | 2026-10-29 | Named-range and chart fixes, idle-resource repair, dependency updates |
| [3.0.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.0.0) | 2026-10-29 | Audit sessions, bounded reads, guarded VBA and CLI migration |
| [2.5.4](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v2.5.4) | 2026-10-25 | First stable fork: process lifetime and portable packages |

## 3.2.0 — 2026-10-31

Set a worksheet's scroll position and selection through MCP or CLI, and recover
access to a workbook when cancellation races with a completed open/create reply.
Based on upstream 2.6.6, with 437 operations instead of 436.

### Worksheet views

- `window set-view` sets the top-left cell and optionally selects a range on the
  named worksheet. For example, use `--sheet-name Dashboard --top-left-cell E20
  --select H30:I32` in the CLI. The MCP action is `window` / `set-view`, with
  `sheet_name`, `top_left_cell` and optional `select`.
- `window get-view` reports Excel's actual scroll origin, selection kind,
  selected range and active cell. The returned fields include `topLeftCell`,
  `selectionKind`, `selectedRangeAddress` and `activeCellAddress`. Non-range
  selections keep the two address fields present with null values.
- Set the view before saving to control where a workbook opens. Existing pane,
  zoom and display settings remain available; returned positions reflect Excel's
  native behavior, including frozen panes, hidden rows and merged cells.

### Cancellation after a completed reply

A workbook could finish opening or being created just as the caller cancelled.
The client then discarded the reply containing its session ID, leaving the book
open but hidden. It now requests `window.show` for that exact completed session
before reporting cancellation, with a separate five-second recovery budget.

This does not save, close, undo or repeat workbook work. A failed reveal is
reported in diagnostics while the original caller cancellation is preserved.
Failed opens and unrelated commands do not trigger the recovery request.

### Upstream and maintenance

- Integrates upstream 2.6.5 and 2.6.6. Upstream accepted our generated MCP response
  optimization in [#1278](https://github.com/sbroenne/mcp-server-excel/pull/1278).
  The fork keeps the single-document conversion already shipped in 3.1.0.
- Includes upstream's explicit CLI worker build dependency for Service tests,
  which the fork already carried, and clarified website/release metadata guidance.
- The internal snapshot-ordering test now uses a synthetic worker client instead
  of injecting a fabricated release into a live worker process. This avoids
  racing its actual shutdown during local validation.
- Test-case reconciliation now treats supplementary Unicode characters and their
  VSTest TRX surrogate escapes consistently, while keeping literal escapes,
  duplicate counts and missing cases distinct.
- The pinned .NET SDK remains 10.0.403. Runtime and development dependency
  versions are unchanged from 3.1.0; the upstream product metadata advances from
  2.6.4 to 2.6.6.

### Upgrade and verification

Use the 3.2.0 MCP/CLI packages from this fork, verify their checksums, and reconnect
MCP clients or restart the relevant CLI service. Existing commands remain
compatible; no breaking migration is needed.

See [3.2.0 verification](VERIFICATION-3.2.0.md). The release's
`VERIFICATION-RESULTS.json` binds actual source, protocol, native Excel and CLI
checks to the shipped executable hashes. Synthetic fixtures do not establish
customer calculation or print equivalence. Portable Windows x64 packages remain
unsigned.

## 3.1.0 — 2026-10-30

Adds chart text formatting, reduces MCP response-processing work and simplifies
PivotCache sharing. Based on upstream 2.6.4, with 436 operations instead of 434.

### MCP responses: 1.60–2.16× faster conversion

Generated MCP calls now pass the normalized Service result straight to reply
conversion. The old path serialized that result and then parsed and normalized
it again. Removing that extra pass preserves returned data, errors, warnings,
partial state and telemetry.

In seven local non-image fixtures, the typed route was **1.60–2.16× faster** and
allocated **about 43–55% fewer managed bytes** per conversion. Selected results:

| Response | String route | Typed route | Speedup | Fewer allocated bytes |
| --- | --- | --- | --- | --- |
| Small grid, about 1 KB | 11.48 µs | 7.08 µs | 1.62× | 43.1% |
| Grid, about 100 KB | 344.28 µs | 215.72 µs | 1.60× | 42.6% |
| Grid, about 224 KB | 437.01 µs | 255.21 µs | 1.71× | 42.6% |
| Failure with partial state | 20.61 µs | 9.53 µs | 2.16× | 54.7% |

These are median per-response timings from five alternating rounds on
Windows/.NET 10.0.14, with 200 warmups and 400 iterations per round
(100 for the larger grids). Both routes used the same candidate assembly and
shared execution helper; output and telemetry parity were checked. The other
fixtures covered scalar, null and nested-failure results.

The measurement covers in-process response conversion, excluding transport,
telemetry delivery and Excel work. It is not a speedup for Excel calculations,
refreshes or complete requests. Screenshot conversion uses its existing path;
the unchanged screenshot control had identical allocated-byte counts.

### PivotCache: direct reassignment instead of item-by-item checks

`share-cache` now asks Excel to assign the selected cache directly. The previous
implementation scanned cached items before the assignment, adding COM calls as
the cache grew. Excel now decides which assignments are valid, with native errors
and resulting-state diagnostics preserved.

In one native Excel test with **200,000 distinct cached items**, the cache-sharing
call took **0.10 seconds**. The target cache contained 200,003 items after a
refresh; the moved report was checked against that target. Creating the tables,
building/refreshing caches, Excel startup and transport are outside that timing.
No old-path timing was collected, so this is a recorded result, not a
before/after speedup ratio.

A moved report uses the selected cache's data. Excel may also change its report
settings during reassignment; preserving the old configuration is not guaranteed.

### Chart text formatting

- Read or change font name, size, bold, italic and color on existing chart areas,
  titles, axes, axis titles, legends and data labels.
- Works with ordinary charts and PivotCharts. MCP uses `chart_config_read`
  with `get-text-format` and `chart_config` with `set-text-format`; CLI uses
  `chartconfig get-text-format` and `chartconfig set-text-format`.
- Missing titles, axes or labels are not created automatically. A partially
  completed formatting operation does not promise rollback.

### Closing workbooks and reporting failures

- Removed unnecessary close-path work and extra status/stop round trips.
  Successful close responses include `closeTimingsMs` so callers can see which
  stages took time.
- Confirmed COM release ends ownership without waiting for the Excel PID to
  disappear. Failed or uncertain release keeps ownership recorded.
- If recording completion fails after a workbook operation has already finished,
  keep its result and explain that the operation must not be repeated.

### Build guards and documentation

- Require the executable's complete `version+commit` stamp to match the selected
  source, including checks for missing or incorrect commit stamps.
- Reject a portable build when its version is already tagged at different source,
  before SDK/build work. Exact tagged-source reproduction remains supported.
- Keep the package README and manifest on the same upstream 2.6.4 base.
- Updated package selection, MCP/CLI setup, contributor and recovery guidance,
  security links, and dated release history.

### Dependencies and upgrading

No dependency or SDK versions changed from fork 3.0.1. The SDK remains 10.0.403
and the MCP SDK remains 2.3.0. Root and extension source-version metadata changed
from upstream 2.6.1 to 2.6.4; no dependency upgrade is implied.

No CLI/MCP schema migration is required from 3.0.1. Switch to the new executable
and reconnect the client to discover the two new chart operations. Download the
ZIP with its manifest, verification receipt and checksum file from this release;
see [the switch procedure](RECOVERY.md#switch-versions).

### Verification

The release checks passed: 12 builder regressions, three packaged MCP protocol
cases, synthetic XLSM workflows through MCP and CLI, and a three-round CLI batch
benchmark. Owned host/Excel processes exited and temporary fixtures were removed.

The earlier 3,438 source tests and 97 native Service tests were reused after
checking unchanged runtime/build/dependency/test inputs; they are not presented
as a fresh full-suite run at the release commit. See
[the verification guide](VERIFICATION-3.1.0.md) and `VERIFICATION-RESULTS.json`
for provenance, package hashes and limits. Synthetic files do not establish
customer-workbook calculation or print equivalence.

## 3.0.1 — 2026-10-29

Maintenance release based on upstream 2.6.1, including its subsequent chart-label
fix. No migration is required from fork 3.0.0.

### Fixes

- **Named-range creation:** create the final name directly instead of adding,
  editing and renaming a temporary workbook name. Invalid names and references
  return an error without leaving temporary names behind. Invariant formula
  handling is retained.
- **Chart labels:** removing all label fields works even when another series
  already has no labels. The operation no longer tries to read an absent label
  collection before disabling it. This fix comes from upstream.
- **Idle MCP resources:** integrated the upstream fix for continuing thread and
  handle growth while idle. Fork EOF, broken-output and owning-client exit
  handling remain in place.

### Workbook guidance and regression coverage

- Clarified that `UsedRange` reports Excel's native bounds, including
  formatting-only cells; it does not identify semantic content or print bounds.
  The returned range behavior is unchanged.
- Added a native regression for clearing part of a merged region, checking that
  the merged cell's stored value is preserved. This adds coverage rather than a
  new clearing implementation.
- Included upstream regression coverage for Excel busy recovery, timed-out
  operations, streamed batch completion and OLE retry boundaries. These checks
  do not establish a fix for an unreproduced cloud AutoSave incident.

### Dependency and toolchain updates

The following versions changed between fork 3.0.0 and 3.0.1. Entries under
development tooling are source/build dependencies, not additional software
installed on a user's machine by the portable packages.

**.NET package pins**

| Package | Previous | 3.0.1 |
| --- | --- | --- |
| `Azure.Core` | 1.63.0 | 1.64.0 |
| `Microsoft.Identity.Client` | 4.91.1 | 4.92.0 |
| `Microsoft.Identity.Client.Extensions.Msal` | 4.91.1 | 4.92.0 |
| `Polly.Core` | 8.8.0 | 8.9.0 |
| `Polly.Extensions` | 8.8.0 | 8.9.0 |
| `Polly.RateLimiting` | 8.8.0 | 8.9.0 |
| `Nerdbank.Streams` | 2.14.408 | 2.14.411 |
| `Microsoft.Extensions.Hosting` | 10.0.13 | 10.0.14 |
| `Microsoft.Extensions.Hosting.Abstractions` | 10.0.13 | 10.0.14 |
| `Microsoft.Extensions.Logging` | 10.0.13 | 10.0.14 |
| `Microsoft.Extensions.Logging.Configuration` | 10.0.13 | 10.0.14 |
| `Microsoft.Extensions.Caching.Abstractions` | 10.0.13 | 10.0.14 |
| `Microsoft.Extensions.Configuration` | 10.0.13 | 10.0.14 |
| `Microsoft.Extensions.Configuration.Binder` | 10.0.13 | 10.0.14 |
| `Microsoft.Extensions.Configuration.EnvironmentVariables` | 10.0.13 | 10.0.14 |
| `Microsoft.Extensions.Diagnostics` | 10.0.13 | 10.0.14 |
| `Microsoft.Extensions.Diagnostics.Abstractions` | 10.0.13 | 10.0.14 |
| `Microsoft.Extensions.FileProviders.Abstractions` | 10.0.13 | 10.0.14 |
| `Microsoft.Extensions.Options` | 10.0.13 | 10.0.14 |
| `Microsoft.Extensions.Options.ConfigurationExtensions` | 10.0.13 | 10.0.14 |
| `Microsoft.Bcl.AsyncInterfaces` | 10.0.13 | 10.0.14 |
| `System.Collections.Immutable` | 10.0.13 | 10.0.14 |
| `System.Reflection.Metadata` | 10.0.13 | 10.0.14 |
| `System.Text.Encoding.CodePages` | 10.0.13 | 10.0.14 |
| `System.Memory.Data` | 10.0.13 | 10.0.14 |
| `System.Security.Cryptography.ProtectedData` | 10.0.13 | 10.0.14 |
| `System.Text.Json` | 10.0.13 | 10.0.14 |
| `System.Threading.RateLimiting` | 10.0.13 | 10.0.14 |

**Build, documentation and evaluation tooling**

| Component | Previous | 3.0.1 | Scope |
| --- | --- | --- | --- |
| .NET SDK | 10.0.402 | 10.0.403 | Pinned build SDK |
| `mkdocs-material` | 9.7.7 | 9.8.0 | Documentation site |
| `@vscode/vsce` | 4.0.0 | 4.1.0 | Extension packaging |
| `@types/node` | 26.6.5 | 26.6.6 | Extension type definitions |
| `oxlint` and its platform bindings | 1.89.0 | 1.90.0 | Extension linting |
| `@azure/msal-node` | 7.1.0 | 7.1.2 | Extension packaging authentication override |
| `@azure/msal-common` used by MSAL Node | 16.14.2 | 16.15.0 | Transitive extension development dependency |
| `pytest-asyncio` | 1.4.0 | 1.4.1 | On-demand evaluation lockfile |
| `github/gh-aw/actions/setup` | 0.93.3 | 0.93.5 | Pinned issue-triage setup action |

The setup action remains commit-pinned (`317d9005` → `32ecafbc`). Updating that
pin does not regenerate or validate the whole inherited workflow. GitHub Actions
remain disabled on this fork, and LLM evaluations were not run for this release.

Dependabot configuration already covered the main dependency groups. Added
weekly coverage for both npm wrapper directories, with grouped updates and
fork-owned labels. Reviewed and merged its evaluation-lockfile and setup-action
updates in [PR #8](https://github.com/Antrakt92/mcp-server-excel/pull/8) and
[PR #9](https://github.com/Antrakt92/mcp-server-excel/pull/9).

### Repository and distribution

- Aligned CODEOWNERS, contributor instructions, issue links, privacy and security
  guidance with this independently maintained fork and its actual local process
  boundaries. Preserved the upstream MIT license and attribution.
- Synced upstream through its chart-label repair without replacing fork changes.
- Updated release preparation and protocol cleanup checks, including a bounded
  idle observation before accepting packaged binaries.
- Distributed portable Windows x64 MCP and CLI ZIPs with bundled .NET runtime,
  checksums, source/package manifest and a separate runtime verification receipt.

### Verification

Local release validation passed: **3,375 Excel-free source cases**, **37 selected
native Excel cases**, three packaged protocol cases and two packaged native
MCP/CLI XLSM workflows. Package checks covered save/reopen, read-only restrictions,
bounded reads, MCP guarded VBA invocation and owned-process cleanup. The full
Release solution build completed with zero warnings/errors.

The CLI benchmark compared eight empty-session queries as separate processes
against one batch, over three warmed rounds. Median workload times were 2.080 s
and 0.260 s. This measures process/connection overhead, not Excel calculation,
workbook reads or printing.

[Release and downloads](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.0.1)
· [Changes since 3.0.0](https://github.com/Antrakt92/mcp-server-excel/compare/v3.0.0...v3.0.1)
· [Verification](VERIFICATION-3.0.1.md)

## 3.0.0 — 2026-10-29

Upgraded the fork from upstream 2.5.3 to 2.6.0 and added inspection and VBA
workflows for agents. The major version reflects incompatible CLI output and
option changes. Existing fork process-lifetime safeguards are retained.

### Added

- **Persistent read-only audit sessions:** reuse one session for inspection;
  refuse agent mutations, explicit calculation, macros, security changes and
  saving before dispatch. Audit identity survives Save As/path reconciliation.
  This API restriction does not prevent interactive Excel edits or calculation
  during opening.
- **Bounded cached reads:** inspect native used-range bounds without reading all
  values, then page through an explicit rectangle. Pages contain at most 10,000
  cells; eager value/formula reads cap geometry at 1,000,000 cells. Responses also
  enforce a conservative 4 MiB budget and preserve cached results without
  requesting calculation.
- **Formatting summaries:** inspect aggregate stored/displayed formatting and
  mixed fields. Detailed per-cell formatting reads are limited to 10,000 cells.
- **Guarded VBA module replacement:** compare the expected source fingerprint
  before deletion, read back stored source, and report the observed partial state
  if editing fails. No automatic source rollback is performed.
- **Typed VBA function results:** `invoke` accepts typed scalar arguments and
  returns scalars or rectangular arrays. Existing `run` behavior is retained;
  invocation does not enable macros. Arrays are limited to 10,000 elements and
  results to a conservative 4 MiB transport budget.

### Fixed and hardened

- Preserve caller cancellation, finite request budgets and infinite pipe-connect
  timeouts instead of confusing one timeout kind with another.
- Reserve workbook ownership across supported physical aliases, including NTFS
  hard links and junctions, to prevent duplicate independent sessions.
- Validate complete VBA procedure replacements before deletion and verify the
  inserted source afterward. Preserve source/read-back failure details.
- Keep legitimate numeric cell values distinct from COM error-code numbers;
  bound formatting reads and long-text responses without silently truncating data.
- Report formula preflight's actual lexical/local-reference scope and unresolved
  checks. Validate target sheet and geometry without scratch writes or calculation.
- Include the CLI worker in Service process-test builds so tests exercise the
  current worker rather than a stale executable.
- Route ADO build-definition checks through the isolated, signed importer when
  configured, preserving the normal machine installation.

### Upstream changes and migration

- CLI commands and batch items now use a common JSON envelope: returned values
  under `data`, failure details under `error`, timing under `durationMs` and batch
  identity under `index`. Scripts using the previous result wrapper must adapt.
- Replace `--sheet` with `--sheet-name` and `--range` with `--range-address`.
  Legacy aliases and optional result-field selection were removed.
- Excel work runs in private workers; status/control paths remain responsive
  while operations are pending. Cancellation stops waiting, not workbook effects.
- Inherited clearer batch/action help, faster absent-service detection, native
  file/URL opening and OneDrive reconnection repairs, invariant named formulas,
  PivotTable labels/totals and a smaller advertised MCP tool catalogue.
- Normal file opens disable macros. Security changes require their dedicated
  approvals; reopening for macro enablement must not discard unsaved work.

See [migration](../../README.md#migrate-from-fork-254) and
[agent workflows](AGENT-WORKFLOWS.md) for exact inputs and limits. The upstream
[2.6.0 changelog](../../CHANGELOG.md#260---2026-10-29) retains individual credits.

### Build and dependency changes

- Pinned extension-packaging `@azure/msal-node` from 6.0.1 to 7.1.0 under Azure
  Identity 4.14.0. Adversarial response-metadata checks failed before the repair
  and passed afterward; real tenant authentication was not performed. This npm
  development dependency is not consumed by the portable .NET executables.
- Updated `Microsoft.Extensions.Logging.Abstractions`,
  `Microsoft.Extensions.DependencyInjection`, `Microsoft.Extensions.ObjectPool`
  and `System.Drawing.Common` from 10.0.13 to 10.0.14.
- Updated `Microsoft.CodeAnalysis.NetAnalyzers` from 10.0.402 to 10.0.403.
- Added pinned `ExcelDna.Interop` 16.0.0 build definitions and retained the
  isolated ADO importer workflow. Build definitions are not new user-facing
  Excel installation requirements.
- Extended package verification with a separate runtime receipt, native XLSM
  smoke checks and CLI batch measurements; expanded protocol cleanup checks.

### Verification

The recorded normal Excel-free suites passed **3,348 cases**. Additional focused
native checks covered audit refusal, cached reads, mixed formatting, VBA source
and result handling, workbook aliases and formula preflight. Extension tests
passed 29 cases, with typecheck, lint, dependency audit, signature checks and
source-channel VSIX package checks. These results do not mean a fork extension
was published to a marketplace.

[Release and downloads](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.0.0)
· [Changes since 2.5.4](https://github.com/Antrakt92/mcp-server-excel/compare/v2.5.4...v3.0.0)
· [Verification](VERIFICATION-3.0.0.md)

## 2.5.4 — 2026-10-25

First stable Antrakt fork release, based on upstream 2.5.3 by Stefan Broenner and
contributors, under the original MIT license.

### Process lifetime and transport

- Shut down the MCP host on stdin EOF, broken stdout or owning-client exit,
  without a competing stdin reader.
- Bound outstanding MCP admission through operation completion and cancellation
  recovery; preserve per-session ownership and response/error behavior.
- Bound CLI connection drain so retained idle clients cannot hold the daemon
  open indefinitely after the response grace period.
- Parse standard MCP result JSON once while retaining text/structured content
  and error compatibility.
- Add fork-specific update checks and portable MCP/CLI downloads.

### Workbook ownership

Preserved upstream ownership rules: attached workbooks remain untouched; unsaved
owned workbooks are handed back open and visible. Cancellation does not authorize
saving, closing, discarding edits or rollback.

### Packaging and verification

Published self-contained Windows x64 MCP and CLI ZIPs, `SHA256SUMS` and
`RELEASE-MANIFEST.json`. The separately labeled legacy 2.4.1 source patch is
archival and must not be applied to the current fork.

Local checks passed: **2,369 Excel-free runtime cases**, 222 ScriptSafety cases,
five packaging guards, four CLI process cases and seven native Excel scenarios.
Actual packaged MCP checks covered protocol initialization, EOF, broken stdout
and owning-client exit.

[Release and downloads](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v2.5.4)
· [Verification](VERIFICATION-2.5.4.md)

## Requirements and verification limits

Portable releases require Windows x64, installed desktop Microsoft Excel and an
interactive desktop. They include the .NET runtime, but not Excel, and are
unsigned. Download the checksum and manifest files from the same release;
3.0.x also provides `VERIFICATION-RESULTS.json`. Follow [recovery](RECOVERY.md)
to verify source revisions and actual executable hashes.

Validation uses local synthetic fixtures. It does not establish customer-workbook
calculation/print equivalence, every Office version, historical RPC/AutoSave
incident resolution or all-day resource stability. Hosted Actions, visible
prompt regressions and LLM evaluations were not run for 3.0.1. These portable
releases do not publish upstream npm, NuGet, extension or registry identities.
Changing binaries does not undo workbook edits.
