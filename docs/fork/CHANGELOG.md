# Antrakt Fork changelog

Release history for [Antrakt92/mcp-server-excel](https://github.com/Antrakt92/mcp-server-excel).
Dates are the original GitHub publication dates, in UTC. Changes are grouped by
the first fork release that contains them. The [upstream changelog](../../CHANGELOG.md)
retains the original project's version numbers and contributor credits.

| Version | Published | Summary |
| --- | --- | --- |
| [3.0.1](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.0.1) | 2026-10-29 | Named-range and chart fixes, idle-resource repair, dependency updates |
| [3.0.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.0.0) | 2026-10-29 | Audit sessions, bounded reads, guarded VBA and CLI migration |
| [2.5.4](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v2.5.4) | 2026-10-25 | First stable fork: process lifetime and portable packages |

## After 3.0.1 — 2026-10-29

These changes are in `main`, after the source used for the 3.0.1 packages.
The published executables and their verification receipts are unchanged.

- The release builder now rejects an executable unless its complete
  `version+commit` stamp matches the selected source commit. The previous check
  accepted a matching version with a wrong or missing commit stamp. A regression
  executes the actual guard against seven metadata cases.
- Reworked the README around requirements, package selection and MCP/CLI setup.
  Added release/license badges and kept the longer download example collapsible.
- Simplified contributor headings and corrected security/recovery links. Download
  verification now explicitly includes the separate runtime verification receipt.
- Updated the repository description, topics and release homepage. Enabled
  automatic deletion of finished PR branches and removed merged working branches.
- Added this dated fork history and expanded the GitHub release descriptions.

[Maintenance PR #12](https://github.com/Antrakt92/mcp-server-excel/pull/12)
contains the builder and presentation changes. Its checks passed: 99 packaging
cases, the seven-case ADO class after correcting the local importer environment,
the remaining 282 ScriptSafety cases, strict documentation build and the 58-page
site audit. Both normal commit hooks passed with zero build warnings/errors.

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
