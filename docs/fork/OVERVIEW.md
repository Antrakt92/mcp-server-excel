# Fork overview

ExcelMCP — Antrakt Fork is maintained at [Antrakt92/mcp-server-excel](https://github.com/Antrakt92/mcp-server-excel). The 3.0.0 release line starts from upstream [v2.6.0](https://github.com/sbroenne/mcp-server-excel/tree/v2.6.0). Fork 2.5.4 was based on upstream 2.5.3. The major version reflects incompatible CLI JSON output and removed options inherited in this upgrade; see [migration](../../README.md#migrate-from-fork-254).

Fork 3.0.1 integrates upstream 2.6.1 and the subsequent chart-label maintenance fix (`efd33740`) while preserving fork lifecycle and workbook fixes. It is a patch release for fixes, dependency updates and guidance; it introduces no new public capability or migration. Published 3.0.0 packages and their verification receipts identify their original 2.6.0 source base.

Fork 4.0.0 retains upstream 2.6.11 and MCP SDK 2.3.1. It adds known-password
VBA project unlocking and verification, exact whitespace-preserving password
validation, and screenshot fixes for normal, frozen and split views. The
catalogue has 445 operations through 64 MCP tools across 33 feature areas. CLI password
actions now require batch JSON on stdin; direct password flags and file batches
containing those actions are refused. See [migration](../../README.md#migrate-password-calls-from-fork-360).

Fork 3.6.0 retained upstream 2.6.11 and MCP SDK 2.3.1. It introduced native VBA project
password configuration through MCP and CLI, with explicit save/reopen and
printable ASCII passwords. Its catalogue has 443 operations through 64 MCP
tools across 33 feature areas. Existing 3.5.1 inputs remained supported.

See the [fork changelog](CHANGELOG.md) for dated release notes, dependency versions,
migration details and changes made after the latest published packages.

Fork 3.1.0 integrates upstream 2.6.4, including chart text formatting, native
PivotCache sharing and close-stage timings. It exposes 436 operations through
63 MCP tools across 33 feature areas. Earlier packages retain their original
source and verification receipts. Downloading a release does not switch an
already running MCP/CLI process.

Fork 3.2.0 integrates upstream 2.6.6 and exposes 437 operations. It adds worksheet
scroll/selection controls and recovery when caller cancellation races with a
completed open/create reply. Recovery asks Excel to show that session without
saving, closing or repeating workbook work. See the [3.2.0 notes](CHANGELOG.md#320--2026-10-31).
Earlier 3.1.0 packages and their verification receipts remain unchanged.

Fork 3.3.0 integrates upstream 2.6.7 and exposes 438 operations. It adds chart
title placement and axis line formatting, protects existing named ranges against
creation aliases, and updates extension packaging dependencies. See the
[3.3.0 notes](CHANGELOG.md#330--2026-11-01).

## Scope

Fork 3.4.1 integrates the accepted upstream numeric-error fix through `edb5eada`.
Value/formula reads and trace previews were already protected in the fork;
the new repair prevents a numeric formula from being reported as a blocked spill.
There are no new dependencies, public contracts or migration steps. See the
[3.4.1 notes](CHANGELOG.md#341--2026-11-02).

Fork 3.4.0 integrates upstream 2.6.8 and its subsequent dependency maintenance.
It restores compact `list-actions` discovery, including session commands, and
removes bundled agent skills while keeping native help and reference pages.
Existing fork workbook safeguards and CLI JSON fixes remain in place. See the
[3.4.0 notes](CHANGELOG.md#340--2026-11-02).

Fork 3.3.1 fixes structured CLI JSON arguments: camelCase fields retain VBA
argument types and range/table sort settings, while PascalCase remains supported.
The upstream base, dependencies and operation counts are unchanged. See the
[3.3.1 notes](CHANGELOG.md#331--2026-11-01).

The fork hardens the MCP process boundary: stdin EOF, terminal stdout failure, owning-client lifetime, a single JSON-RPC stdout stream, and bounded outstanding request admission through cancellation cleanup. It uses native Excel for workbook operations. Current fork additions include persistent audit sessions, bounded range reads and formatting summaries, guarded whole-module replacement and typed VBA invocation. [Agent workflows](AGENT-WORKFLOWS.md) describes their inputs and limits.

CLI connection drain is bounded after shutdown: clients receive the existing response grace, then retained transport pipes are closed after two seconds. This closes connections, not workbook operations. Generated MCP replies reuse the normalized Service result, preserving public content, error classification and telemetry. Screenshot conversion keeps its existing path.

Upstream 2.6.0 isolates Excel work in private workers and defines workbook lifetime and cancellation recovery. Normal shutdown hands unsaved workbooks back open and visible, attached workbooks remain untouched, and timeout/cancellation does not authorize closing, saving or discarding workbooks. The governing design is [ADR-005](../ADR-005-SESSION-AND-FAILURE-OWNERSHIP.md). Older forced-close or automatic-save behavior must not be transplanted from the 2.4.1 patch.

The earlier cumulative 2.4.1 lifecycle patch is historical work. If distributed as a legacy recovery asset, it is labeled separately and is not the recommended current runtime. Its stress results and timings do not establish correctness or performance for the current release.

## Distribution

The MCP and CLI ZIPs are portable Windows x64 distributions with a bundled .NET runtime. Installed desktop Excel and an interactive Windows session are still required. Fork downloads come from this fork's GitHub releases; upstream npm, extension, marketplace and installer links remain upstream distributions unless explicitly identified otherwise.

Each release includes `SHA256SUMS`, `RELEASE-MANIFEST.json` and `VERIFICATION-RESULTS.json`. Preserve them with the downloaded packages. GitHub's automatic source archives provide the tagged source snapshot. See [4.0.0 verification](VERIFICATION-4.0.0.md) and [recovery](RECOVERY.md) for verification and switching versions.

## Build and verify

Use the SDK in `global.json` and PowerShell 7.4 or later. Follow [scripts/AGENTS.md](../../scripts/AGENTS.md), [testing guidance](../../tests/AGENTS.md) and the repository build instructions. The fork packaging entry point is [scripts/Build-ForkRelease.ps1](../../scripts/Build-ForkRelease.ps1); inspect its help for current parameters.

Fork standalone releases use that local builder with explicit version metadata, followed by authenticated GitHub CLI publication. The inherited upstream Actions publication routes remain disabled on this fork. No upstream npm, NuGet, marketplace or registry identity is published by this procedure.

Verification must distinguish compilation and tests without Excel, actual stdio/process lifetime checks, and tests using desktop Excel. Passing protocol tests does not prove workbook behavior, native dialogs, rendering or numerical correctness. A release manifest records what was actually checked and any omitted checks; older receipts do not replace new validation.

For lifecycle changes, include EOF, broken stdout with stdin still open, and parent exit while another writer retains stdin. Warm the same workload before comparing resources. Measure a bounded workload and shutdown; do not advertise an all-day leak guarantee or performance improvement without matching evidence.

## Maintenance boundary

Keep fork changes separate from unrelated upstream behavior. Preserve upstream license attribution, source contracts, calculation behavior and workbook ownership. Public reports and packages must exclude personal configuration, workbook data, private paths, secrets and unsanitized diagnostic output.
