# Antrakt Fork changelog

Release history for [Antrakt92/mcp-server-excel](https://github.com/Antrakt92/mcp-server-excel).
Dates are the original GitHub publication dates, in UTC. Changes are grouped by
the first fork release that contains them. The [upstream changelog](../../CHANGELOG.md)
retains the original project's version numbers and contributor credits.

| Version | Published | Summary |
| --- | --- | --- |
| [5.9.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v5.9.0) | 2026-11-22 | Close history, worker readiness, and VBA readback diagnostics |
| [5.8.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v5.8.0) | 2026-11-21 | VBA operation receipts, save and close diagnostics, and range search fixes |
| [5.7.1](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v5.7.1) | 2026-11-21 | Safer workbook access changes and fixes for worksheet, formatting, VBA and CLI commands |
| [5.7.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v5.7.0) | 2026-11-21 | MCP shutdown diagnostics, Excel recovery and safer workbook edits |
| [5.6.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v5.6.0) | 2026-11-19 | Literal text writes, VBA label warnings and workbook timings |
| [5.5.2](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v5.5.2) | 2026-11-19 | VeryHidden worksheet deletion and NuGet updates |
| [5.5.1](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v5.5.1) | 2026-11-19 | VBA return values, timeout tests and workbook inspection guidance |
| [5.5.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v5.5.0) | 2026-11-18 | Paged visibility, workbook safeguards and Excel session recovery |
| [5.4.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v5.4.0) | 2026-11-17 | Windows MCP bundle, workbook edits, and batch validation |
| [5.3.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v5.3.0) | 2026-11-17 | Saved sheet identities, relationship ownership and parent-style diagnostics |
| [5.2.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v5.2.0) | 2026-11-16 | Workbook package checks and hidden defined names |
| [5.1.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v5.1.0) | 2026-11-16 | Cell inspection and edits, write receipts, and workbook recovery |
| [5.0.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v5.0.0) | 2026-11-12 | Workbook and file safety, faster DAX reads, and recoverable Excel sessions |
| [4.1.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v4.1.0) | 2026-11-05 | Native exit history, clearer failure diagnostics and safer disconnected batches |
| [4.0.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v4.0.0) | 2026-11-04 | Known-password VBA authentication, CLI stdin migration and worksheet screenshot fixes |
| [3.6.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.6.0) | 2026-11-03 | Native VBA project password configuration and verified dialog cleanup |
| [3.5.1](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.5.1) | 2026-11-03 | MCP SDK maintenance and locked video tooling updates |
| [3.5.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.5.0) | 2026-11-03 | MCP batch, custom slicer styles and cancellable CLI startup waits |
| [3.4.1](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.4.1) | 2026-11-02 | Correct numeric formula classification in spill inspection |
| [3.4.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.4.0) | 2026-11-02 | Compact CLI discovery, upstream 2.6.8 and simpler agent packages |
| [3.3.1](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.3.1) | 2026-11-01 | CLI typed JSON arguments and range/table sort settings |
| [3.3.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.3.0) | 2026-11-01 | Chart title placement, axis line formatting and named-range overwrite protection |
| [3.2.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.2.0) | 2026-10-31 | Worksheet scroll/selection controls and completed-session cancellation recovery |
| [3.1.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.1.0) | 2026-10-30 | Chart text formatting, native PivotCache sharing, close timings and MCP response optimization |
| [3.0.1](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.0.1) | 2026-10-29 | Named-range and chart fixes, idle-resource repair, dependency updates |
| [3.0.0](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v3.0.0) | 2026-10-29 | Audit sessions, bounded reads, guarded VBA and CLI migration |
| [2.5.4](https://github.com/Antrakt92/mcp-server-excel/releases/tag/v2.5.4) | 2026-10-25 | First stable fork: process lifetime and portable packages |

## 5.9.0 — 2026-11-22

### Closing and recovery

- `file_read list` and `excelcli session list` now include `recentCloseAttempts`. For a known session, the history shows whether the worker confirmed close, whether the controller released the workbook claims, whether a save during close was confirmed, and any matching native Excel exit. It correlates the exit by PID and exact process creation time.
- A command failure and a failure recording its completion can now be read separately. A successful retry retains `lastFailure` and any earlier worker status diagnostic. The controller keeps up to 16 recent session/process pairs in memory.
- `closeReadiness` distinguishes a known blocker from a failed status check. A previously reported blocker stays visible. When readiness is unconfirmed, an authorized close of the same session rechecks Excel using its existing safeguards. Worker diagnostics include the failing stage, exception type, COM HRESULT, and last matching native observation when available.

### VBA source edits

- If the text read back after an import, module replacement, or procedure replacement differs from the supplied code, `partialState.firstDifferentLine` points to the first difference in the normalized input. `partialState.mismatchKind` reports `codeSpacing` when the difference is limited to recognized code spacing. The response omits the source line. The source fingerprint and strict read-back checks still apply.

The upstream base remains **2.6.13**. Dependency versions are unchanged.

[Full comparison](https://github.com/Antrakt92/mcp-server-excel/compare/v5.8.0...v5.9.0)

## 5.8.0 — 2026-11-21

### VBA and CLI batches

- Give a `vba.invoke` batch entry a fresh UUID as its `operationId` to look up its outcome after a timeout or a lost reply. Use `session operation-result` in the CLI or `file_read` with `operation-result` in MCP. The receipt can include the typed return value, and a duplicate ID is rejected while the receipt is retained. The same controller holds up to 128 receipts. Finished receipts expire after 15 minutes, and results over 8 KiB are omitted.
- Ctrl+C stops `excelcli batch` from starting further commands and preserves the results already written. An interrupted command that may have reached Excel is reported as unconfirmed.

### Saving, closing and discovery

- When an explicit save returns `partialState.save`, it separates Excel's save-event report from the workbook's unsaved state. If Excel reports success before a COM error, the service preserves the original error and skips an automatic retry. Check the saved file when the workbook remains dirty.
- Save As releases a stale destination reservation once Excel confirms that the workbook is still at its original location. If the outcome remains unclear for a workbook that has never been saved, its destination stays reserved until the workbook location is confirmed or the session closes.
- When Excel's readiness or an RPC disconnect leaves closing unconfirmed, the response explains how to check the session and find an open workbook. An RPC disconnect also prompts you to verify the saved file and, when the process identity is available, match it to the same controller's recent exits by PID and creation time. Confirmed exit records remain visible after the workbook session disappears.
- Partial discovery explains how to proceed when another controller or a different ExcelMcp build owns an Excel instance. Sessions already managed by the current connection remain usable.

### Search and data validation

- Range search checks values and formulas when both sources are enabled. Each matching cell is counted once, even with a limited preview.
- Replace First rejects unsupported partial-cell wildcard patterns even when the range has no matches. Escape a literal `*` or `?` with `~`, or use Replace All for wildcard matching.
- Replacing data validation preserves the existing rules when Excel rejects the replacement on the first target cell.

The upstream base remains **2.6.13**. Dependency versions are unchanged.

[Full comparison](https://github.com/Antrakt92/mcp-server-excel/compare/v5.7.1...v5.8.0)

## 5.7.1 — 2026-11-21

### Workbook access

- If Excel keeps the selected workbook in Protected View after an approved edit request, MCP reports that state and keeps the existing session attached.
- MCP and CLI batches stop when Excel cannot confirm an editing, connection or macro-reopen operation. They also stop if Excel fails to restore its opening settings, even when the workbook operation itself completed. The error preserves the original cause and any restoration failures for inspection before the next command.

### Worksheet, VBA and CLI fixes

- Moving a sheet honors `afterSheet` when `beforeSheet` is blank.
- Writing a matrix of number formats requires one rectangular range. A disjoint selection is rejected before any formatting changes. A single format can still be applied to disjoint ranges.
- VBA module deletion matches names case-insensitively, like module reads.
- CLI batches return a shutdown warning in `partialState.shutdownWarning` when the service stops with a warning. The warning no longer causes the batch response to fail JSON parsing.

### Regression coverage

Desktop Excel tests now check that rejected range and table inputs preserve workbook contents, locked files reopen after the lock is released, and closing without saving cannot produce a save receipt on retry. Dependency versions are unchanged.

[Full comparison](https://github.com/Antrakt92/mcp-server-excel/compare/v5.7.0...v5.7.1)

## 5.7.0 — 2026-11-21

Set `EXCELMCP_STDIO_DIAGNOSTICS=1` to record why the MCP server stopped. The optional stderr line gives the first observed reason and the exit code, which helps distinguish closed input, a failed output write, and a parent process exit. Standard output stays reserved for MCP messages.

### Excel recovery

- If Excel exits during a callback, before a new command, or during a readiness check after a timeout, the service can return the observed process identity and exit code. It reports the original COM error when Excel is still running.
- Workbook discovery keeps a worker's newest registration when an earlier registration becomes stale. The CLI daemon handles lost startup connections without treating an incomplete startup as a usable session.

### Workbook operations

- Range and table writes reject JSON numbers that overflow to infinity, as well as nonfinite values, before writing cells or resizing a table. Matrix inputs reject null rows and requests that supply both a nonempty inline matrix and a file. Empty inline arrays still allow file input.
- Regular PivotTable and table slicers apply matching names from a mixed list. If no names match, the existing filter stays in place and the call returns an error. Empty lists still clear filters. Table slicers report the actual number of selected items.
- Screenshot capture accepts one rectangle at a time, including when the range comes from a defined name. A disjoint range is rejected before the active sheet, selection or viewport changes.
- Special Cells skips the per-cell visibility scan when Excel confirms a single row or column is hidden. Workbook path checks use smaller buffers for ordinary paths and still handle long paths.

### Dependencies

This release incorporates upstream **2.6.13**. The portable builds pick up OpenTelemetry PersistentStorage Abstractions and FileSystem **1.1.3 → 1.2.0**, Azure.Core **1.64.0 → 1.65.0**, Polly Core, Extensions and RateLimiting **8.9.0 → 8.9.1**, Nerdbank.MessagePack **1.4.95 → 1.4.104**, Microsoft.NET.StringTools **18.11.0 → 18.12.2**, and System.ClientModel **1.16.0 → 1.17.0**. The affected Microsoft.Bcl and System libraries move from **10.0.14 to 10.0.15**, including System.Text.Json and System.Threading.RateLimiting. Test tooling moves to xunit.analyzers **2.3.0** and coverlet.collector **10.1.1**. The upstream evaluation lockfile was also refreshed.

[Full comparison](https://github.com/Antrakt92/mcp-server-excel/compare/v5.6.0...v5.7.0) · [Upstream 2.6.13](https://github.com/sbroenne/mcp-server-excel/releases/tag/v2.6.13)

## 5.6.0 — 2026-11-19

Bulk range writes and table appends now have a literal text mode. Pass `literal_text: true` in MCP, `--literal-text true` in the CLI, or `literalText: true` in batch JSON. Leading zeros, apostrophes, date-like strings and text beginning with `=` survive saving and reopening. Numbers and Booleans keep their types. Existing calls keep their formula and date interpretation.

- **Full-length cell text.** Formula and error reads work alongside 32,767-character cells. Oversized literal writes fail before changing cells, and overwrite checks identify occupied destinations.
- **Empty tables.** An append with a nonblank first row can succeed when Excel leaves its blank insert row outside the table data. In that state the first row needs at least one value.
- **VBA inspection.** Use `vba_read list` with `include_label_warnings: true` or `excelcli vba list --include-label-warnings` to find labels that may have been intended as calls. Results identify the module, procedure and source line. The scan leaves the project unchanged.
- **Workbook diagnostics.** `file_read test` and `excelcli session test` report `testTimingsMs` for measured phases, including Excel activation, workbook open and cleanup. Failed startup cleanup has a separate field. Controller and worker intervals can overlap.

No dependency versions changed in this release.

[Implementation and verification](https://github.com/Antrakt92/mcp-server-excel/pull/81) · [Full comparison](https://github.com/Antrakt92/mcp-server-excel/compare/v5.5.2...v5.6.0)

## 5.5.2 — 2026-11-19

Deleting a VeryHidden worksheet now works through MCP and the CLI. If Excel refuses or cancels the deletion and the original sheet still exists, the command restores its VeryHidden state before returning the error. This also covers a workbook where the only visible tab is a chart sheet.

### Dependencies

| Package | Previous | New |
| --- | --- | --- |
| ModelContextProtocol | 2.3.1 | 2.4.0 |
| OpenTelemetry, OpenTelemetry.Api, OpenTelemetry.Extensions.Hosting, OpenTelemetry.Exporter.InMemory, OpenTelemetry.Exporter.OpenTelemetryProtocol | 1.20.1 | 1.21.0 |
| Microsoft.Extensions.Caching.Abstractions, Configuration, Configuration.Binder, Configuration.EnvironmentVariables, Diagnostics, Diagnostics.Abstractions, FileProviders.Abstractions, Hosting, Hosting.Abstractions, Logging, Logging.Abstractions, Logging.Configuration, Options, Options.ConfigurationExtensions | 10.0.14 | 10.0.15 |
| System.Drawing.Common | 10.0.14 | 11.0.0 |
| Microsoft.Extensions.Resilience | 10.11.0 | 11.0.0 |
| Microsoft.NET.Test.Sdk (test tooling) | 18.11.0 | 18.12.0 |

The dependency updates come from the upstream project. Existing command arguments and the Windows x64 requirement carry forward.

[Worksheet fix](https://github.com/Antrakt92/mcp-server-excel/pull/78) · [Upstream sync](https://github.com/Antrakt92/mcp-server-excel/pull/79) · [Full comparison](https://github.com/Antrakt92/mcp-server-excel/compare/v5.5.1...v5.5.2)

## 5.5.1 — 2026-11-19

This maintenance release documents how to check a VBA call that returns null,
diagnose a print display issue and recover after a macro times out.

- **VBA function results.** A function in a worksheet module can change the
  workbook and still return null through Excel's `Application.Run`. The guide
  shows when to use an explicit standard-module wrapper and when to read back
  workbook state. New desktop Excel tests verify the return value, print area
  and call count for both entry points.
- **VBA timeouts.** The regression test now saves and reopens its disposable
  workbook through the normal macro approval path. It verifies that Excel can
  finish the macro after the caller times out and that the late workbook
  change can be read before closing the session.
- **Workbook inspection.** The guide shows how to compare displayed cell text
  with its number format, width, wrapping and shrink-to-fit settings. It also
  explains which ActiveX layout details the drawing inventory returns.
- **Dependabot.** NuGet minor and patch updates now use separate groups for
  test, compiler and runtime dependencies. Major updates still open
  individually. Package versions and the upstream 2.6.12 base carry forward.

[Full comparison](https://github.com/Antrakt92/mcp-server-excel/compare/v5.5.0...v5.5.1)

## 5.5.0 — 2026-11-18

Read visibility across large ranges a page at a time, and copy or move sheets between files from the CLI without opening a workbook session. This release also fixes cases where rejected inputs could leave changes behind in Excel.

### What changed

- **Range visibility.** Use `offset` and `limit` with MCP `range_format_read` → `get-visibility`, or `--offset` and `--limit` with `excelcli rangeformat get-visibility`. Start at offset zero and reuse the same range, axis and limit with the returned `nextOffset` until it is absent. `totalDimensions` counts distinct rows or columns across the full scope.
- **Batch responses.** In a local benchmark with a synthetic 2,040,924-byte response, median JSON processing fell from 30.37 ms to 6.84 ms and allocations from 17.7 MB to 4.45 MB. The benchmark measures response processing, including the final JSON write.
- **Saved workbook comparisons.** Named parent styles are indexed by definition. This cuts repeated work when comparing saved files with many styles and preserves the reported matches.

### Workbook changes and exports

- Missing data-validation formulas are rejected before an existing rule is replaced. Creating a table with an unknown style leaves the source cells alone. Adding a table column rejects a header that duplicates an existing one regardless of case, or exceeds Excel's 255-character limit, before inserting the column.
- Adding or editing a hyperlink rejects tooltips longer than 255 characters before changing the cell or link. Replacing manual page breaks preserves the worksheet's print scaling.
- Structured references that do not resolve to a range return a clear failure. A PDF or XPS export still reports success if the file was written and only opening the viewer failed.
- The CLI checks `--output` against the explicit file paths in `workbook save-as`, `workbook save-copy-as`, `workbook export-fixed-format`, `chart export-image`, and both workbooks in `sheet copy-to-file` and `sheet move-to-file`. If a file alias only resolves after the operation, the CLI preserves the artifact and reports that the operation completed and the response file could not be written.

### Excel sessions and command errors

- A delayed readiness error stays with the operation that started the check. When an Excel helper exits, the service records the lost session and releases its file reservations before removing it from the list. It leaves worker records with invalid or mismatched identity metadata untouched.
- When workbook validation opens the file in a temporary Excel instance, it reports whether that process exited. Malformed validation replies stay in the normal error envelope, and explicitly empty enum arguments are rejected instead of receiving a default.
- Concurrent first launches can read the same saved telemetry identity while another process finishes publishing it. Unused VBA parser and service-request wrappers were removed.

### Upgrade notes

- Scopes larger than 10,000 rows or columns now require a page limit. Choose a limit from 1 to 10,000. Smaller scopes can still be read without paging.
- The portable CLI, MCP ZIP and MCPB target Windows x64 and use installed desktop Excel. The packages declare their upstream base from the exact source revision used to build them.

### Dependencies

| Scope | Dependency | Update |
| --- | --- | --- |
| Runtime | `Microsoft.Extensions.DependencyInjection` | 10.0.14 → 11.0.0 |
| Runtime | `Microsoft.Extensions.ObjectPool` | 10.0.14 → 11.0.0 |
| LLM tests | `pytest-asyncio` | 1.4.1.post1 → 1.5.0 |
| Documentation | `mkdocs-material` | 9.8.0 → 9.8.1 |
| GitHub Actions | `github/gh-aw/actions/setup` | pin `34f03fba65c7` → `399a445fe2d8` |

Upstream base remains ExcelMcp 2.6.12. [Full comparison](https://github.com/Antrakt92/mcp-server-excel/compare/v5.4.0...v5.5.0)

## 5.4.0 — 2026-11-17

### Installation

- The Windows x64 `.mcpb` contains the same self-contained server as the MCP ZIP. Clients with a desktop extension installer can launch it directly. Installed desktop Microsoft Excel and an interactive Windows session are required.
- The release includes `server.json` for the fork's MCP Registry identity. The manifest, verification receipt and checksums identify the exact published bundle and binary.
- The README includes a sample workbook and a read-only audit that checks the saved file's SHA-256. Manual MCP and CLI ZIP downloads remain available.

### Workbook behavior

- Cross-file worksheet moves with no explicit destination position now place the worksheet after the final tab, including chart tabs.
- Named-range writes preserve literal text beginning with an equals sign or an apostrophe. Large value previews are omitted while list responses keep the range metadata.
- Table creation without headers keeps the first data row. Structured references escape special characters in column names and emit valid current-row references. Totals apply the selected minimum, maximum or variance function. Invalid totals arguments leave a hidden totals row hidden.
- Saved workbook comparisons resolve relationships that differ only in case, accept ZIP comments containing footer-like bytes and bound memory use when checking the structural report size.

### Commands and recovery

- MCP and CLI reject duplicate JSON properties in a batch before the first operation runs. Individual raw commands also reject duplicate keys.
- Worker shutdown preserves a workbook session opened during controller exit. The CLI retains the original shutdown warning if a second error occurs while stopping.
- The CLI skips telemetry response processing when telemetry is off. A malformed service reply keeps its failure classification through the CLI invocation.

The release retains upstream ExcelMcp 2.6.12. Runtime and development dependency versions are unchanged.

## 5.3.0 — 2026-11-17

Saved workbook comparisons can now explain structural changes after a native Excel save. This helps inspect sheet moves, deleted tabs and renumbered package parts in complex XLSX and XLSM files.

### Structural package diagnostics

- Set `include_structure: true` on MCP `file_read` with action `compare-packages`, or add `--include-structure` to `excelcli session compare-packages`. Batch arguments use `includeStructure`.
- The optional `structure` report includes both saved tab orders, sheet names, `sheetId`, relationship IDs, part names, sheet types and `CodeName`. Defined names include their `localSheetId` and resolved sheet scope, so a moved sheet's filter scope can be checked against the new tab order.
- Relationship records identify the source owner and resolved internal target, including missing owners and targets. Graph-unreachable parts are listed for inspection. External URLs, defined-name formulas and cell values are excluded from the report.
- Named parent styles are compared using their names and fingerprints of the parent definitions, including referenced fonts, fills, borders and number formats. The report identifies shifted parent indices, removed styles, changed definitions and ambiguous candidates.
- Structural diagnostics retain the exact part allowlist. A parent-style match describes the saved definitions and does not establish rendered equivalence or approve rewriting the package. The command inspects saved files without opening Excel.

### Limits and compatibility

The option is off by default. Existing hash-only comparisons and checked `save-copy-as` calls keep their current behavior. Structural metadata, XML depth, style candidate counts and report size have explicit limits. Unsupported leading sheet extensions and ambiguous internal targets return an error. See [saved package comparisons](../reference/workbook.md#compare-saved-workbook-packages) for the limits and result semantics.

The catalogue remains at 449 operations across 64 MCP tools. Runtime and development dependency versions are unchanged from 5.2.0. The portable builds retain upstream ExcelMCP 2.6.12.

## 5.2.0 — 2026-11-16

ExcelMCP 5.2.0 adds package comparison for saved workbooks, a check before publishing native Excel copies, and metadata reads for defined names. The portable builds include upstream ExcelMCP 2.6.12.

### Defined names

- Set `visible: false` when creating a name to keep a persistent marker out of Excel's visible name list. Existing calls still create visible names by default.
- Use `metadata_only: true` with MCP `namedrange_read` or `--metadata-only true` with `excelcli namedrange read`. Look up the exact name to see its visibility, original `RefersTo` formula and resolution status, even when one area contains `#REF!`. Metadata mode skips cell values.

### Workbook packages

- Compare saved XLSX and XLSM files with MCP `file_read` → `compare-packages` or `excelcli session compare-packages`. It reports added, removed and changed parts with SHA-256 hashes. Use `matchesAllowlist` to check an exact list of allowed part names. The command works in batches alongside workbook sessions.
- Comparison rejects overlapping ZIP records and invalid compression or size claims. ZIP timestamp and compression changes alone do not count as workbook part changes.
- Set `package_check` on MCP `workbook` → `save-copy-as`, or pass `--package-check` JSON to the CLI. Give it a saved baseline and exact allowed part names. Excel writes a temporary copy, and ExcelMCP publishes the checked file only if its package differences fit the list. If the check refuses publication, the destination remains intact and the response includes the diff. If another process locks the staged file, the error gives its path.
- For a standard VBA workbook, an allowlist containing only `xl/vbaProject.bin` will reject a copy if Excel also rewrites a worksheet or any other part. The check covers package bytes. VBA behavior, references and signatures still require their own verification.

The catalogue now has 449 operations across 64 MCP tools. Runtime dependency versions are unchanged.

## 5.1.0 — 2026-11-16

This release adds two range commands for inspecting and editing specific cells, and operation receipts for checking whether a save or table append completed after a timeout. It includes upstream ExcelMCP 2.6.12.

### Range inspection and edits

- `inspect-cells` (MCP `range_read`, CLI `range inspect-cells`) reports values, formulas and displayed text for selected cells.
- `range.patch-cells` updates only the specified cells. Provide expected cell contents to detect conflicting edits before writing. The result identifies cells that need inspection if Excel interrupts a write.
- Merging cells checks whether the merge would discard existing contents. To discard them deliberately, set `overwrite_policy` to `allow` in MCP or use `--overwrite-policy allow` in the CLI.
- CSV writes and appends preserve quoted commas, escaped quotes and embedded newlines. Malformed input reports its source line before any cells are written. A validation formula longer than Excel’s 255-character limit is rejected before replacing the current rule.

### Saves and recovery

- After a timeout or cancellation, look up the receipt before retrying a save or table append. Give the original batch entry an `operationId` UUID and use `file_read` → `operation-result` in MCP or `excelcli session operation-result` in the CLI. Save status distinguishes a confirmed save from a workbook that still needs inspection.
- Timed-out Save As requests release an unused destination once Excel confirms the workbook stayed at its original path. Session status also reports the verified workbook path when available.
- Stale discovery results and alternate hard-link paths are checked before opening a workbook. Workers retry temporary Windows access errors when checking their controller so they can still recover or reveal retained workbooks when its exit is confirmed.

### Worksheets and bounded reads

- Copying a worksheet next to hidden tabs renames the new copy. Moving a worksheet to the end puts it after the last tab, including chart sheets. Unprotecting a worksheet through MCP or CLI no longer opens a password prompt; pass the password explicitly when the worksheet requires one.
- Named-range, table and PivotTable reads check the size of the requested result before returning it. Power Query evaluation reads in blocks, checks its response budget, and cleans up temporary query objects. To preview a large query, use `Table.FirstN` or `Table.SelectColumns`.

### CLI and dependencies

- CLI batch commands accept mixed-case names. VBA password actions and aliases enforce the stdin-only policy consistently. The plugin launcher supports Windows Node installations with `npx.exe`, including Volta, and preserves quoted arguments through both executable and CMD shims.
- The Python evaluation lockfile updates `opentelemetry-api` from 1.45.1 to 1.45.2 to address [GHSA-mf5v-w3w3-f6cg](https://github.com/advisories/GHSA-mf5v-w3w3-f6cg). Evaluation tests run without the retired generated skills. The portable .NET runtime packages are unchanged.

## 5.0.0 — 2026-11-12

This release fixes cases where ExcelMCP could lose a worksheet or a query load, change the wrong named object, or say a workbook was saved when the save was unconfirmed. It includes upstream ExcelMCP 2.6.12.

### Upgrade notes

- Use absolute Windows paths for MCP arguments that refer to local input files. CLI code, data, formula, query, image, and theme input paths resolve from the directory where the command or batch was submitted. Inline JSON and query text remain supported.
- Split MCP batches into groups of at most 128 commands. When accumulated results reach 8 MiB and more commands remain, the batch keeps the result that reached the limit and stops before the next command. Check `stopReason`, `stoppedEarly`, `executedCount`, and the per-command results before resuming.
- Power Query trace lists return up to 512 matches. Large scans report the limit they reached. Use a newer `since` or a process filter to narrow the scan. The list has no pagination for older matches. Incomplete summaries omit the continuation cursor.
- If a chart or PivotTable name is duplicated, supply a worksheet selector where available. MCP uses `chart_sheet_name` for charts, `sheet_name` for PivotTables, and `pivot_sheet_name` for slicer PivotTable membership. CLI uses the corresponding kebab-case option. Rename the object before using an action that has no selector.
- The default CLI pipe now belongs to the current Windows sign-in session. Close and save work through the matching installation before switching versions. A same-session legacy daemon can be reached explicitly with its old pipe name when needed.

### Workbooks and files

- A cross-file worksheet move saves the destination before removing and saving the source. If the source save then fails, the error reports that a copy may exist in both files.
- A worksheet copied next to a chart tab is renamed using the full tab order. The existing worksheet and chart keep their names.
- Save and close persists edits that arrive after the first save. If Excel vetoes the final save or cannot safely close a workbook with multiple windows, the workbook remains available for inspection. Retrying a closed session with save enabled requires an explicit save receipt.
- New workbooks, workbook copies, PDF and XPS exports, and chart images use a staging file before publication. With overwrite disabled, a file created by another process remains intact. If creation publishes the workbook but Excel cannot reopen it, the error gives the published path and confirms that no session was established. An output already in Excel can still publish after the caller times out. Inspect the session and destination before retrying.
- Live Save As checks the destination again immediately before handing off to Excel. Excel's native Save As still has a narrow race with outside writers. Use Save Copy As with overwrite disabled for an atomic same-format snapshot.
- If opening VBA code moves a reopened workbook to another file, the original session is quarantined. Excel retains the live workbook, and session status identifies its verified current location.

### Ranges, tables, queries and reports

- First-only range replacement changes one match. Partial first-only replacements accept literal text and escaped wildcards. Use whole-cell matching for wildcard patterns. Find can search Excel notes and threaded comments and counts a matching cell once.
- Table append checks the full expansion area and preserves totals and neighboring cells. Table reads keep an existing single blank row, and unfiltered reads skip row visibility calls.
- Power Query checks the requested destination before removing earlier loads. Listing queries scans worksheet and Data Model destinations once per request. ExcelMCP serializes its native-query approval reads, writes and verification. It detects external archive changes where possible and keeps displaced bytes for recovery.
- Chart and PivotTable operations reject duplicate workbook-wide names. Supported operations can select a worksheet explicitly. Slicer membership is checked against the selected PivotTable.
- Ordinary PivotTable filters validate all member names before changing the current filter and return what Excel applied. Single-PivotTable refresh honors the requested wait deadline and reports a native refusal.
- Upstream 2.6.12 adds readable blocking-dialog details to errors, improves workbook and Power Query refresh-all reporting for separate PivotCaches, and includes a short Excel CLI quick-start in the upstream plugin.

### Performance

- DAX and DMV rows are read in bounded ADO batches. On this Windows machine, a warm DAX query returning 1,500 rows by four columns took a median of 8.48 seconds on baseline `e79838b8` and 31.0 milliseconds with batching, about 273 times faster for that query. Each build was measured 14 times with the same synthetic workbook and query.
- MCP initialization fell from 5,125 to 2,419 bytes in an isolated stdio comparison of the guidance change. Both runs listed 64 tools and returned equally sized discovery responses.

### Service and MCP

- Batches stop on uncertain completion and retain the original response for recovery. Partial MCP session lists use `workbook_session_id` consistently.
- A cancelled MCP call can return promptly while its worker finishes. Its admission slot remains held until then, and local status reads have reserved capacity.
- The controller preserves exact worker recovery and exit records, reclaims dead utility slots, and reconciles cancellation around worker claims without releasing an unconfirmed workbook.

### Dependencies and build

The portable runtime uses the same .NET package versions as 4.1.0. Upstream updated the Actions pins for .NET setup 6.0.0 → 6.1.0, CodeQL init/analyze 4.39.1 → 4.39.3, and Agentic Workflows setup 1.0.1 → 1.1.0. Test-only Python pins moved to pytest-asyncio 1.4.1.post1 and pytest-skill-engineering 1.0.8. The fork's CI mappings now cover the MCP examples, briefing assets, and synthetic OLAP helper. Upstream publishing workflows are scoped to the upstream repository.

## 4.1.0 — 2026-11-05

ExcelMCP 4.1.0 adds clearer evidence for Excel automation failures and keeps recent process exits available after their sessions disappear. It also stops dependent batch commands when an automation disconnect leaves the previous operation's outcome uncertain.

### New diagnostics

- Disconnect responses can include `partialState.diagnostics.nativeProcess`, captured through a retained process handle checked against the exact PID and creation time. `state` distinguishes `Alive`, `Exited` and `Indeterminate`; `exitCode` and `nativeError` remain separate. An unavailable exit code is `null`, never an assumed zero.
- `session.list` includes `recentExcelExits` in both complete results and partial results. It retains up to 16 distinct process identities for the current controller, including exits whose sessions have already been removed. Repeated observations do not consume extra entries.
- Public `startedAtUtcFileTime` values in these diagnostics are decimal strings. Preserve the digits when comparing identities; converting them to JavaScript numbers can lose precision. Native exit codes remain unsigned JSON numbers.
- `partialState.diagnostics.execution` separates an operation failure from a later visibility-snapshot failure. It reports `callbackReturned` and the operation/snapshot exception types and HRESULTs when those facts are available. `complete:false` marks an incomplete or ambiguous exception chain. Existing error causes, cancellation tokens and observed VBA state are preserved.

### Fixes

- CLI and MCP batches stop after a retained `ExcelDisconnected` result, including batches configured to continue after ordinary errors. A callback having returned does not make it safe to start the next command after a disconnect.
- Native exit observations are recorded before process ownership is released. A failed diagnostic write leaves teardown retryable. Closing an already exited normal workbook without saving can finish without another COM readiness probe; this shortcut does not report a requested save as successful.
- Repeated teardown reuses an identical observation that the worker has already persisted. An unnecessary second write can no longer turn completed cleanup into a new failure. A richer observation still requires a successful write before replacing the stored state.
- The controller can recover a newer persisted exit observation when its known worker dies before confirming release or delivering its final response. Recovery checks the worker, controller, generation, build, pipe and known Excel identity. It retains diagnostics without adopting a process, recreating a session or acknowledging an unconfirmed release.
- Disconnect handling captures available native evidence before dead-session cleanup can remove it. A process whose exit cannot be confirmed is described as uncertain rather than reported as still running.
- Locked VBA source errors now point to `vba.unlock-project` and the CLI batch-stdin password route instead of saying that projects cannot be unlocked.

### MCP scalar arguments

The server already accepts native JSON numbers and Booleans in `vba` invocation arguments. Some clients present `arguments[].value` as string-only. The workflow guide now explains how to use MCP `batch` action `run` with a nested `vba.invoke` entry to preserve the intended types. Do not quote numbers or Booleans to satisfy that client presentation. New protocol regressions cover discovery, direct invocation and batch forwarding.

### Maintenance

Removed unused result models and legacy version/progress reporters. The generators now share their word-separator conversion while preserving existing CLI kebab-case and MCP snake_case names. Historical 4.0.0 dependency notes were corrected to identify the existing transitive tensor dependency accurately.

Dependencies are unchanged from 4.0.0: upstream ExcelMCP **2.6.11**, MCP SDK **2.3.1**, `System.Numerics.Tensors` **10.0.14**, and build SDK **10.0.403**. The catalogue remains **445 operations**, **64 MCP tools** and **33 feature areas**.

### Upgrading and limits

Existing 4.0.0 command inputs remain supported. The password workflow still requires serialized CLI batch JSON on stdin. Reconnect the MCP client and use the matching CLI installation after switching packages.

Exit history belongs to the current controller and is cleared when that controller restarts. Exit code zero confirms normal native process exit; it does not prove that workbook changes were saved. Cancellation, timeout and a lost connection do not undo edits or authorize a retry.

The historical RPC failures still have no established cause. Bounded synthetic form/timer checks passed without reproducing those incidents. Preserve the incident records and capture the new diagnostics if a failure recurs.

See `VERIFICATION-RESULTS.json` in this release for the checks completed against the published packages and their source revision. Private workbook diagnostics are not included in release assets.

[Full changes since 4.0.0](https://github.com/Antrakt92/mcp-server-excel/compare/v4.0.0...v4.1.0)

## 4.0.0 — 2026-11-04

Adds known-password VBA authentication and fixes worksheet screenshot capture
in normal, frozen and split views. This is a major release because existing
CLI password calls must migrate to batch JSON on stdin.

### Added

- `vba.unlock-project` opens current-session access with a known password.
  `vba.verify-password` checks a real locked-project challenge; accepted input
  also opens current-session access. Neither action removes saved protection,
  saves, closes or reopens the workbook.
- Verification reports recognized rejection as `passwordVerified:false`.
  Unlocking with a wrong password returns `Permissions`. For an already
  accessible project, unlocking reports that no authentication occurred;
  verification returns `Prerequisite` rather than claiming a password match.
- Errors retain observed authentication and cleanup state. Unknown outcomes
  are not guessed mismatches, and cancellation does not undo access already
  granted. Inspect that state before retrying.

### Fixed

- Controller status publishes workbook paths after recording their ownership
  reservations. This closes a race that could admit a duplicate open during
  session adoption or rename.
- Session listing reports worker errors that cannot be shown on a known session
  row. Cleanup retains reservations still owned by newer worker state.
- Discovery and first Save As of an attached unsaved workbook keep names such
  as `Book1` out of the file-reservation map. This prevents path-validation
  errors from interrupting cleanup or masking an already completed save.
- Normal session shutdown confirms that the exact owned Excel process has
  exited before reporting completed cleanup. An unconfirmed exit keeps the
  process tracked and the session quarantined; closing again after confirmed
  exit reconciles cleanup without saving or closing the workbook a second time.
- Range captures in frozen and split views use the correct worksheet pixels
  and preserve pane layout, zoom and actual visible origins. Maximized windows
  retain their window state. Oversized cells, empty visible ranges and
  incomplete crops fail explicitly instead of returning clipped success.
- Captures refresh the current worksheet view before taking the final image,
  preventing stale viewport pixels after scrolling or tiling. This covers
  offscreen styled cells and embedded charts in normal, frozen and split views.
- Password validation now preserves nonempty whitespace-only passwords instead
  of treating them as missing. Missing, null and empty values remain invalid;
  ordinary required text inputs retain their existing whitespace checks.
- Password-action help includes each action's description, canonical batch
  command, argument names and workbook selection. Batch parse errors retain
  type and numeric position without echoing input tokens.
- Historical protection guidance now accurately says unsupported Unicode input
  is rejected before protection changes.

### Dependencies

- Updated the MCP server's existing `System.Numerics.Tensors` dependency from
  **10.0.13 to 10.0.14** through central transitive pinning. No direct project
  dependency was added; MCP SDK remains **2.3.1**.

### CLI migration

`protect-project`, `unlock-project` and `verify-password` require batch JSON on
stdin. Direct password flags and batch input files are refused before connecting;
a file batch containing any password action runs no commands. Ordinary file
batches remain supported. MCP tool inputs remain available.

With the target `$sessionId` already selected, use PowerShell 7:

```powershell
$password = Read-Host 'Known VBA project password' -MaskInput
@{
    command = 'vba.unlock-project'
    sessionId = $sessionId
    args = @{ password = $password }
} | ConvertTo-Json -Depth 4 -Compress | excelcli batch --input -
Remove-Variable password
```

For password configuration, use command `vba.protect-project` with the new value
in `args.password`. Keep passwords out of arguments, files and logs. Omitting
`--input` also reads stdin; a JSON serializer preserves spaces, quotes and
backslashes.

### Limits and verification

Authentication requires an interactive desktop, a normal writable owned
session, printable ASCII input and recognized English native dialogs. Audit,
read-only and attached sessions are refused. Unknown dialogs are not answered.
The action does not enable macros or change Trust Center settings. Native
length overflow is refused; exact maximum-length acceptance across Office
versions is not established.

Screenshot capture briefly brings Excel to the foreground and requires an
interactive desktop. Hidden or minimized Excel
windows are temporarily shown and their original visibility and window state
are restored. Native freeze-divider lines may remain visible. Existing large-range
limits still apply, with top-left truncation reported in the result message.

Unknown/cancellation transport tests do not establish native failure injection
or every cleanup/cancellation phase. See [4.0.0 verification](VERIFICATION-4.0.0.md).

The upstream baseline remains 2.6.11
and MCP SDK remains 2.3.1. The catalogue contains 445 operations through 64 MCP
tools across 33 feature areas.

## 3.6.0 — 2026-11-03

### Added

- `vba.protect-project` sets a password and enables the VBA project's viewing
  lock through Excel's own Project Properties dialog. Available through MCP
  `vba` and CLI batch stdin. Saving and reopening remain explicit operations;
  the result reports both as required before persisted locking is checked.
- The driver verifies the selected project's name and file before each dialog,
  checks the lock option separately, and restores the previous editor project
  and visibility after cleanup. Failures retain the observed partial state.
- Native acceptance covers save/reopen, correct and incorrect passwords,
  whitespace/case/escape differences, source and cell preservation, another
  project in the same Excel, and owned process shutdown.

### Fixed

- Read-only workbook write refusals now report the `Permissions` category.

### Upstream

Integrates the [2.6.11 release metadata](https://github.com/sbroenne/mcp-server-excel/commit/32174cbf6fd6c83af6a39e10722da5eb3dcbff7c).
Its CLI startup-cancellation change was already included in fork 3.5.0; this
sync adds no new runtime changes or dependency updates.

### Limits

Passwords use printable ASCII. VBA's project-specific legacy encoding can map
different Unicode passwords to the same native value, so non-ASCII input is
rejected before protection changes. Existing locked projects and audit
sessions are rejected. This action does not unlock projects, encrypt files,
change Trust Center settings or save pending work automatically.

No dependencies were added or upgraded. The upstream baseline is 2.6.11.
The source catalogue has 443 operations through 64 MCP tools.

## 3.5.1 — 2026-11-03

Updates the MCP SDK and puts both video projects under weekly Dependabot
monitoring. MCP tools, CLI commands and workbook behaviour are unchanged.

### Dependencies

| Dependency | Previous | New | Scope |
| --- | --- | --- | --- |
| ModelContextProtocol | 2.3.0 | 2.3.1 | MCP runtime |
| HyperFrames, briefing video | 0.8.184 | 0.8.223 | Development only |
| undici, HyperFrames override | 7.25.0 introduced by the update | 7.30.1 | Development only; removes audit findings in the new tool graph |
| esbuild, briefing video | 0.25.12 | 0.28.2 | Transitive video tooling, including platform packages |
| acorn, briefing video | Not previously included | 8.19.0 | Transitive video tooling |

The SDK patch hardens authenticated HTTP sessions in the upstream SDK. These
portable packages use stdio; this release does not claim that the HTTP issue
was reachable through our server. See the [SDK release notes](https://github.com/modelcontextprotocol/csharp-sdk/releases/tag/v2.3.1).

### Maintenance

- Dependabot now checks the introduction and briefing video projects weekly.
- Video scripts use locally installed, lockfile-pinned HyperFrames instead of
  fetching a CLI through `npx`. Use `npm ci` before running them. The introduction
  video remains on 0.8.214.
- Briefing tooling pins the patched undici release within major version 7.
  The initial HyperFrames update produced audit findings; the merged graph
  reported zero vulnerabilities on 2026-11-03.

### Compatibility and verification

No migration is required. The SDK remains 10.0.403, the upstream product baseline
is 2.6.10, and the catalogue remains 442 operations and 64 MCP tools.
The [verification guide](VERIFICATION-3.5.1.md) defines package and installation
checks; release assets record actual results and source hashes.

The briefing's original footage is intentionally absent from Git. Its build
and check commands stop on missing clips, so full video rendering is not claimed.
This does not affect the portable MCP/CLI packages. No new performance claim or
resolution of historical RPC/StatusBar observations is made.

## 3.5.0 — 2026-11-03

Agents can send a known sequence of MCP operations in one `batch` call and
create workbook-owned slicer styles. This release integrates upstream 2.6.9,
2.6.10 and the subsequent CLI startup-lock fix, while preserving the fork's
workbook and lifecycle safeguards.

### Added

- MCP `batch` accepts an optional default workbook session, returns a result for
  each attempted command and supports `stop_on_error`. Timeouts, cancellation
  and unknown completion stop later commands. Earlier edits are not rolled back.
- Batch retains audit-only sessions, bounded range pages and supported session
  identifier aliases, including nested command arguments.
- Create a slicer style by copying an existing style, edit button and header
  colours, borders and bold text, then apply it without changing layout or
  filters. List available styles through either MCP or `excelcli`.
- The catalogue grows from 438 to 442 operations and from 63 to 64 MCP tools;
  the 33 feature areas are unchanged.

### Fixed and maintained

- Cancelling CLI startup while another caller holds the startup mutex now
  interrupts the wait promptly. A cancelled waiter does not start a daemon or
  run the protected action; acquired and abandoned locks are still released.
- Include upstream's packaging follow-ups after bundled skills were removed.
  Optional reference documentation remains available; product packages do not
  install agent skills.
- Integrate the accepted upstream histories of our named-range overwrite
  protection and camelCase CLI sort-option fixes. Both fixes were already in
  3.4.1; they are not new fixes in this release.

### Dependencies and compatibility

Runtime and development dependency graphs are unchanged from 3.4.1. The SDK
remains 10.0.403. Upstream product metadata advances from 2.6.8 to 2.6.10; this
is separate from the fork's portable version. No migration is required.

### Verification and limits

[The verification guide](VERIFICATION-3.5.0.md) defines the final-source,
packaged MCP/CLI and native Excel gates. Release assets include the exact source
identity, package hashes and machine-readable results. Historical RPC failure
causes and Excel StatusBar behaviour are not claimed fixed. No new performance
percentage is claimed for this release.

## 3.4.1 — 2026-11-02

A formula returning `-2146826243` could be reported as a blocked `#SPILL!`
because its numeric value matches Excel's COM error code. Spill inspection now
checks the cell's native error state before classifying it. Actual spill errors
keep their existing diagnostics; cell values and formulas are not changed.

This integrates accepted upstream [PR #1285](https://github.com/sbroenne/mcp-server-excel/pull/1285)
through `edb5eada`. The fork already made this distinction for value/formula
reads and trace previews. Those paths retain their bounded reads and budgets;
they now share the same native error helper with spill inspection.

The baseline native group passed 17 cases and failed the numeric-spill case.
After the repair all 18 cases passed, along with 95 existing value, formula,
trace and spill checks. Discovery and execution matched; none were skipped.
These integration results are separate from the final release checks recorded
in `VERIFICATION-RESULTS.json`.

Dependencies, 438 operations, 63 MCP tools and 33 feature areas are unchanged.
All 3.4.0 discovery and packaging changes remain included. No migration is
required. Historical COM/RPC failures remain unresolved and no new performance
claim is made. See [3.4.1 verification](VERIFICATION-3.4.1.md).

## 3.4.0 — 2026-11-02

Agents can discover exact command names with `excelcli list-actions` without
starting Excel. The compact JSON includes session actions as well as workbook
commands. This release integrates upstream 2.6.8 and subsequent maintenance
through `426bae6f`, preserving the fork's 438 operations and existing fixes.

- CLI/plugin/extension packages no longer bundle discovery or report-formatting
  skills. Native `--help`, MCP schemas, launchers and optional reference pages
  remain available. Personal skills installed separately are unaffected.
- The fork retains camelCase CLI list arguments, named-range overwrite
  protection, bounded reads, guarded VBA edits and process-ownership checks.
- Diagnostic guidance now distinguishes pending work, macro errors, COM
  transport failures, native process exit and cleanup. A passing rerun does not
  establish the cause of an intermittent failure.
- Source-tag plugin rebuilding supports both older tags with skill generation
  and newer tags without it.

### Dependencies and tooling

| Component | Previous fork | This release | Scope |
| --- | --- | --- | --- |
| Upstream base | 2.6.7 | 2.6.8 plus maintenance | CLI discovery, packaging and guidance |
| pytest | 9.1.1 | 9.2.0 | Optional agent-evaluation tooling |
| pytest-skill-engineering | 1.0.3 | 1.0.6 | Optional agent-evaluation tooling |
| github-copilot-sdk | 1.0.24 | 1.0.28 | Transitive evaluation dependency |
| `github/gh-aw/actions/setup` | `32ecafbc` | `0627fdbe` (v1.0.1) | Inherited issue-triage workflow |

The fork already used .NET SDK 10.0.403, mkdocs-material 9.8.0 and
pytest-asyncio 1.4.1; they are retained, not new upgrades. NuGet runtime packages
and the npm dependency graph are unchanged. Upstream npm/extension metadata is
2.6.8; the portable fork packages are versioned 3.4.0.

No workbook or CLI JSON migration is required. Use native discovery/help in
place of the removed bundled skills. Historical COM/RPC failures remain under
investigation; this release does not claim to fix them. The three recent
workbook-specific repairs are separate local deliveries.

See [3.4.0 verification](VERIFICATION-3.4.0.md) and the release's
`VERIFICATION-RESULTS.json` for actual checks and executable identities. No new
performance claim is made.

## 3.3.1 — 2026-11-01

CLI commands now retain camelCase fields in structured JSON options. For example,
`--arguments '[{"type":"number","value":2.5}]'` previously reached the Service
with an empty argument type. The same parser could lose sort column selections
and turn `ascending:false` into its default value.

- Typed VBA arguments preserve number, Boolean, string, null and missing values.
- Range and table sorting preserve the specified columns, their order and
  descending settings. Existing PascalCase field names remain supported.
- Batch JSON and MCP dispatch are unchanged. No migration is required.
- Contributor guidance now clarifies Excel process ownership and matching source,
  test and worker builds during native verification.

The upstream base remains 2.6.7, with 438 operations, 63 MCP tools and 33 feature
areas. Dependency versions are unchanged. Three CodeQL alerts in the inherited
issue-triage workflow were reviewed as intentional data flows and dismissed as
false positives; this release does not claim a security code fix for them.

See the [3.3.1 verification route](VERIFICATION-3.3.1.md) and the release's
`VERIFICATION-RESULTS.json` for fresh checks of the packaged executables. Earlier
timings and investigation of historical native failures remain separate evidence;
this patch makes no performance or crash-resolution claim.

## 3.3.0 — 2026-11-01

Position chart titles, format axis lines and create named ranges without changing
an existing definition by mistake. Based on upstream 2.6.7, with 438 operations
through 63 MCP tools across 33 feature areas.

### Chart titles and axis lines

- `chart_config` / `chartconfig` `set-title` accepts `left`, `top` and
  `alignment`. Coordinates are points from the chart's own edges; alignment
  controls text inside the title box. Omitted settings are preserved. An empty
  title hides it and cannot be combined with layout settings.
- The new `set-axis-format` action sets an existing axis line's colour, weight
  and visibility. It supports primary and secondary category/value axes. Changing
  colour or weight alone keeps the current visibility; tick labels and gridlines
  are preserved.
- `set-area-format` also accepts `line_visible` / `--line-visible` for the chart
  or plot-area border. Use `false` to hide the border without changing its fill.
- These controls work with ordinary charts and PivotCharts. See the
  [chart guide](../reference/chart.md#title-placement-and-chart-lines) for inputs
  and the distinction between title placement and alignment.

### Named-range creation

Creating an existing name could silently replace its definition with `=0` before
Excel rejected a formula. The guard now recognises Excel's aliases: surrounding
ASCII spaces, quoted worksheet scopes, active-sheet local names and equivalent
Unicode spellings such as `Straße` / `STRASSE`. Hidden names are included. A
rejected duplicate preserves its formula and visibility, including after reopen.

New names are created with their final name, then assigned through Excel's
English R1C1 formula interface. This avoids temporary-name residue and localised
formula assignment. Testing on an actual German Excel configuration remains a
separate check; changing only the caller culture does not establish that coverage.
If formula assignment fails after a new name was
created, that new name remains for inspection; the error reports the partial
state. Existing names are not changed or deleted to recover from the failure.

### Dependencies and tooling

| Dependency | Previous | New | Scope |
| --- | --- | --- | --- |
| `minimatch` under VSCE | 10.2.6 | 11.0.3 | Extension packaging; scoped override for [GHSA-g3hv-75wg-m249](https://github.com/advisories/GHSA-g3hv-75wg-m249) |
| `@isaacs/nonrecursive-glob-match` | Not present | 1.0.3 | Transitive dependency of the patched glob matcher |
| `@types/vscode` | 1.140.0 | 1.144.0 | Extension development types; minimum VS Code version unchanged |
| `@changesets/format` | 0.1.2 | 0.1.3 | Release-note tooling |
| Introduction video HyperFrames CLI | 0.8.184 | 0.8.214 | Video tooling; separately audited `npx` dependency graph |

The minimatch advisory concerns optimisation level 2. VSCE's current matching
path uses the default level 1; the update removes the affected component without
claiming a demonstrated exploit through that packaging command. Regression tests
cover preprocessing limits, exclusions and ignore-file negation. The extension's
packaged file list is unchanged.

Two unused NuGet pins referenced versions that were never published:
`System.Net.Http` 4.3.5 and `System.Text.RegularExpressions` 4.3.2. They are removed;
neither package was in the resolved dependency graph. Resolved .NET dependencies
and the pinned .NET SDK 10.0.403 are unchanged.

The root and extension npm audits and all-project NuGet audit reported no
vulnerabilities on 2026-11-01. HyperFrames 0.8.226 was rejected because its resolved
graph contains vulnerable Undici 7.25.0. Other deferred tooling updates and their
remaining checks are recorded in [BUILD.md](BUILD.md#deferred-tooling-updates).

### Upstream, upgrade and verification

Includes upstream's chart controls and the named-range regional-formula and
error-reporting fixes. Upstream accepted our ServiceClient cancellation/timeout
changes in [#1280](https://github.com/sbroenne/mcp-server-excel/pull/1280); the fork
already carried that behavior. Our named-range follow-up remains under review
in [#1279](https://github.com/sbroenne/mcp-server-excel/pull/1279).

Use the 3.3.0 portable MCP/CLI packages, verify their checksums and reconnect MCP
clients or restart the relevant CLI service. No breaking migration is required.
Upstream npm, extension and registry metadata is 2.6.7; this fork release publishes
the portable Windows x64 packages only.

See [3.3.0 verification](VERIFICATION-3.3.0.md) and the release's
`VERIFICATION-RESULTS.json` for exact source, protocol and native checks against
the shipped binaries. This release makes no new performance claim. Synthetic
fixtures do not establish customer calculation or print equivalence; portable
packages remain unsigned.

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
