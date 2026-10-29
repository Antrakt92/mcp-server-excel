# Agent workflows

These workflows describe the development branch. Discover the running tool
schema or CLI help before using an option with an older installed release.

## Persistent inspection

For an extended inspection, open with MCP `file` action `open` and
`read_only:true`, or CLI `session open <FILE> --read-only`. Batch JSON uses
`readOnly:true`. Reuse the returned session; do not reopen it for each read.

The response and session list report `auditOnly:true`. Excel opens the file
read-only with opening macros and external-link updates disabled. The Service
refuses workbook mutations, explicit calculation, macro execution, security
changes and saving before dispatching them to Excel. Close an unchanged owned
audit with `save:false`. Explicit close discards unsaved interactive edits; resolve
their owner and handover/save policy first if they appeared. Editing requires a separately authorized normal
session after the audit session closes.

This is a restriction on the agent API, not a lock against interactive Excel
edits or a guarantee that Excel never calculates during opening. The audit flag
survives native Save As and path reconciliation. Normal shutdown preserves
unsaved interactive edits through the ordinary handover policy; it does not
save or discard them automatically.

## Interrupted work

Cancellation and timeout stop the caller's wait. Inspect the selected session's
pending operation and late outcome before issuing another request. Do not retry
a mutation merely because its response was interrupted. Recovery and attachment
retain the workbook's actual state; see [recovery](RECOVERY.md).

## VBA source and function results

Read a module with `vba_read` action `view` before a full replacement. Pass its
returned `sourceHash` as `expected_source_hash` to `vba` action `replace-module`,
with `module_name` and the new `vba_code` or `vba_code_file`. CLI uses
`vba replace-module --expected-source-hash` and `--vba-code-file`.

A stale fingerprint refuses the edit before deletion. Successful edits read
back the actual stored source. A failed edit reports its phase, observed module
identity, fingerprint and line count when available, plus independent read-back
errors. It does not automatically remove a created module or restore old code.
Inspect that state before explicitly authorizing a recovery replacement.

For a function result, use `vba` action `invoke`; the existing `run` action
continues to discard return values. `arguments` accepts typed objects with
`type` set to `string`, `number`, `boolean`, `null` or `missing`, and a scalar
`value` where required. At most 30 arguments are supported. The result reports
`returnKind` and either `value` or rectangular `values`. Macro execution still
requires the existing security approval; invoking does not enable macros.
Returned arrays are limited to 10,000 elements and results have a conservative
4 MiB transport budget. If return conversion is refused, execution has already
completed: inspect its workbook effects before considering another invocation.

## Bounded workbook reads

Use `range_read` action `get-used-range-info` for native bounds without values;
formatting-only cells can inflate UsedRange. Fix a rectangular scope and call
`get-values-page` with zero-based `row_offset`/`column_offset` and at most 10,000
`page_rows * page_columns`. Follow the returned next offsets until `hasMore` is
false. Separate calls are not snapshot isolation. CLI options use hyphens; batch
JSON uses camelCase.

Eager value/formula reads cap geometry at 1,000,000 cells, and all value/formula
responses have a conservative 4 MiB transport budget. Long text can require much
smaller ranges. Oversized reads fail without returning truncated values; reduce
the explicit scope instead of retrying the same request. Native arrays are read
in tiles of at most 1,000 cells. Reads preserve cached results without requesting
calculation. Excel errors retain canonical representation and actual cell
addresses; ordinary numeric constants matching COM error codes remain numbers.

Use `range_format_read` action `get-format-summary` for aggregate stored,
displayed or both views, including mixed fields. It is not a per-cell histogram;
detailed `get-format` permits at most 10,000 cells.