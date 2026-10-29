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
changes and saving before dispatching them to Excel. Close with `save:false`
when inspection is finished. Editing requires a separately authorized normal
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
