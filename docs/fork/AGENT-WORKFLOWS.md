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
