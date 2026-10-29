# Verify, upgrade and recover

Use packages from [this fork's releases](https://github.com/Antrakt92/mcp-server-excel/releases). Download the MCP or CLI ZIP together with `SHA256SUMS`, `RELEASE-MANIFEST.json` and `VERIFICATION-RESULTS.json` from the **same release**.

## Verify a download

Run `Get-FileHash <downloaded-file> -Algorithm SHA256` for the ZIP, manifest and verification receipt, and compare each full value with its matching filename in `SHA256SUMS`. Check the manifest's version, upstream base and source revision before extracting. Match the verification receipt's source revision and executable hashes to the manifest and the extracted executables; see [release verification](OVERVIEW.md#distribution). Hashes verify that your download matches the publisher's release; they do not independently establish publisher trust or workbook correctness.

Extract into a new version-specific directory. Do not overwrite a running executable. Keep the previous verified package until the new version has passed the checks relevant to your work.

## Switch versions

1. Inspect active work and the intended Excel windows. Decide whether to save changes explicitly; shutdown is not permission to save or discard them.
2. Stop the relevant MCP connection using its client, or use the CLI's documented service controls. Do not kill all Excel or MCP processes.
3. Point the client configuration or command to the new extracted executable. Restart or reconnect that connection.
4. Check the new executable's version/help and MCP initialization through the intended client. A changed path or hash alone does not prove the client activated the new process.
5. Discover/list the intended workbook and inspect its state before continuing. Avoid opening another copy of a workbook already available to attach.

To roll back, repeat these steps with the previous verified release. Preserve workbook state first. Switching binaries cannot undo workbook edits.

## Migrate automation from fork 2.5.4

Fork 3.0.0 uses upstream 2.6.0 and requires CLI consumer changes:

- Read returned values from `data` and failure details from `error`; the common envelope also includes `command`, `success` and `durationMs`. Batch entries retain their `index`.
- Replace `--sheet` and `--range` with `--sheet-name` and `--range-address`. Remove obsolete result-field selection from saved requests; current action help is authoritative.
- Check formula preflight `validationScope` and `unverifiedChecks`; `isValid` does not prove Excel compilation, UDF availability or safe overwriting.
- Rediscover MCP tools after reconnecting. Tool results retain JSON text and structured content; do not infer success from transport delivery alone.
- Normal opens disable VBA macros. The explicit enable-macros approval and saved-workbook reopening gate applies to both `vba run` and `vba invoke`; audit sessions cannot enable or execute macros.

Finish or explicitly hand over current workbook work with the old version before switching. Test migrated automation on synthetic data before using operational workbooks. See [agent workflows](AGENT-WORKFLOWS.md) for current inspection and VBA limits. Older verification records describe their own release only.

## After cancellation, timeout or a lost connection

Cancellation and timeout stop waiting; Excel's work may continue. Check the Excel window and discover/list sessions before retrying. Do not assume the operation failed, undo partial changes automatically, or close a workbook as cleanup. If the host exits, unsaved workbooks may remain open and visible and attached workbooks stay untouched, following upstream 2.6.0 ownership rules. Pending work can remain in a retained worker; inspect its reported state using the matching build before retrying. Changing binaries does not finish, reverse or authorize cleanup of that work. AutoSave may already have saved changes independently.

If recovery requires a process-level action, establish the exact owned process identity and current workbook state first. A process name alone does not establish ownership. Unrelated Excel instances must remain untouched.

## Rebuild or report a problem

The separately labeled [historical 2.4.1 patch](LEGACY.md) preserves the earlier
development line. It is archival source and must not be applied to current releases.

Use the tagged source archive or Git checkout matching the release manifest. Follow [the build and verification boundary](OVERVIEW.md#build-and-verify) and `scripts/Build-ForkRelease.ps1`. Preserve newly generated checksums and manifest; a rebuild can differ from the published binary and needs its own verification.

For a fork issue, include the version, package checksum, Windows/Excel versions, entry point and minimal reproduction. Describe observed state and which checks ran. Remove workbook contents, credentials, user-specific paths and private logs before publishing.
