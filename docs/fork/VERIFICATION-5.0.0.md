# Verify ExcelMCP 5.0.0

Build the portable Windows x64 release from clean, merged source using the pinned SDK and Microsoft ADO importer.

    pwsh -NoProfile -File scripts/Build-ForkRelease.ps1 -Version 5.0.0

Verify the source commit and tree, asset and executable SHA-256 hashes, x64 PE headers, and file versions. Extract the verified ZIPs before running package checks.

- `scripts/tests/fork_protocol_smoke.py` with `--idle-seconds 15` checks MCP initialization, command and error parity, idle resources, EOF, broken stdout, and parent exit.
- `scripts/tests/fork_packaged_excel_smoke.py` checks the packaged CLI and MCP through native Excel, including workbook writes, persistence, and lifecycle cleanup.
- `scripts/tests/fork_vba_authentication_smoke.py` checks explicit VBA authentication, incorrect-password cleanup, stdin-only CLI handling, and saving and reopening a synthetic macro workbook.

Record the actual results in `VERIFICATION-RESULTS.json`, bound to the exact package executables. Keep private workbooks, workbook contents, and local paths out of the public receipt.

Publish the two ZIPs, `RELEASE-MANIFEST.json`, `VERIFICATION-RESULTS.json`, and `SHA256SUMS`. Download the published files and verify that all five match the local release outputs.

After switching the installation, verify the configured CLI and a fresh MCP process. Existing chats may need to reconnect. Retire an older host after checking that its sessions and pending operations are clear.
