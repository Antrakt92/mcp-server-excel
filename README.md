<h1 align="center">ExcelMCP — Antrakt Fork</h1>

<p align="center">Control desktop Microsoft Excel through MCP or the excelcli command line.</p>

<p align="center">
  <a href="https://github.com/Antrakt92/mcp-server-excel/releases/latest"><strong>Download</strong></a>
  · <a href="#connect-an-mcp-client"><strong>MCP setup</strong></a>
  · <a href="#use-the-cli"><strong>CLI</strong></a>
  · <a href="docs/fork/AGENT-WORKFLOWS.md">Agent workflows</a>
</p>

<p align="center">
  <a href="https://github.com/Antrakt92/mcp-server-excel/releases/latest"><img src="https://img.shields.io/github/v/release/Antrakt92/mcp-server-excel" alt="Latest release"></a>
  <img src="https://img.shields.io/badge/platform-Windows%20x64-0078D4" alt="Windows x64">
  <a href="LICENSE"><img src="https://img.shields.io/github/license/Antrakt92/mcp-server-excel" alt="MIT license"></a>
</p>

This independent fork of [ExcelMcp by Stefan Broenner](https://github.com/sbroenne/mcp-server-excel) adds read-only audit sessions, bounded range reads and VBA edits with source checks. Workbook operations use Excel's native COM interface.

**Requires Windows x64, installed Microsoft Excel 2016 or later, and an interactive desktop.** Portable release packages include their .NET runtime; Excel itself is not included.

[Features](FEATURES.md) · [Changelog](docs/fork/CHANGELOG.md) · [Verification](docs/fork/VERIFICATION-3.0.1.md)

## Choose a package

| Package | Use |
| --- | --- |
| [ExcelMcp-Antrakt-MCP-3.0.1-win-x64.zip](https://github.com/Antrakt92/mcp-server-excel/releases/download/v3.0.1/ExcelMcp-Antrakt-MCP-3.0.1-win-x64.zip) | A local stdio MCP server for a compatible AI client. |
| [ExcelMcp-Antrakt-CLI-3.0.1-win-x64.zip](https://github.com/Antrakt92/mcp-server-excel/releases/download/v3.0.1/ExcelMcp-Antrakt-CLI-3.0.1-win-x64.zip) | The portable `excelcli` command line and its background service. |

Download **SHA256SUMS**, **RELEASE-MANIFEST.json** and **VERIFICATION-RESULTS.json** from the same release. The manifest identifies the source revision, upstream base, package hashes and packaging checks. The separate verification receipt records protocol, native Excel and CLI checks against the exact executable hashes. Match their source revision and executable hashes. GitHub also provides source ZIP and tar archives on the release page.

<details>
<summary>Download with PowerShell</summary>

```powershell
$release = 'https://github.com/Antrakt92/mcp-server-excel/releases/download/v3.0.1'
Invoke-WebRequest "$release/ExcelMcp-Antrakt-MCP-3.0.1-win-x64.zip" -OutFile 'ExcelMcp-Antrakt-MCP-3.0.1-win-x64.zip'
Invoke-WebRequest "$release/ExcelMcp-Antrakt-CLI-3.0.1-win-x64.zip" -OutFile 'ExcelMcp-Antrakt-CLI-3.0.1-win-x64.zip'
Invoke-WebRequest "$release/SHA256SUMS" -OutFile 'SHA256SUMS'
Invoke-WebRequest "$release/RELEASE-MANIFEST.json" -OutFile 'RELEASE-MANIFEST.json'
Invoke-WebRequest "$release/VERIFICATION-RESULTS.json" -OutFile 'VERIFICATION-RESULTS.json'
Get-FileHash .\ExcelMcp-Antrakt-*-3.0.1-win-x64.zip -Algorithm SHA256
Get-FileHash -LiteralPath 'RELEASE-MANIFEST.json', 'VERIFICATION-RESULTS.json' -Algorithm SHA256
```

</details>

Compare each hash with `SHA256SUMS` before extracting. Keep releases in separate folders when upgrading. See [recovery and verification](docs/fork/RECOVERY.md) for the full procedure.

## Connect an MCP client

Extract the MCP ZIP, then configure your client to launch `Sbroenne.ExcelMcp.McpServer.exe`. For clients using an `mcpServers` configuration, the following is an example; replace the executable path with your extraction location:

```json
{
  "mcpServers": {
    "excel-antrakt": {
      "command": "C:/Tools/ExcelMCP/MCP/Sbroenne.ExcelMcp.McpServer.exe",
      "env": { "EXCELMCP_TELEMETRY_MODE": "off" }
    }
  }
}
```

Restart or reconnect that client connection after changing its executable. Updating files does not replace a server already running in memory. Client configuration formats vary; use the equivalent local stdio command setting in your client.

## Use the CLI

Extract the CLI ZIP and run:

```powershell
.\excelcli.exe --help
```

MCP and CLI provide the same underlying Excel capabilities through different interfaces. Their command names and parameter spelling differ; use current CLI help or the connected MCP tool descriptions.

## Migrate from fork 2.5.4

Fork 3.0.0 was based on upstream 2.6.0 and introduced the CLI changes below. Fork 3.0.1 includes upstream 2.6.1 maintenance and the later chart-label fix; no migration is needed from 3.0.0. Published 3.0.0 archives and verification receipts keep their original 2.6.0 source base.

CLI commands and batch entries now use a common JSON envelope. Read workbook results under `data`, failure details under `error`, and operation time under `durationMs`; check top-level `success` and the process exit code. Batch entries retain `index`. Update scripts that used the old result wrapper or command summary.

Replace CLI `--sheet` with `--sheet-name` and `--range` with `--range-address`. The old aliases are rejected. Optional result-field selection has also been removed; use current action help to update saved requests. MCP clients should rediscover the tool surface after reconnecting; replies continue to provide JSON text and structured content.

Normal file opens disable VBA macros. Running a macro does not enable them: the separate `file enable-macros` operation requires explicit approval and permission to reopen a saved owned workbook (CLI: `session enable-macros --approved --allow-reopen`). Audit sessions reject macro execution, workbook edits, calculation and saving. See [agent workflows](docs/fork/AGENT-WORKFLOWS.md) and [upgrade/recovery](docs/fork/RECOVERY.md) before continuing an existing automation.

## What this fork changes

- MCP shutdown observes stdin EOF, terminal stdout failure and owning-client exit without a competing stdin reader.
- Stdout remains a single JSON-RPC protocol stream.
- MCP request admission is bounded through the underlying operation and cancellation cleanup.
- CLI shutdown drains client connections within a bounded grace period; retained idle clients cannot hold the daemon open indefinitely.
- MCP response conversion shares its normalized JSON parse between result construction and error classification.
- Persistent audit sessions refuse agent edits, explicit calculation, macros and saving.
- Bounded range pages and formatting summaries avoid materializing oversized reads.
- Guarded whole-module VBA replacement detects stale source; typed invocation returns bounded scalar or rectangular results.

Excel remains the calculation and workbook engine. Read [the fork scope](docs/fork/OVERVIEW.md) for version boundaries and validation limits. Release-specific results belong in the release verification receipt (`VERIFICATION-RESULTS.json`); measurements from the older 2.4.1 patch are not performance evidence for this release.

### Workbook safety

Discover and attach to an already-open workbook when supported; do not reopen it just to create another automation session. Operate only on the intended workbook in your own Windows session.

Normal shutdown closes saved workbooks opened by the host, leaves unsaved workbooks open and visible without saving or discarding them, and leaves attached workbooks untouched. Cancellation stops the caller's wait; it does not authorize saving, closing, discarding or rolling back workbook changes. Inspect the current session and Excel window before retrying. Excel AutoSave may independently persist edits.

Do not terminate unrelated Excel or MCP processes during recovery. See [recovery](docs/fork/RECOVERY.md).

### Command coverage

The combined catalogue provides **63 MCP tools across 33 feature areas, with 436 operations**. See [FEATURES.md](FEATURES.md) and current tool descriptions for all 436 operations. These counts are checked against the compiled contracts.

## Source and contributions

The fork is maintained at [Antrakt92/mcp-server-excel](https://github.com/Antrakt92/mcp-server-excel). Report fork-specific lifecycle or packaging problems in [this fork's issue tracker](https://github.com/Antrakt92/mcp-server-excel/issues), with the release version and a minimal reproduction. Exclude workbook contents, credentials, private paths and unsanitized logs.

Contributions target `main`; start with [the contributor guide](CONTRIBUTING.md). Report suspected vulnerabilities through [the security policy](SECURITY.md), rather than public issues. Build and verification ownership is described in [the fork overview](docs/fork/OVERVIEW.md). The original upstream README is preserved in [README.upstream.md](README.upstream.md); its upstream installers and marketplace links install upstream distributions.

## Credits and license

ExcelMcp was created by **Stefan Broenner** and its upstream contributors. This fork retains the original [MIT license](LICENSE) and copyright notice. Fork maintenance is by **Antrakt92**. This is an independent project and does not claim endorsement by the upstream author, Microsoft or OpenAI. See [FORK-NOTICE.md](FORK-NOTICE.md).
