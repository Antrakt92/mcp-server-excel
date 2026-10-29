# ExcelMCP — Antrakt Fork

Automate desktop Microsoft Excel through MCP or the `excelcli` command line, using Excel's native COM interface.

This is an independently maintained fork of [ExcelMcp by Stefan Broenner](https://github.com/sbroenne/mcp-server-excel). The **2.5.4 fork release line is based on upstream 2.5.3**. The development branch integrates upstream 2.6.0 while retaining fork transport shutdown, bounded request admission and connection drain. The download links below remain the last published 2.5.4 release until the new candidate is verified.

**Requires Windows x64, installed Microsoft Excel 2016 or later, and an interactive desktop.** Portable release packages include their .NET runtime; Excel itself is not included.

[Download releases](https://github.com/Antrakt92/mcp-server-excel/releases/latest) · [Fork overview](docs/fork/OVERVIEW.md) · [Verification](docs/fork/VERIFICATION-2.5.4.md) · [Recovery](docs/fork/RECOVERY.md) · [Upstream documentation](README.upstream.md) · [MIT license](LICENSE)

## Choose a package

| Package | Use |
| --- | --- |
| [ExcelMcp-Antrakt-MCP-2.5.4-win-x64.zip](https://github.com/Antrakt92/mcp-server-excel/releases/latest/download/ExcelMcp-Antrakt-MCP-2.5.4-win-x64.zip) | A local stdio MCP server for a compatible AI client. |
| [ExcelMcp-Antrakt-CLI-2.5.4-win-x64.zip](https://github.com/Antrakt92/mcp-server-excel/releases/latest/download/ExcelMcp-Antrakt-CLI-2.5.4-win-x64.zip) | The portable `excelcli` command line and its background service. |

Download **SHA256SUMS** and **RELEASE-MANIFEST.json** from the same release. The manifest identifies the source revision, upstream base, package hashes and recorded verification. GitHub also provides source ZIP and tar archives on the release page.

PowerShell download example:

```powershell
$release = 'https://github.com/Antrakt92/mcp-server-excel/releases/latest/download'
Invoke-WebRequest "$release/ExcelMcp-Antrakt-MCP-2.5.4-win-x64.zip" -OutFile 'ExcelMcp-Antrakt-MCP-2.5.4-win-x64.zip'
Invoke-WebRequest "$release/ExcelMcp-Antrakt-CLI-2.5.4-win-x64.zip" -OutFile 'ExcelMcp-Antrakt-CLI-2.5.4-win-x64.zip'
Invoke-WebRequest "$release/SHA256SUMS" -OutFile 'SHA256SUMS'
Invoke-WebRequest "$release/RELEASE-MANIFEST.json" -OutFile 'RELEASE-MANIFEST.json'
Get-FileHash .\ExcelMcp-Antrakt-*-2.5.4-win-x64.zip -Algorithm SHA256
```

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

## What this fork changes

- MCP shutdown observes stdin EOF, terminal stdout failure and owning-client exit without a competing stdin reader.
- Stdout remains a single JSON-RPC protocol stream.
- MCP request admission is bounded through the underlying operation and cancellation cleanup.
- CLI shutdown drains client connections within a bounded grace period; retained idle clients cannot hold the daemon open indefinitely.
- Standard MCP responses are parsed once while preserving their text, structured content and error classification.

These are lifecycle changes, not a replacement calculation engine. Read [the fork scope](docs/fork/OVERVIEW.md) for version boundaries and validation limits. Release-specific results belong in the release manifest; measurements from the older 2.4.1 patch are not performance evidence for this release.

### Workbook safety

Discover and attach to an already-open workbook when supported; do not reopen it just to create another automation session. Operate only on the intended workbook in your own Windows session.

Upstream 2.5.3 behavior is retained: normal shutdown closes saved workbooks opened by the host, leaves unsaved workbooks open and visible without saving or discarding them, and leaves attached workbooks untouched. Cancellation stops the caller's wait; it does not authorize saving, closing, discarding or rolling back workbook changes. Inspect the current session and Excel window before retrying. Excel AutoSave may independently persist edits.

Do not terminate unrelated Excel or MCP processes during recovery. See [recovery](docs/fork/RECOVERY.md).

### Command coverage

The inherited catalogue provides **63 MCP tools across 33 feature areas, with 434 operations**. See [FEATURES.md](FEATURES.md) and current tool descriptions for all 434 operations. These upstream counts are checked against the compiled contracts.

## Source and contributions

The fork is maintained at [Antrakt92/mcp-server-excel](https://github.com/Antrakt92/mcp-server-excel). Report fork-specific lifecycle or packaging problems in [this fork's issue tracker](https://github.com/Antrakt92/mcp-server-excel/issues), with the release version and a minimal reproduction. Exclude workbook contents, credentials, private paths and unsanitized logs.

Build and verification ownership is described in [the fork overview](docs/fork/OVERVIEW.md). The original upstream README is preserved in [README.upstream.md](README.upstream.md); its upstream installers and marketplace links install upstream distributions.

## Credits and license

ExcelMcp was created by **Stefan Broenner** and its upstream contributors. This fork retains the original [MIT license](LICENSE) and copyright notice. Fork maintenance is by **Antrakt92**. This is an independent project and does not claim endorsement by the upstream author, Microsoft or OpenAI. See [FORK-NOTICE.md](FORK-NOTICE.md).
