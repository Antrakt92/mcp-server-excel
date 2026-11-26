<h1 align="center">Excel MCP Server for Windows</h1>

<p align="center">Inspect, edit, and verify workbooks in desktop Microsoft Excel from an MCP client or the excelcli command line.</p>

<p align="center">
  <a href="https://github.com/Antrakt92/mcp-server-excel/releases/latest"><strong>Download</strong></a>
  · <a href="#connect-an-mcp-client"><strong>MCP setup</strong></a>
  · <a href="#try-a-read-only-audit"><strong>Try a workbook audit</strong></a>
  · <a href="docs/fork/AGENT-WORKFLOWS.md">Agent workflows</a>
</p>

<p align="center">
  <a href="https://github.com/Antrakt92/mcp-server-excel/releases/latest"><img src="https://img.shields.io/github/v/release/Antrakt92/mcp-server-excel" alt="Latest release"></a>
  <img src="https://img.shields.io/badge/platform-Windows%20x64-0078D4" alt="Windows x64">
  <a href="LICENSE"><img src="https://img.shields.io/github/license/Antrakt92/mcp-server-excel" alt="MIT license"></a>
</p>

ExcelMCP — Antrakt Fork is built on [Stefan Broenner’s ExcelMcp](https://github.com/sbroenne/mcp-server-excel). It uses Excel’s native COM interface for workbook operations, including formulas, charts, PivotTables, Power Query, and VBA.

**Requires Windows x64, installed Microsoft Excel 2016 or later, and an interactive desktop.** Portable release packages include their .NET runtime; Excel itself is not included.

[Command catalogue](FEATURES.md) · [Changelog](docs/fork/CHANGELOG.md) · [Release verification](https://github.com/Antrakt92/mcp-server-excel/releases/download/v5.4.0/VERIFICATION-RESULTS.json)

## What this fork adds

- **Audit an existing workbook.** Open a persistent read-only session and page through large ranges. The server refuses Excel edits, explicit calculation, macro execution, and saves in audit mode.
- **Update a VBA module safely.** Pass the source hash from your last read when replacing a module. If the code changed in the meantime, the replacement is refused. The server reads the stored source back after a successful edit.
- **Check what a save changed.** Compare two saved XLSX or XLSM files by package part. Optional structural details show sheet identities, relationships, local name scopes, and parent-style references.

See [agent workflows](docs/fork/AGENT-WORKFLOWS.md) for examples and [the fork overview](docs/fork/OVERVIEW.md) for lifecycle changes and scope.

## Choose a package

Use the packages below to install this fork. The npm, NuGet, and marketplace packages linked from the [upstream README](README.upstream.md) install the upstream project.

| Package | Use |
| --- | --- |
| [Windows MCP bundle](https://github.com/Antrakt92/mcp-server-excel/releases/download/v5.4.0/ExcelMcp-Antrakt-MCPB-5.4.0-win-x64.mcpb) | Install as a desktop extension in a client that supports MCPB. Contains the Windows x64 server and disables telemetry in its launch configuration. |
| [Windows MCP ZIP](https://github.com/Antrakt92/mcp-server-excel/releases/download/v5.4.0/ExcelMcp-Antrakt-MCP-5.4.0-win-x64.zip) | Configure a compatible MCP client to launch the local stdio server. |
| [Windows CLI ZIP](https://github.com/Antrakt92/mcp-server-excel/releases/download/v5.4.0/ExcelMcp-Antrakt-CLI-5.4.0-win-x64.zip) | The portable `excelcli` command line and its background service. |

Download **SHA256SUMS**, **RELEASE-MANIFEST.json** and **VERIFICATION-RESULTS.json** from the same release. The manifest identifies the source revision, upstream base, package hashes and packaging checks. The separate verification receipt records protocol, native Excel and CLI checks against the exact executable hashes. Match their source revision and executable hashes. GitHub also provides source ZIP and tar archives on the release page.

<details>
<summary>Download with PowerShell</summary>

```powershell
$release = 'https://github.com/Antrakt92/mcp-server-excel/releases/download/v5.4.0'
Invoke-WebRequest "$release/ExcelMcp-Antrakt-MCP-5.4.0-win-x64.zip" -OutFile 'ExcelMcp-Antrakt-MCP-5.4.0-win-x64.zip'
Invoke-WebRequest "$release/ExcelMcp-Antrakt-CLI-5.4.0-win-x64.zip" -OutFile 'ExcelMcp-Antrakt-CLI-5.4.0-win-x64.zip'
Invoke-WebRequest "$release/SHA256SUMS" -OutFile 'SHA256SUMS'
Invoke-WebRequest "$release/RELEASE-MANIFEST.json" -OutFile 'RELEASE-MANIFEST.json'
Invoke-WebRequest "$release/VERIFICATION-RESULTS.json" -OutFile 'VERIFICATION-RESULTS.json'
Get-FileHash .\ExcelMcp-Antrakt-*-5.4.0-win-x64.zip -Algorithm SHA256
Get-FileHash -LiteralPath 'RELEASE-MANIFEST.json', 'VERIFICATION-RESULTS.json' -Algorithm SHA256
```

</details>

Compare each hash with `SHA256SUMS` before extracting. Keep releases in separate folders when upgrading. See [recovery and verification](docs/fork/RECOVERY.md) for the full procedure.

## Connect an MCP client

For MCPB, verify the downloaded archive against `SHA256SUMS` and use your desktop client's extension installer. In Claude Desktop, open **Settings → Desktop app → Manage Extensions → Advanced settings → Install Extension** and select the bundle. If you already configured ExcelMCP manually, use one connection so the tool list stays unambiguous.

For a manual stdio connection, extract the MCP ZIP and configure your client to launch `Sbroenne.ExcelMcp.McpServer.exe`. Replace the executable path with your extraction location in this example for clients that use `mcpServers`.

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

If a client reports `Transport closed`, see [how to reconnect and recover an
open workbook](docs/fork/RECOVERY.md#when-an-mcp-connection-closes).

## Use the CLI

Extract the CLI ZIP and run:

```powershell
.\excelcli.exe --help
```

MCP and CLI provide the same underlying Excel capabilities through different interfaces. Their command names and parameter spelling differ; use current CLI help or the connected MCP tool descriptions.

## Try a read-only audit

In PowerShell, run this from the extracted CLI folder. Point `$path` to a closed test copy of an XLSX or XLSM workbook.

```powershell
$path = (Resolve-Path 'C:\Work\Example.xlsx' -ErrorAction Stop).Path
$before = (Get-FileHash -LiteralPath $path).Hash
$opened = .\excelcli.exe -q session open $path --read-only | ConvertFrom-Json
if (-not $opened.success) { throw $opened.error.message }

$session = $opened.data.sessionId
$listed = $false
try {
    if (-not $opened.data.auditOnly) { throw 'Expected an audit session.' }
    .\excelcli.exe -q sheet list --session $session
    if ($LASTEXITCODE -ne 0) { throw 'Sheet listing failed.' }
    $listed = $true
}
finally {
    .\excelcli.exe -q session close --session $session
    $closeExit = $LASTEXITCODE
    if ($closeExit -ne 0 -and -not $listed) {
        Write-Warning "Closing failed. Check session $session before retrying."
    }
}
if ($closeExit -ne 0) { throw 'Closing the audit session failed.' }
if ((Get-FileHash -LiteralPath $path).Hash -ne $before) {
    throw 'The saved workbook changed.'
}
'Workbook SHA-256 unchanged.'
```

The output includes the sheet list and close result. In an MCP client, use `file` action `open` with `read_only: true`. Pass the returned session ID as `workbook_session_id` to `worksheet_read` action `list`, then use `file` action `close` with the same ID and `save: false`. See [persistent inspection](docs/fork/AGENT-WORKFLOWS.md#persistent-inspection) for the session limits. To try an edit, use the [sample workbook script](examples/README.md).

## Upgrading

The portable release **5.4.0** is based on upstream **2.6.12**. See the [5.4.0 notes](docs/fork/CHANGELOG.md#540--2026-11-17) for the release contents.

## Migrate password calls from fork 3.6.0

Replace direct `vba protect-project --password ...` calls and password-bearing
batch files with serialized batch JSON on stdin. This applies to `protect-project`,
`unlock-project` and `verify-password`. A file batch containing any password action
runs no commands; ordinary file batches remain supported. MCP tool inputs remain
available.

In PowerShell 7, with `$sessionId` already set to the target session:

```powershell
$password = Read-Host 'Known VBA project password' -MaskInput
@{
    command = 'vba.unlock-project'
    sessionId = $sessionId
    args = @{ password = $password }
} | ConvertTo-Json -Depth 4 -Compress | .\excelcli.exe batch --input -
Remove-Variable password
```

Use command `vba.protect-project` with the new value in `args.password` when
configuring protection. Omitting `--input` also reads stdin. Password text is
preserved exactly, including nonempty all-space values; missing, null and empty
passwords are invalid. Keep credentials out of arguments, files and logs.

Unlocking an already accessible project does not verify the supplied password.
Verification requires a real locked-project challenge; accepted input also opens
current-session access. Neither action removes protection settings, saves or
reopens the workbook. Inspect unknown/error state before retrying: cancellation
is not rollback. See [agent workflows](docs/fork/AGENT-WORKFLOWS.md#vba-project-passwords).

## Migrate from fork 2.5.4

Fork 3.0.0 was based on upstream 2.6.0 and introduced the CLI changes below. Fork 3.0.1 includes upstream 2.6.1 maintenance and the later chart-label fix; no migration is needed from 3.0.0. Published 3.0.0 archives and verification receipts keep their original 2.6.0 source base.

CLI commands and batch entries now use a common JSON envelope. Read workbook results under `data`, failure details under `error`, and operation time under `durationMs`; check top-level `success` and the process exit code. Batch entries retain `index`. Update scripts that used the old result wrapper or command summary.

Replace CLI `--sheet` with `--sheet-name` and `--range` with `--range-address`. The old aliases are rejected. Optional result-field selection has also been removed; use current action help to update saved requests. MCP clients should rediscover the tool surface after reconnecting; replies continue to provide JSON text and structured content.

Normal file opens disable VBA macros. Running a macro does not enable them: the separate `file enable-macros` operation requires explicit approval and permission to reopen a saved owned workbook (CLI: `session enable-macros --approved --allow-reopen`). Audit sessions reject macro execution, workbook edits, calculation and saving. See [agent workflows](docs/fork/AGENT-WORKFLOWS.md) and [upgrade/recovery](docs/fork/RECOVERY.md) before continuing an existing automation.

## Workbook safety

Discover and attach to an already-open workbook when supported; do not reopen it just to create another automation session. Operate only on the intended workbook in your own Windows session.

Normal shutdown closes saved workbooks opened by the host, leaves unsaved workbooks open and visible without saving or discarding them, and leaves attached workbooks untouched. Cancellation stops the caller's wait; it does not authorize saving, closing, discarding or rolling back workbook changes. Inspect the current session and Excel window before retrying. Excel AutoSave may independently persist edits.

Do not terminate unrelated Excel or MCP processes during recovery. See [recovery](docs/fork/RECOVERY.md).

## Command coverage

The combined catalogue provides **64 MCP tools across 33 feature areas, with 450 operations**. See [FEATURES.md](FEATURES.md) and current tool descriptions for all 450 operations. These counts are checked against the compiled contracts.

## Source and contributions

The fork is maintained at [Antrakt92/mcp-server-excel](https://github.com/Antrakt92/mcp-server-excel). Report fork-specific lifecycle or packaging problems in [this fork's issue tracker](https://github.com/Antrakt92/mcp-server-excel/issues), with the release version and a minimal reproduction. Exclude workbook contents, credentials, private paths and unsanitized logs.

Contributions target `main`; start with [the contributor guide](CONTRIBUTING.md). Report suspected vulnerabilities through [the security policy](SECURITY.md), rather than public issues. Build and verification ownership is described in [the fork overview](docs/fork/OVERVIEW.md). The original upstream README is preserved in [README.upstream.md](README.upstream.md); its upstream installers and marketplace links install upstream distributions.

## Credits and license

ExcelMcp was created by **Stefan Broenner** and its upstream contributors. This fork retains the original [MIT license](LICENSE) and copyright notice. Fork maintenance is by **Antrakt92**. This is an independent project and does not claim endorsement by the upstream author, Microsoft or OpenAI. See [FORK-NOTICE.md](FORK-NOTICE.md).
