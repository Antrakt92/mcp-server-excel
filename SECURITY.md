# Security Policy

This policy covers the independently maintained
[Antrakt ExcelMCP fork](https://github.com/Antrakt92/mcp-server-excel).
Upstream ExcelMcp is maintained separately by
[Stefan Broenner and contributors](https://github.com/sbroenne/mcp-server-excel).

## Supported releases and reporting

Use the latest stable [fork release](https://github.com/Antrakt92/mcp-server-excel/releases).
Older packages remain available for rollback and reproducibility; this fork does
not maintain parallel security-fix branches for them. Fork portable MCP and CLI
packages are separate from upstream npm, NuGet, marketplace and registry
releases. Identify the repository, exact version and package used in a report.

Report suspected vulnerabilities privately through
[the fork's vulnerability reporting form](https://github.com/Antrakt92/mcp-server-excel/security/advisories/new).
Do not put confidential reports in public issues, pull requests, profile comments
or diagnostic attachments. GitHub does not provide a general private direct
message channel. If the private form is unavailable, do not substitute a public
issue containing exploit details or private data.

Include the affected version, entry point, minimal synthetic reproduction,
expected boundary, observed impact and any relevant sanitized diagnostics.
Exclude credentials, connection strings, customer workbooks and personal data.
Use the private advisory discussion to coordinate disclosure. There is no
promised response or remediation deadline.

## System and trust boundaries

ExcelMCP controls installed desktop Microsoft Excel through native COM in the
current Windows user's session. MCP stdio and the CLI are equal entry points;
controllers delegate Excel work to local worker processes. Excel, Office add-ins,
Windows permissions, the chosen MCP client and the AI assistant are part of the
operating environment.

The product runs with the current user's permissions. It is not a filesystem
sandbox or a boundary against other software already running as that user. Absolute-path normalization and physical workbook identity
tracking do not restrict access to an approved directory. There is no documented
global 1 GB workbook size limit. Creating new workbooks is restricted to `.xlsx`
and `.xlsm`; opening existing workbooks delegates format support to Excel.

Treat workbook contents, formulas, VBA and Power Query code, external connections
and tool arguments as untrusted inputs. Authorize operations for the selected
workbook. Returned tool results can include requested workbook data, source code
and paths; the MCP client and assistant determine how they handle those results.
Telemetry is a separate data flow described in [PRIVACY.md](PRIVACY.md).

## Controls and limitations

- Windows named-pipe ACLs restrict controller access to the current user's SID.
  Private workers also validate their recorded owner process and creation time.
  The CLI daemon intentionally accepts commands from other processes running as
  that same user. A pipe name alone is not an authorization control.
- Product IPC uses stdio and Windows named pipes. This does not make Excel or
  workbook operations network-isolated: network paths, remote workbooks, data
  connections, Office services and add-ins may involve external systems using
  the user's access. Do not assume that UNC paths or network drives are blocked.
- Normal workbook opens disable VBA macros and external-link updates. Enabling
  macros requires the separate approved saved-workbook reopen workflow; opening
  event code can then execute. VBA project access is a separate Office setting.
  Excel and organization policies still govern protected content and editing.
- Persistent audit sessions (`read_only:true` in MCP, `--read-only` in the CLI)
  reject agent mutations, explicit calculation, macro execution, security changes
  and saving before dispatch. This API restriction does not prevent interactive
  editing, every Excel opening/background action, or all network activity.
- Requested visibility, authentication and recovery can show Excel windows.
  `Visible=false` and `DisplayAlerts=false` are not permanent security guarantees.
- Cancellation and timeout do not undo workbook edits or establish that Excel has
  stopped. Session ownership, pending work and partial failures must remain
  observable. Inspect current state before retrying a mutation or closing it.
- Guarded VBA edits check the expected source fingerprint and read back stored
  code. A partial failure reports observed state; it does not automatically roll
  back or discard workbook changes. VBA execution remains an authorized code
  execution capability, not a sandbox for untrusted macros.
- Bounded range and VBA result limits constrain specific responses. They do not
  guarantee a total memory, execution-time or resource bound for every inherited
  Excel operation. Diagnostic output is not a complete durable security audit log.

See [agent workflows](docs/fork/AGENT-WORKFLOWS.md),
[recovery](docs/fork/RECOVERY.md) and
[workbook opening notices](docs/reference/behavioral-rules.md#opening-notices-and-blocked-content)
for the implemented workflows and their limits.

## Validation and repository safeguards

The repository contains compiler/analyzer checks, dependency review, Dependabot
configuration and CodeQL workflows. Configuration files describe intended checks;
they do not prove that a particular run succeeded, that repository protection is
active, or that there are no vulnerabilities. Fork publication uses the local
portable build and verification workflow; inherited hosted Actions publication
routes are disabled.

Required checks and branch protection are repository settings that must be
verified separately. Passing compilation or static analysis does not establish
native Excel correctness, complete COM cleanup, workbook integrity or runtime
acceptance. Release receipts bind actual package checks to source and executable
hashes; see [verification](docs/fork/VERIFICATION-3.0.0.md).

Report reproducible defects in these boundaries with their actual caller,
workbook state and impact. Do not include private workbook data in public
reproductions.
