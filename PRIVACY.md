# Privacy Policy

**Last Updated:** October 29, 2026

## Overview

This policy describes the [Antrakt ExcelMCP fork](https://github.com/Antrakt92/mcp-server-excel),
which lets agents use installed desktop Excel through MCP or the CLI. See
[FORK-NOTICE.md](FORK-NOTICE.md) for upstream attribution and distribution scope.

The fork's portable release builder does not embed an Application Insights
connection string, so those packages default to telemetry **off**. You can also
set `EXCELMCP_TELEMETRY_MODE=off` explicitly. The optional inherited telemetry
implementation is described below; packages from other distribution channels
must be checked against their own build configuration and policy.

## Data Collection Summary

ExcelMcp's MCP Server and CLI support **limited, pseudonymous telemetry** when
reporting is enabled. The existing CLI background daemon sends operation metrics;
short-lived `excelcli` command processes do not send telemetry.
The statements below describe telemetry collection, not the workbook data
that a tool returns to your chosen AI assistant.

### Telemetry fields when reporting is enabled

- **Tool usage statistics** - Which tools and actions are used (e.g., "range/get-values").
  Only known ExcelMcp commands are reported; anything else is reported as "other/other"
- **Entry point** - Whether the operation was requested through the CLI or MCP Server
- **Performance metrics** - How long operations take (duration in milliseconds)
- **Invocation outcome** - Whether an operation succeeded, returned an expected
  negative diagnostic result, or failed
- **Failure class** - A fixed privacy-safe label for input or state, an external
  dependency, timeout or cancellation, Excel runtime, an internal product fault,
  or an unclassified failure
- **Session information** - A random telemetry session ID generated each time the CLI daemon or MCP Server starts
- **Pseudonymous user ID** - A random ID stored at
  `%LOCALAPPDATA%\ExcelMcp\telemetry-user-id`. The MCP Server and CLI share it
  for the same Windows user profile, and it persists across restarts until the
  file is deleted. Background reporting setup creates or reads it only when
  telemetry is enabled; disabled reporting never accesses the file.
  If the file cannot be read or written, that process uses a temporary random
  ID instead.
- **One-time migration link** - If telemetry is active when the ID file is
  first created, one `UserIdMigration` event is sent to the active destination
  with the new ID and the previous 16-character, unsalted hash of the machine
  name, user name, and platform. The raw names are not sent. The hash is
  included only to connect older usage events to the new ID; weekly reports do
  not publish it.
- **Application version** - Which version of ExcelMcp is running
- **Pseudonymous instance label** - `instance-` followed by the first eight
  characters of the random user ID. CLI and MCP events and requests
  use this label instead of the computer name.
- **Unhandled exceptions (MCP Server)** - Error type names, approved source, and
  project-owned failure site only (never exception messages or stack traces)

### Data excluded from telemetry

- **File contents** - Workbook content is never included in telemetry
- **File names or paths** - File paths are not included in telemetry
- **Personal information** - No raw names, emails, or account information
- **Spreadsheet data** - Cell values, formulas, and query results are never included in telemetry
- **Prompts or messages** - Requests and conversations are never transmitted
- **Host or client details** - Computer names, working directories, and MCP client names are not collected
- **User accounts** - No registration or sign-in required
- **Error details** - Error messages, response content, and stack traces are
  never included. Exception type names appear only in the unhandled-exception
  reports listed above, not in invocation outcome telemetry
- **Workbook session IDs or tool arguments** - Telemetry never includes the
  Excel workbook `workbook_session_id` (MCP) / `sessionId` (CLI) value, workbook path, request
  arguments, or user content. The random process telemetry session ID listed
  above is separate from workbook identity.

### Purpose of Telemetry

Pseudonymous telemetry is intended to:

- Understand which features are most used
- Identify and fix performance issues
- Prioritize development of new features
- Detect and fix bugs

Invocation telemetry reads only the structured category produced by ExcelMcp; it
does not inspect response or error text. Unknown or missing categories remain
**unclassified**.

### Telemetry Infrastructure

A build with an embedded connection string defaults to sending telemetry to
**Azure Application Insights**, a Microsoft service. The fork portable builder
leaves that connection string empty. Azure reporting uses HTTPS; storage and
retention depend on the configured receiving service. Its operator is responsible
for the destination and its data-handling settings.

The repository's upstream Azure infrastructure configuration defines ingestion
filters that discard framework trace logs and raw exception records, and retain
only explicitly sanitized failure events. These are deployment controls;
building this fork does not deploy or verify an Azure ingestion configuration.

Set `EXCELMCP_TELEMETRY_MODE=off` to disable telemetry collection and delivery in
either the CLI or MCP Server. Contributor builds without an embedded connection
string default to off.

Optional `EXCELMCP_TELEMETRY_MODE=local` sends only the approved operation
telemetry above to a loopback OpenTelemetry receiver on your own machine. It
overrides the embedded Azure destination and does not create an Azure telemetry
client. It does not enable framework logs, raw exception details, or additional
workbook data collection. Access to the local dashboard and any data it retains
is your responsibility. See [local telemetry setup](docs/DEVELOPMENT.md#local-development-with-telemetry).

Disabled reporting does not create the ID or a migration link. In local mode,
the link goes only to the local receiver and
is not deferred for Azure. In either case, older Azure events may remain under
the previous ID in weekly reports. Delete
`%LOCALAPPDATA%\ExcelMcp\telemetry-user-id` and restart the MCP Server to reset
the stored ID; the next active telemetry use creates a new one.

Reporting never waits for setup, saved-ID storage, export, or shutdown. Records
before setup is ready, overflowed batches, and unsent reports at process exit
are dropped. No shutdown flush or offline telemetry storage is used. Sampling,
live metrics, SDK self-statistics, performance counters, and automatic standard
metrics are disabled. CLI reporting starts in the existing daemon, not in each
command process. Help/version, failures before a request reaches the daemon,
and checks that find the daemon stopped are not reported. Internal readiness
pings are excluded. CLI timing and outcomes describe service execution, not
client-side connection, output-file writing, or argument parsing.

## How It Works

ExcelMcp drives the Excel application on your local machine:

1. **Local Processing** - All Excel operations are performed locally via Microsoft's COM API
2. **Workbook locations** - Excel can read or write selected local, network or
   cloud-hosted workbooks using the current user's file permissions and Office
   authentication. Opening or refreshing a workbook may contact its data sources
   or Office services; local COM automation is not network isolation.
3. **Tool Results** - Data requested by your AI assistant is returned through
   your MCP client or CLI process. Results can include requested cell data,
   formulas, VBA/Power Query code and paths; your client and assistant's privacy
   policies govern how they handle those results.
4. **Optional remote code services** - M/DAX formatting and Python in Excel use
   the external services described below when you request those features.
5. **Telemetry** - Release builds of the MCP Server and CLI can send the
   pseudonymous usage metrics listed above to Azure Application Insights

## Data Flow

When you use ExcelMcp with an AI assistant (like Claude):

1. You send a request to the AI assistant
2. The AI assistant calls ExcelMcp tools on your local machine
3. ExcelMcp performs the requested Excel operations locally
4. Requested results are returned to the AI assistant through your MCP client
5. The MCP Server or CLI can send pseudonymous usage telemetry to Azure Application Insights

Workbook access and data connections can involve network storage, cloud sync,
Office authentication and external data sources. Their permissions, retention
and privacy terms are controlled by those services and the workbook configuration.

Some explicitly requested code operations have additional data flows:

- Setting `formatMCode=true` sends the supplied M code to
  [powerqueryformatter.com](https://powerqueryformatter.com/).
- Setting `formatDax=true` sends the supplied DAX formula to
  [DAX Formatter](https://www.daxformatter.com/).
- Python in Excel sends Python code and referenced worksheet data to Microsoft's
  cloud execution environment. This feature requires Microsoft 365, internet
  access, and Python in Excel to be enabled.

Remote formatting is disabled by default and requires explicit consent.
ExcelMcp preserves M and DAX code locally unless you opt in.

**Note:** The AI assistant you use (for example, Claude or GitHub Copilot) has
its own privacy policy governing conversations, tool calls, and returned
workbook data.

## Local Power Query trace files

Power Query tracing follows the current Windows user's setting, which you can
read or change through the workbook-free Settings tools. Excel writes trace files to
the current user's local application data folder. These files are created and
retained by Excel, not sent by ExcelMcp telemetry. They can contain data-source
names and paths, native query text, and exception details.

Trace-file lists and summaries are returned only to the requesting MCP client or
CLI. Query text and exception messages are hidden from summaries by default;
requesting `include_query_text: true` includes those details in the tool result,
which is then subject to the AI assistant's privacy policy. Trace files remain
in the local folder until you remove them outside ExcelMcp. Tracing changes
require restarting Excel to take effect. The setting remains as chosen until
you change it; closing workbooks, sessions, or ExcelMCP does not restore it.
Other Excel processes started while it is on may also write trace files.

## Third-Party Services

- **Azure Application Insights** - A build configured for Azure reporting can
  send pseudonymous telemetry to this service. Fork portable packages default to
  off. See [Microsoft's Privacy Statement](https://privacy.microsoft.com/privacystatement).
- **Microsoft Excel** - ExcelMcp requires Microsoft Excel installed on your machine. Excel is subject to Microsoft's privacy policy.
- **Microsoft Python in Excel** - Python code and referenced worksheet data run
  in Microsoft's cloud when you use the Python in Excel operations.
- **Power Query Formatter** - Receives M code only when you set
  `formatMCode=true`.
- **DAX Formatter** - Receives DAX formulas only when you set
  `formatDax=true`.
- **AI Assistants** - When used with AI assistants like Claude, those services have their own privacy policies.

## Open Source

ExcelMcp is open source software. You can review the complete source code at:
https://github.com/Antrakt92/mcp-server-excel

## Security

- ExcelMcp runs with the same permissions as your user account
- It can only access files and Excel instances that your user account can access
- No elevated privileges are required or requested

## Changes to This Policy

If we make changes to this privacy policy, we will update the "Last Updated" date above and publish the updated policy in our GitHub repository.

## Contact

For questions about this privacy policy or the ExcelMcp project:

- **GitHub Issues:** https://github.com/Antrakt92/mcp-server-excel/issues
- **Repository:** https://github.com/Antrakt92/mcp-server-excel

Use [private security reporting](SECURITY.md#supported-releases-and-reporting) for
suspected vulnerabilities. Do not include private workbook data or credentials
in public issues.

---

**Summary:** ExcelMcp controls local Excel, which can access selected local or
remote workbook locations and data sources. Requested tool results can be
returned to your chosen AI assistant. Remote code formatting and Python in Excel
have the separate service flows above. Fork portable packages default to telemetry
off; configured builds can send pseudonymous usage telemetry.
It excludes workbook contents, file names, paths, arguments, and raw machine or
user names. A one-time migration event may include the previous name-derived
hash so older usage can be counted under the new random ID.
