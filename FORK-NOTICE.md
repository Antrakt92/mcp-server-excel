# Fork notice

**ExcelMCP — Antrakt Fork** is an independent fork maintained by [Antrakt92](https://github.com/Antrakt92) at [Antrakt92/mcp-server-excel](https://github.com/Antrakt92/mcp-server-excel).

The 3.0.0 fork release line is based on upstream ExcelMcp v2.6.0 from [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel). ExcelMcp was created by **Stefan Broenner** and its contributors. The original README is preserved as [README.upstream.md](README.upstream.md).

The original [MIT license](LICENSE), including its copyright notice, remains unchanged. Fork additions are distributed under that license. No affiliation with or endorsement by the upstream author, Microsoft, OpenAI or other vendors is claimed.

Fork changes cover MCP transport, request admission, bounded CLI connection drain, persistent audit sessions, bounded range inspection and guarded VBA workflows. They retain upstream workbook ownership, unsaved-work handoff and cancellation recovery. The historical 2.4.1 patch is a separate legacy line; it does not define current session behavior. See [the fork overview](docs/fork/OVERVIEW.md).
