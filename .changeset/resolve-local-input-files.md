---
"excelmcp": patch
---

CLI code, data, and formula files now resolve relative to the directory where you run the CLI. MCP requires absolute Windows paths for these files. A batch can use an input file created by an earlier command because ExcelMCP checks file existence when the command runs.
