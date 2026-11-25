---
"excelmcp": patch
---

Excel MCP requires absolute paths for Save As, Save Copy As, workbook exports, chart image exports, and sheet transfers between files. The CLI resolves relative paths for these commands from its working directory. A checked workbook copy also rejects a CLI response output path that points to its baseline workbook.
