---
"excelmcp": patch
---

The CLI resolves code, data, formula, image, QueryTable text, and workbook theme input paths from its working directory. MCP requires absolute Windows paths for these inputs. A batch can use an input file created by an earlier command because ExcelMCP checks file existence when the command runs.
