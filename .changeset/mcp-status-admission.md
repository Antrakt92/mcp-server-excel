---
"excelmcp": patch
---

Keep MCP file listing available when workbook operations fill the outstanding operation limit, so callers can inspect pending sessions. Status requests have their own bounded capacity.
