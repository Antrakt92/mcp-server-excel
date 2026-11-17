---
"excelmcp": patch
---

Keep a newly opened workbook session alive if its controller exits during startup. If another error occurs during `excelcli service stop`, the CLI still reports the original shutdown warning.
