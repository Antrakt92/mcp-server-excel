---
"excelmcp": patch
---

Ctrl+C now stops `excelcli batch` from starting additional commands and preserves the results already written. If an interrupted command may have reached Excel, the batch reports that its outcome is unknown.
