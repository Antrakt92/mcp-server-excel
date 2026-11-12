---
"excelmcp": patch
---

ExcelMcp retains the last recorded Excel exit when a helper stops during status refresh. A temporary lock on the helper’s registry record is reported and retried while other workers continue to be monitored.
