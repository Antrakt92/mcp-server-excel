---
"excelmcp": patch
---

Background cleanup rechecks the worker before shutting it down. If a workbook opens while cleanup is pending, its session remains available.
