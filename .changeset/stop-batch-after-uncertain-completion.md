---
"excelmcp": patch
---

CLI and MCP batches stop when workbook completion is uncertain even if recording the worker response also fails. The error retains the original response so you can inspect the affected session before continuing.
