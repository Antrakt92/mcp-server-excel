---
"excelmcp": patch
---

Keep cancelled MCP requests counted against the outstanding operation limit while their worker dispatch remains pending. Cancelling the caller still ends its wait and signals the worker.
