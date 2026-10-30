---
"excelmcp": patch
---

Avoid repeated JSON normalization when converting Service responses into generated MCP replies. Preserve errors, warnings, partial state, telemetry and matching text/structured content while reducing conversion work and allocations.
