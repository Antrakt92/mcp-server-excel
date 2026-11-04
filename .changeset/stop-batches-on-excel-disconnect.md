---
"excelmcp": patch
---

Stop CLI and MCP batches after an Excel automation disconnect, even when continuing after ordinary errors is enabled. Retained-session errors now report unknown completion while preserving observed VBA edit and protection state, so later commands cannot run under the assumption that the previous operation finished.
