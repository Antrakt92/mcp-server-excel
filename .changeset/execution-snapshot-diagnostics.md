---
"excelmcp": patch
---

Report operation and visibility-snapshot failures separately in structured diagnostics. Responses preserve the original exception and cancellation token, identify whether the callback returned, and retain observed VBA state. Incomplete or ambiguous exception chains are marked incomplete instead of assigning an operation outcome.
