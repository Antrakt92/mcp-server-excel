---
"excelmcp": patch
---

Power Query refresh waits honor cancellation even if Excel has already become idle. If Excel also fails to cancel its native refresh, the failure is reported alongside the original cancellation.
