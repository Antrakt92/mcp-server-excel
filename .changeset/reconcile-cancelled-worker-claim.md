---
"excelmcp": patch
---

Cancelling workbook discovery before a worker claim is sent leaves the retained worker alone. Claims already sent can finish recording their ownership without starting workbook discovery after cancellation.
