---
"excelmcp": patch
---

Retrying a close with saving enabled succeeds only when the service confirmed that the earlier close saved the workbook. A close without saving, a lost worker response, or failed teardown cannot be mistaken for a saved close.
