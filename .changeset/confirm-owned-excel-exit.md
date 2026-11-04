---
"excelmcp": patch
---

Confirm the exact owned Excel process has exited before normal session shutdown releases tracking. Keep an unconfirmed exit tracked and quarantined so a later close can reconcile cleanup without saving or closing the workbook again.
