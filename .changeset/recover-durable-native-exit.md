---
"excelmcp": patch
---

Keep an already-recorded Excel exit in the controller's diagnostic history when its worker dies before acknowledging release. Recovery requires the exact known worker, controller, generation, build and Excel identities; it does not recreate sessions or grant ownership.
