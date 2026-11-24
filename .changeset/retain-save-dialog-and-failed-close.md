---
"excelmcp": patch
---

When a dialog blocks a workbook save, the session status and close history can report what kind of window was observed. If closing fails and Excel exits later, the controller releases the quarantined session after confirming its process identity and keeps the save and close outcomes unconfirmed.
