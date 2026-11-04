---
"excelmcp": patch
---

Refresh the rendered worksheet view before taking the full-content screenshot. This prevents old pixels from hiding offscreen cells, embedded charts and newly scrolled tiles in normal, frozen and split views. Discard the preparatory bitmap, preserve existing view/workbook state, and reject observed window movement or resizing instead of applying an outdated crop.
