---
"excelmcp": patch
---

Render frozen and split screenshot panes through the standard whole-window path so scrolling captures the current viewport rather than an older frame. Keep the existing full-content renderer for ordinary zoomed and tiled captures. Add a strict non-active-sheet regression that checks pixel geometry, both sheet views and unchanged cell contents.
