---
"excelmcp": patch
---

`range patch-cells` accepts Excel's confirmed scalar formula write when Excel adds an implicit intersection marker during readback. It still detects later formula changes and preserves adjacent cells.
