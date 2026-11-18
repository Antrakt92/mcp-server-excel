---
"excelmcp": patch
---

The CLI rejects `--output` when it points to the destination for `workbook save-as`, `workbook save-copy-as`, `workbook export-fixed-format`, or `chart export-image`. The same check protects both workbooks in `sheet copy-to-file` and `sheet move-to-file`.
