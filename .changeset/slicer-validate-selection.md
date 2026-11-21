---
"excelmcp": patch
---

Table and non-OLAP PivotTable slicers leave the filter unchanged and return an error when none of the requested item names match. Requests with at least one match still apply the matching names. Table slicers report how many items are selected after the update.
