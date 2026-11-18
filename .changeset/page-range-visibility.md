---
"excelmcp": minor
---

Use `offset` and `limit` with the `get-visibility` action in MCP `range_format_read` or CLI `rangeformat`. The result includes `totalDimensions` and, when another page is available, `nextOffset`. Scopes larger than 10,000 rows or columns require a page limit.
