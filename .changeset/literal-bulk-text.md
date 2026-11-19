---
"excelmcp": minor
---

Use `literalText` in batch JSON, `literal_text` in MCP, or `--literal-text true` in the CLI to preserve strings in range value writes and table appends. Formula-like text, initial apostrophes, dates supplied as strings, and leading zeros stay text. Numeric and Boolean inputs still write as their native types. Empty strings still write blank cells. Existing calls keep their current formula and date interpretation. A populated first row can be appended when Excel leaves an empty table without a data row after resizing.
