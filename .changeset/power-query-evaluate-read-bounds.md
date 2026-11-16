---
"excelmcp": patch
---

Power Query evaluate reads Excel results in smaller blocks and rejects responses that exceed the result budget. Use `Table.FirstN` or `Table.SelectColumns` in the M expression to preview a large query.
