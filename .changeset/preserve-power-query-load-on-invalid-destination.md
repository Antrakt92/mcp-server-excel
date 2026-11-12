---
"excelmcp": patch
---

Power Query `load-to` checks the worksheet destination before removing existing loads. An invalid destination or occupied target cell leaves the current worksheet and Data Model loads intact.
