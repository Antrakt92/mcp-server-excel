---
"excelmcp": patch
---

If Excel cannot select visible cells, the fallback checks individual rows and columns before reporting an empty result. It preserves the Excel error when it cannot verify the result within the visibility check limit.
