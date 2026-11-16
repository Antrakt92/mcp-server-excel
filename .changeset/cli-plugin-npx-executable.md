---
"excelmcp": patch
---

The Excel CLI plugin launcher now works with Node installations that expose `npx.exe`, including Volta on Windows. It still passes JSON and other quoted arguments through unchanged.
