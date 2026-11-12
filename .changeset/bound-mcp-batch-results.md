---
"excelmcp": patch
---

MCP accepts up to 128 commands per batch and stops before another command once results reach 8 MiB of JSON. The reply keeps every processed result, including the one that reached the limit, and reports when commands remain. Split larger batches and continue from the reported count after checking the results.
