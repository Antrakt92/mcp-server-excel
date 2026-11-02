---
"excelmcp": minor
---

Add the MCP `batch` tool for running a known sequence of commands in one call, including read-only sessions and bounded range pages. Each entry retains its normal result and error details; cancellation, timeout or unknown completion stops later commands without rolling back earlier edits.
