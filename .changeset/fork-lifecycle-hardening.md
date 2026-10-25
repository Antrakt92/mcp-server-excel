---
"excelmcp": patch
---

Publish the independent Antrakt fork based on upstream 2.5.3, with hardened MCP shutdown for stdin EOF, broken stdout and owning-client exit, one JSON-RPC stdout stream, and bounded outstanding requests through cancellation cleanup. Preserve upstream unsaved-workbook handoff and cancellation recovery. Add portable fork packages, release checksums, a release manifest and recovery guidance.

Bound CLI connection drain so retained idle clients cannot hold shutdown indefinitely. Avoid duplicate JSON parsing for standard MCP responses while preserving result compatibility. Route update checks and downloads to this fork.
