---
"excelmcp": patch
---

Do not rewrite an identical native observation that this worker already persisted. Repeated teardown no longer fails on an unnecessary second write after release; a richer observation still requires a successful write before replacing the stored state.
