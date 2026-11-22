---
"excelmcp": minor
---

Session lists now include recent close attempts by session ID and exact Excel process identity. Each record reports separately whether the worker confirmed close, whether the controller released its claims, whether save was confirmed, and whether that Excel process exited. A later confirmation can update the record while preserving the previous close failure.
