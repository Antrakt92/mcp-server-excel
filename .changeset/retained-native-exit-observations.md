---
"excelmcp": minor
---

Retain an identity-checked native process handle throughout Excel ownership and record confirmed exit state before releasing it. Disconnect errors include native process diagnostics, and session lists retain the last 16 observed exits for that controller, even after the sessions are gone. Unknown exit codes remain null; process creation times use exact decimal strings to avoid JavaScript rounding.

Failed diagnostic writes keep ownership retryable. Closing an already exited normal workbook without saving can complete its teardown without a COM readiness probe; a requested save is never reported as successful through that path. Attached and preserved Excel applications keep their existing ownership rules.
