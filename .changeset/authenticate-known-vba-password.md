---
"excelmcp": major
---

Add `vba.unlock-project` and `vba.verify-password` for authenticating a locked VBA project with its known password. Accepted input opens access in the current editor session without changing protection settings, saving, or reopening the workbook. Verification distinguishes a recognized wrong password from an unknown outcome, and refuses to verify an already accessible project. The native driver validates the selected project and its expected dialogs, then closes those dialogs and restores the previous editor context.

CLI password actions now require JSON batch input on stdin, including the existing `vba.protect-project` action. Migrate direct password flags and password-bearing input files to `excelcli batch --input -`; ordinary file batches remain supported. This is an incompatible CLI workflow change. Passwords are passed exactly, and malformed batch input cannot echo its contents in parse diagnostics.
