---
"excelmcp": major
---

CLI daemons now use a separate default pipe for each Windows session. ExcelMcp checks the Windows session before sending or accepting commands, including on custom pipes. Use the previous CLI or an explicit pipe override in the same session to finish work on an already-running legacy daemon.
