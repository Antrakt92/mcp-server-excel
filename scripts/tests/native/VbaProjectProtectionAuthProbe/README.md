# Synthetic VBA project authentication probe

Test-only Windows helper for verifying a password set through ExcelMCP. It has
no product entry point, COM reference, package dependency or Excel shutdown code.
The caller owns one disposable, otherwise idle Excel fixture and its shutdown.

Build with the repository-pinned SDK:

```powershell
dotnet build scripts/tests/native/VbaProjectProtectionAuthProbe/VbaProjectProtectionAuthProbe.csproj -c Release
```

Invoke the built executable with four named arguments:

```text
--identity <identity.json> --mode correct|wrong|observe --ready <new-ready.json> --result <new-result.json>
```

The identity JSON contains `processId` and `startedAt` (the process creation time
as a UTC Windows FileTime). The helper retains the exact process handle and
checks PID, creation time, executable name and current Windows session. There
must be no existing standard dialog in that Excel process before readiness.

Send the synthetic password to submit as **UTF-8 stdin**, then close stdin. Do
not append a newline, place the password in arguments or persist it in a file.
The helper rejects empty/control-character input and input exceeding 1,024 UTF-16
code units. Both modes submit that exact value; mode only selects the expected
authentication outcome. For wrong mode, the caller supplies a distinct password
with the same length and Unicode category as the correct password, so the probe
can detect lossy character conversion rather than merely reject ASCII input.

`observe` inventories an owned synthetic challenge's caption and control IDs,
then cancels it without reading stdin or submitting a password. Use this mode
to establish a native dialog signature before a production-driver acceptance
run. Its receipt reports `passwordSubmitted:false`; it does not verify a password.
Captions may contain the synthetic project name. Keep this test helper on
disposable fixtures and keep its receipts out of customer/public reports.

Wait for the newly created ready JSON before asking the product to open the
locked fixture's Project Properties command. Correct mode expects the standard
Project Properties dialog. Wrong mode expects the exact English `Project Locked`
dialog with `Invalid password`, dismisses its OK button (ID 2), then cancels the
authentication dialog (ID 2). Other locales/signatures fail closed.

Only addressed `WM_CHAR` input and asynchronous button clicks are used. There
are no foreground keystrokes, coordinates or password/control-value dumps. The
helper cancels its introduced Project Properties dialog (tab ID 12320, Cancel
ID 2) and independently attempts expected-dialog cleanup on failure. Unknown
dialogs remain for the caller to inspect; they are never answered automatically.

The new result JSON records `success`, `expectedOutcomeObserved`,
`cleanupConfirmed`, identity confirmation, mode and nonsecret error codes. The
nullable `passwordControlUnicode` and `windowsAnsiCodePage` fields record the
confirmed authentication edit's window type and the Windows ANSI code page;
they remain null if the edit was not confirmed. No password text is recorded.
The nullable `passwordControlSupportsUnichar` field records the response to a
bounded `WM_UNICHAR` / `UNICODE_NOCHAR` support query, which inserts no text.
Exit 0 requires the expected authentication outcome and confirmed dialog cleanup.
The helper never closes a workbook, quits Excel or terminates a process. The
caller must separately prove the product's workbook/worker/Excel shutdown and
retain failed receipts.
