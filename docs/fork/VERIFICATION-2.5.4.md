# Fork 2.5.4 verification

Base: upstream v2.5.3 (`a92b11b94ab6fa04e7de129c4d22504cd873f82d`).
Local verification was performed on 2026-10-25 using .NET SDK 10.0.402,
Windows x64 build 10.0.26300.0 and installed Excel 16.0.20527.20088.

## Source checks

| Check | Result |
| --- | --- |
| Full Release solution build | Passed, zero warnings and errors |
| Excel-free CLI suite | 507 passed |
| Excel-free ComInterop suite | 221 passed |
| Excel-free Core suite | 748 passed |
| Excel-free MCP suite | 705 passed |
| Excel-free Service suite, after connection-drain repair | 188 passed |
| Script safety / PreCommit group, including fork selection | 222 passed |
| Fork builder refusal, source provenance and preservation guards | 5 passed |
| Real CLI startup, concurrent startup, stop and restart | 4 passed |
| Native MCP cancellation, unsaved handover/discard and independent ownership | 6 passed |
| Native MCP create/write/save/reopen workflow | 1 passed |

The 40 daemon-host cases, included in the Service suite above, cover bounded idle
connection shutdown, pending-handler ownership, response grace, request saturation
and per-session exclusion. The added MCP regressions cover admission retention,
factory/shutdown races, transport failures, parent identity, result parity and the
fork update channel. These focused cases are included in the suite totals.

Useful reproduction commands, from the repository root:

```powershell
dotnet restore Sbroenne.ExcelMcp.sln -p:NuGetAudit=true -p:NuGetAuditMode=all
dotnet build Sbroenne.ExcelMcp.sln -c Release --no-restore
./scripts/Invoke-ExcelFreeTests.ps1 -Group Fast
dotnet test tests/ExcelMcp.ScriptSafety.Tests -c Release --filter 'Feature=PreCommit'
dotnet test tests/ExcelMcp.Packaging.Tests -c Release --filter 'FullyQualifiedName~ForkReleaseBuildTests'
dotnet test tests/ExcelMcp.McpServer.Tests -c Release --filter 'FullyQualifiedName~McpSessionLifetimeTests|FullyQualifiedName~McpProcessOwnershipTests|FullyQualifiedName~McpServerSmokeTests.Smoke_WorkbookLifecycle_PersistsValues'
```

Run native tests on an interactive desktop. They create disposable synthetic
workbooks and explicitly close their own sessions. They are not a reason to
terminate another Excel instance or use a customer workbook as a fixture.

## Packaged process checks

Both Windows x64 self-contained archives were built and checked for executable
architecture, version metadata and their license/notice contents. The preflight
MCP package passed actual-process protocol initialization, 63-tool discovery,
empty session listing, a missing-file diagnostic and text/structured parity.
It remained alive during a 15-second idle interval and exited on stdin EOF.
Broken stdout with stdin retained and parent exit with a separately retained
stdin writer both terminated within the smoke's eight-second deadline.

Recheck the final extracted release executable using Python 3.11 or later:

```powershell
python scripts/tests/fork_protocol_smoke.py --exe '<extracted MCP executable>' --report protocol-results.json
```

The report contains the tested executable SHA-256. The release manifest records
the final packaged-binary checks separately from this source verification record.

## Limits

These checks cover the changed lifecycle and distribution boundaries. They do
not establish every inherited Excel operation, all Office versions, ARM64,
customer workbook calculation equivalence, print layouts or all-day resource
stability. There is no quantified overall speedup claim. No hosted Actions tests
were run, and no npm, NuGet, extension or MCP-registry package was published.
Release executables are unsigned; verify their hashes against the same release.
Older 2.4.1 measurements apply only to the [archival patch](LEGACY.md).
