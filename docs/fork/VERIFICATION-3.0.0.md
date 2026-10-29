# Fork 3.0.0 verification

Base: upstream v2.6.0 (`49e04ec3a2fadce6c2b3fc3c7c8678165799fa6d`).
Implementation checks were performed locally on 2026-10-29 with .NET SDK
10.0.402, Windows x64 and installed Excel 16.0. No hosted Actions runs were used.

## Source checks

| Check | Result |
| --- | --- |
| Upstream integration full Release solution build | Passed, zero warnings/errors |
| Selected Release builds before each implementation commit | Passed, zero warnings/errors |
| Final normal Excel-free CLI / ComInterop / Core suites | 813 / 222 / 898 passed |
| Final normal Excel-free MCP / Service suites | 753 / 262 passed |
| Final normal skill-generation / packaging suites | 16 / 98 passed |
| Final normal script-safety suite | 286 passed after isolated importer routing repair |
| VS Code extension tests / typecheck / lint | 29 passed / passed / passed |
| Extension dependency audit and signature validation | 0 vulnerabilities; no invalid or missing signatures |
| Actual source-channel VSIX package checks | x64 and ARM64 passed |
| VBA source/invocation units | 68 passed |
| Service VBA class and source-failure transport | 41 passed, includes desktop Excel |
| Persistent audit session native regressions | 3 passed |
| Bounded range units | 18 passed |
| Bounded range native regressions | 10 passed |
| Caller cancellation and total deadlines | 10 passed; new regressions failed before repair |
| Final MCP discovery contract regression | 58 passed |
| Benchmark/packaged-smoke/protocol-cleanup Python regressions | 8 / 12 / 5 passed |
| Formula lexical units / native cases | 26 / 21 passed |
| Number-format limits and numeric trace native regressions | 70 passed |
| Physical workbook identity ComInterop / Service regressions | 222 / 5 passed; 3 Service cases use native Excel |
| CLI worker build-closure process regressions | 40 passed |

The extension dependency fix replaces MSAL Node 6.0.1 with 7.1.0 under Azure
Identity 4.14.0. Adversarial response-metadata checks failed on the old version;
the patched boundary and legitimate request adaptation pass. A real state
mismatch is refused before any network call. No live tenant authentication was
performed, and exploit reachability through normal vsce use was not established.
The portable .NET runtimes do not consume that npm development dependency.

These suites overlap. Native fixtures are synthetic and explicitly close their
owned sessions; source results do not establish final packaged-binary identity.
The audit cases include mutation refusal, unchanged saved hashes, sticky audit
identity after Save As and unsaved interactive handover. Range cases include
canonical errors versus legitimate numeric constants, long-text response limits,
manual cached values, page boundaries and mixed displayed/stored formatting.
VBA cases cover stale fingerprints, normalized source read-back, typed results,
output limits after completed execution and structured partial failure state.

## Reproduce packaged checks

Build from clean committed source using [BUILD.md](BUILD.md). Extract both ZIPs
into separate directories, then run with Python 3.11 or later:

```powershell
python scripts/tests/fork_protocol_smoke.py --exe '<MCP_EXE>' --report protocol-results.json --idle-seconds 15
python scripts/tests/fork_packaged_excel_smoke.py --mcp-exe '<MCP_EXE>' --cli-exe '<CLI_EXE>' --report native-results.json
python scripts/tests/fork_cli_batch_benchmark.py --cli '<CLI_EXE>' --source-revision '<SOURCE_COMMIT>' --output cli-performance.json
```

The first check never opens Excel: it verifies actual stdio initialization,
discovery, result parity, EOF, broken stdout while stdin remains open, parent
exit with a retained stdin writer and bounded idle resources. The native check
uses disposable XLSM files through each public entry point, including save/reopen,
audit refusal, hash preservation, bounded reads and MCP guarded VBA invocation.
The benchmark uses a private daemon and empty session lists; its three warmed
rounds compare eight separate CLI processes with one batch process. It records
startup, each subprocess, shutdown, cleanup and unchanged binary hashes. Native
smoke cleanup retains exact owned process handles and checks worker/Excel exit;
failed cases preserve sanitized failure and cleanup evidence. Idle resource
counters are observations of one bounded window, not leak assertions.

Each published release supplies `VERIFICATION-RESULTS.json` with the exact
packaged executable hashes and observed outcomes. `RELEASE-MANIFEST.json` records
what the builder checked; its test fields explicitly say the packaging command
does not run runtime tests. Match the verification receipt and manifest to the
same ZIPs and source revision. Do not substitute an older installation receipt.

## Limits

These checks do not establish every inherited operation, every Office version,
customer workbook numerical equivalence, operational printing speed or all-day
resource stability. CLI batch timings measure process/connection overhead.
They do not measure Excel calculation or printing. No npm, NuGet, extension or
MCP-registry identity is published by this portable fork workflow. Source-channel
version metadata remains upstream 2.6.0; portable packages explicitly use 3.0.0.