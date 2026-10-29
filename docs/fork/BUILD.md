# Build portable fork packages

From a clean checkout, use Windows, PowerShell 7.4 or later, and the SDK pinned in `global.json`:

```powershell
pwsh -NoProfile -File scripts/Build-ForkRelease.ps1 -Version 3.0.0 -OutputDirectory artifacts/fork/3.0.0
```

The command builds both self-contained Windows x64 executables and packages each with the original license, fork notice and a standalone README. It writes `SHA256SUMS` and `RELEASE-MANIFEST.json`. It never publishes a release, invokes hosted CI, launches the packaged applications or runs desktop Excel tests. Normal .NET dependency restore may access the configured package feeds; `-NoRestore` is available when the required runtime-specific assets are already restored.

The inherited local CLI build hook stops background CLI services belonging to
this source checkout before rebuilding. It validates their executable paths and
process identities; it does not save workbooks or terminate Excel. Inspect and
finish your own active CLI work before building. Use a separate source checkout
for release preparation. Do not suppress that hook by pretending to be hosted CI.

Use `-DotNetExecutable` to select the pinned SDK executable explicitly. The builder checks its reported version. Existing final output directories are refused; choose a new directory rather than overwriting a release.

`-AllowDirtySource` is for development experiments only. Its manifest records `sourceDirty: true`; `sourceCommit` and `sourceTree` identify HEAD, not the uncommitted filesystem. Public release packages should come from the committed, clean source revision being tagged.

Build in a checkout with no other writers. The builder rechecks HEAD, its tree
and Git status before installing packages and refuses observed source changes.
Development builds with dirty source do not identify or reproduce uncommitted
file contents through their HEAD metadata.

The builder checks native PE architecture and version metadata. These are packaging checks, not protocol, lifecycle, workbook or numerical acceptance. Run the relevant tests and actual process checks separately and retain sanitized release verification evidence. Rebuild after any source change; the manifest must describe the packages actually uploaded.

For archive verification and switching versions, see [recovery](RECOVERY.md).

The canonical upstream npm/registry/extension metadata is at 2.6.1. The
portable fork builder explicitly applies 3.0.0 to all compiled assemblies and
release package metadata; it does not publish those other channels.

Native ADO interop generation also requires Microsoft's TlbImp 4.8.9037.0 from
the .NET Framework 4.8.1 SDK. Use the official Microsoft developer pack or a
reviewed isolated extraction of its signed tools. The local
`AdoTypeLibraryImporterPath` MSBuild property can select that executable without
changing global SDK resolution; validate it with
`scripts/Import-AdoDefinitions.ps1 -ImporterPath <TLBIMP_EXE> -ValidateOnly`.
Do not bypass compiler checks or change Office/security policy to build.
When using an isolated importer, set the process environment variable
`AdoTypeLibraryImporterPath` to that same reviewed executable before running
`Invoke-ExcelFreeTests.ps1`. The ADO definition tests forward it as a default
parameter to direct script calls; explicit importer arguments still take
precedence, and Microsoft signature/version checks remain mandatory. Without
this variable, direct script tests use the standard Developer Pack location.

The current `main` branch includes upstream 2.6.1. Existing fork 3.0.0 release
archives and their verification receipts retain their original upstream 2.6.0
base; source synchronization does not rebuild or replace published assets.
