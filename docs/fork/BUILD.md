# Build portable fork packages

From a clean checkout, use Windows, PowerShell 7.4 or later, and the SDK pinned in `global.json`:

```powershell
pwsh -NoProfile -File scripts/Build-ForkRelease.ps1 -Version 2.5.4 -OutputDirectory artifacts/fork/2.5.4
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
