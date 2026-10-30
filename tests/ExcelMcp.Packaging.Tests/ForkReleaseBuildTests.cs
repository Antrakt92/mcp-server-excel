using Xunit;
using static Sbroenne.ExcelMcp.Tests.Infrastructure.PackagingScriptTestHelper;

namespace Sbroenne.ExcelMcp.Packaging.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "Packaging")]
[Trait("Speed", "Fast")]
public sealed class ForkReleaseBuildTests
{
    private static string Builder => Path.Combine(RepoRoot, "scripts", "Build-ForkRelease.ps1");

    [Theory(Timeout = 45000)]
    [InlineData("conflicting-lightweight", true)]
    [InlineData("conflicting-annotated", true)]
    [InlineData("conflicting-clean", true)]
    [InlineData("matching-lightweight", false)]
    [InlineData("matching-annotated", false)]
    [InlineData("absent", false)]
    public async Task VersionTagAdmission_RejectsDifferentSourceBeforeDirtyCheckAndAllowsExactHead(
        string tagCase, bool conflict)
    {
        var sandbox = CreateSandbox("fork-version-tag");
        try
        {
            var runner = Path.Combine(sandbox, "admission.ps1");
            var builder = Builder.Replace("'", "''", StringComparison.Ordinal);
            var helpers = Path.Combine(RepoRoot, "scripts", "PackageHelpers.ps1")
                .Replace("'", "''", StringComparison.Ordinal);
            File.WriteAllText(runner, $$"""
                $ErrorActionPreference = 'Stop'
                Set-StrictMode -Version Latest
                . '{{helpers}}'
                $fixture = Join-Path $PSScriptRoot 'repo'
                $scripts = Join-Path $fixture 'scripts'
                [IO.Directory]::CreateDirectory($scripts) | Out-Null
                Copy-Item -LiteralPath '{{builder}}' -Destination (Join-Path $scripts 'Build-ForkRelease.ps1')
                Copy-Item -LiteralPath '{{helpers}}' -Destination (Join-Path $scripts 'PackageHelpers.ps1')
                function Git {
                    param([string[]]$Arguments)
                    (Invoke-PackageCommand $fixture 'git' $Arguments -TimeoutSeconds 10).Trim()
                }
                $null = Git -Arguments @('init', '--quiet')
                $null = Git -Arguments @('config', 'core.hooksPath', (Join-Path $PSScriptRoot 'no-hooks'))
                $null = Git -Arguments @('config', 'user.name', 'Fork Release Fixture')
                $null = Git -Arguments @('config', 'user.email', 'fork-fixture@example.invalid')
                $null = Git -Arguments @('config', 'commit.gpgsign', 'false')
                $null = Git -Arguments @('config', 'tag.gpgsign', 'false')
                $payload = Join-Path $fixture 'payload.txt'
                [IO.File]::WriteAllText($payload, 'committed source')
                $null = Git -Arguments @('add', '.')
                $null = Git -Arguments @('commit', '--quiet', '-m', 'initial fixture source')
                $tagCase = '{{tagCase}}'
                if ($tagCase -ne 'absent') {
                    if ($tagCase.EndsWith('annotated')) {
                        $null = Git -Arguments @('tag', '-a', 'v3.0.1', '-m', 'fixture release')
                    } else {
                        $null = Git -Arguments @('tag', 'v3.0.1')
                    }
                }
                if ($tagCase.StartsWith('conflicting')) {
                    $null = Git -Arguments @('commit', '--quiet', '--allow-empty', '-m', 'later source')
                    if ($tagCase -ne 'conflicting-clean') {
                        [IO.File]::WriteAllText($payload, 'uncommitted sentinel')
                    }
                }
                $headBefore = Git -Arguments @('rev-parse', 'HEAD')
                $refsBefore = Git -Arguments @('show-ref')
                $statusBefore = Git -Arguments @('status', '--porcelain=v1', '--untracked-files=normal')
                $payloadBefore = [IO.File]::ReadAllText($payload)
                $outsideSentinel = Join-Path $PSScriptRoot 'sentinel.txt'
                [IO.File]::WriteAllText($outsideSentinel, 'preserve prior evidence')
                $output = Join-Path $PSScriptRoot 'release'
                $missingSdk = Join-Path $PSScriptRoot 'must-not-run-dotnet.exe'
                $failure = $null
                try {
                    & (Join-Path $scripts 'Build-ForkRelease.ps1') -Version '3.0.1' `
                        -OutputDirectory $output -DotNetExecutable $missingSdk
                } catch {
                    $failure = $_.Exception.Message
                }
                if ($null -eq $failure) { throw 'Builder unexpectedly accepted the missing SDK.' }
                if ([IO.Path]::Exists($output)) { throw 'Admission failure created release output.' }
                if ([IO.File]::ReadAllText($outsideSentinel) -cne 'preserve prior evidence' -or
                    [IO.File]::ReadAllText($payload) -cne $payloadBefore) { throw 'Admission changed sentinel evidence.' }
                if ((Git -Arguments @('rev-parse', 'HEAD')) -cne $headBefore -or
                    (Git -Arguments @('show-ref')) -cne $refsBefore -or
                    (Git -Arguments @('status', '--porcelain=v1', '--untracked-files=normal')) -cne $statusBefore) {
                    throw 'Admission changed source, refs, or checkout state.'
                }
                Write-Output "ADMISSION-ERROR: $failure"
                Write-Output 'PASS source, refs, sentinels, and absent release output preserved'
                """);
            var result = await RunPowerShellFileAsync(runner, [], timeoutMs: 30000);
            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            Assert.Contains("PASS source, refs, sentinels, and absent release output preserved",
                result.CombinedOutput, StringComparison.Ordinal);
            if (conflict)
            {
                Assert.Contains("Version 3.0.1 is already tagged at", result.CombinedOutput, StringComparison.Ordinal);
                Assert.DoesNotContain("Release source is dirty", result.CombinedOutput, StringComparison.Ordinal);
                Assert.DoesNotContain("must-not-run-dotnet.exe", result.CombinedOutput, StringComparison.Ordinal);
            }
            else
            {
                Assert.Contains("must-not-run-dotnet.exe", result.CombinedOutput, StringComparison.Ordinal);
                Assert.DoesNotContain("already tagged at", result.CombinedOutput, StringComparison.Ordinal);
            }
        }
        finally
        {
            foreach (var file in Directory.EnumerateFiles(sandbox, "*", SearchOption.AllDirectories))
            {
                File.SetAttributes(file, FileAttributes.Normal);
            }
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Theory(Timeout = 30000)]
    [InlineData("invalid", "valid package version")]
    [InlineData("2.5.4-preview", "stable three-part numeric version")]
    public async Task InvalidVersion_RejectsBeforeSdkOrBuild(string version, string expectedError)
    {
        var sandbox = CreateSandbox("fork-invalid-version");
        var output = Path.Combine(sandbox, "release");
        try
        {
            var result = await RunPowerShellFileAsync(Builder,
                ["-Version", version, "-OutputDirectory", output,
                 "-DotNetExecutable", Path.Combine(sandbox, "must-not-run-dotnet.exe")], timeoutMs: 15000);
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains(expectedError, result.CombinedOutput, StringComparison.Ordinal);
            Assert.False(Directory.Exists(output));
            Assert.Empty(Directory.GetFileSystemEntries(sandbox));
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Fact(Timeout = 30000)]
    public async Task SourceTreeOutput_RejectsAndPreservesSentinel()
    {
        var output = Path.Combine(RepoRoot, "src", $"fork-output-safety-{Guid.NewGuid():N}");
        Directory.CreateDirectory(output);
        var sentinel = Path.Combine(output, "sentinel.txt");
        File.WriteAllText(sentinel, "retained source sentinel");
        try
        {
            var result = await RunPowerShellFileAsync(Builder,
                ["-Version", "2.5.4", "-OutputDirectory", output,
                 "-DotNetExecutable", Path.Combine(output, "must-not-run-dotnet.exe")], timeoutMs: 15000);
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("Unsafe package output directory", result.CombinedOutput, StringComparison.Ordinal);
            Assert.Equal("retained source sentinel", File.ReadAllText(sentinel));
            Assert.Equal(sentinel, Assert.Single(Directory.GetFileSystemEntries(output)));
        }
        finally { DeleteDirectoryIfExists(output); }
    }

    [Fact(Timeout = 30000)]
    public async Task ExistingFinalOutput_RejectsAndPreservesSentinel()
    {
        var sandbox = CreateSandbox("fork-existing-output");
        var output = Path.Combine(sandbox, "release");
        Directory.CreateDirectory(output);
        var sentinel = Path.Combine(output, "sentinel.txt");
        File.WriteAllText(sentinel, "last verified release");
        try
        {
            var result = await RunPowerShellFileAsync(Builder,
                ["-Version", "2.5.4", "-OutputDirectory", output,
                 "-DotNetExecutable", Path.Combine(sandbox, "must-not-run-dotnet.exe")], timeoutMs: 15000);
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains("Use a new release output directory", result.CombinedOutput, StringComparison.Ordinal);
            Assert.Equal("last verified release", File.ReadAllText(sentinel));
            Assert.Equal(sentinel, Assert.Single(Directory.GetFileSystemEntries(output)));
            Assert.Equal(output, Assert.Single(Directory.GetFileSystemEntries(sandbox)));
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

    [Fact(Timeout = 45000)]
    public async Task SourceGuard_RejectsWorktreeAndCommittedChangesWithoutTouchingOutput()
    {
        var sandbox = CreateSandbox("fork-source-guard");
        try
        {
            var runner = Path.Combine(sandbox, "guard.ps1");
            var builder = Builder.Replace("'", "''", StringComparison.Ordinal);
            var helpers = Path.Combine(RepoRoot, "scripts", "PackageHelpers.ps1")
                .Replace("'", "''", StringComparison.Ordinal);
            File.WriteAllText(runner, $$"""
                $ErrorActionPreference = 'Stop'
                Set-StrictMode -Version Latest
                . '{{helpers}}'
                $tokens = $null; $errors = $null
                $ast = [Management.Automation.Language.Parser]::ParseFile('{{builder}}', [ref]$tokens, [ref]$errors)
                if ($errors.Count) { throw 'Builder parse failed.' }
                foreach ($name in @('Read-ForkGit', 'Assert-ForkSourceUnchanged')) {
                    $functions = @($ast.FindAll({ param($node)
                        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
                    }, $true))
                    if ($functions.Count -ne 1) { throw "Expected exactly one product function: $name" }
                    . ([scriptblock]::Create($functions[0].Extent.Text))
                }
                $root = Join-Path $PSScriptRoot 'repo'
                [IO.Directory]::CreateDirectory($root) | Out-Null
                $null = Read-ForkGit -GitArguments @('init', '--quiet')
                $null = Read-ForkGit -GitArguments @('config', 'core.hooksPath', (Join-Path $PSScriptRoot 'no-hooks'))
                $payload = Join-Path $root 'payload.txt'
                [IO.File]::WriteAllText($payload, 'original')
                $null = Read-ForkGit -GitArguments @('add', 'payload.txt')
                $commitArgs = @('-c', 'commit.gpgsign=false', '-c', 'user.name=Fork Release Fixture',
                    '-c', 'user.email=fork-fixture@example.invalid', 'commit', '--quiet', '-m')
                $null = Read-ForkGit -GitArguments ($commitArgs + 'initial')
                $commit = Read-ForkGit -GitArguments @('rev-parse', 'HEAD')
                $tree = Read-ForkGit -GitArguments @('rev-parse', 'HEAD^{tree}')
                $sourceStatus = Read-ForkGit -GitArguments @('status', '--porcelain=v1', '--untracked-files=normal')
                Assert-ForkSourceUnchanged
                $output = Join-Path $PSScriptRoot 'release'
                [IO.Directory]::CreateDirectory($output) | Out-Null
                $sentinel = Join-Path $output 'sentinel.txt'
                [IO.File]::WriteAllText($sentinel, 'verified prior output')
                function Assert-SourceChangeRejected {
                    $rejected = $false
                    try { Assert-ForkSourceUnchanged }
                    catch {
                        if ($_.Exception.Message -notlike '*Source changed during packaging; release output was not installed.*') { throw }
                        $rejected = $true
                    }
                    if (-not $rejected) { throw 'Changed source was accepted.' }
                    if ([IO.File]::ReadAllText($sentinel) -cne 'verified prior output' -or
                        @(Get-ChildItem -LiteralPath $output -Force).Count -ne 1) { throw 'Prior output changed.' }
                }
                [IO.File]::WriteAllText($payload, 'uncommitted change')
                Assert-SourceChangeRejected
                [IO.File]::WriteAllText($payload, 'original')
                Assert-ForkSourceUnchanged
                $null = Read-ForkGit -GitArguments ($commitArgs + @('new revision', '--allow-empty'))
                Assert-SourceChangeRejected
                Write-Output 'PASS clean source, worktree change, restored content, and changed HEAD guard'
                """);
            var result = await RunPowerShellFileAsync(runner, [], timeoutMs: 30000);
            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            Assert.Contains("PASS clean source, worktree change, restored content, and changed HEAD guard",
                result.CombinedOutput, StringComparison.Ordinal);
        }
        finally
        {
            // Git object files are read-only on Windows; normalize only this owned fixture.
            foreach (var file in Directory.EnumerateFiles(sandbox, "*", SearchOption.AllDirectories))
            {
                File.SetAttributes(file, FileAttributes.Normal);
            }
            DeleteDirectoryIfExists(sandbox);
        }
    }

    [Fact(Timeout = 30000)]
    public async Task VersionMetadata_RequiresExactSourceCommit()
    {
        var sandbox = CreateSandbox("fork-version-metadata");
        try
        {
            var runner = Path.Combine(sandbox, "metadata.ps1");
            var builder = Builder.Replace("'", "''", StringComparison.Ordinal);
            File.WriteAllText(runner, $$"""
                $ErrorActionPreference = 'Stop'
                Set-StrictMode -Version Latest
                $tokens = $null; $errors = $null
                $ast = [Management.Automation.Language.Parser]::ParseFile('{{builder}}', [ref]$tokens, [ref]$errors)
                if ($errors.Count) { throw 'Builder parse failed.' }
                $guards = @($ast.FindAll({ param($node)
                    $node -is [Management.Automation.Language.IfStatementAst] -and
                    $node.Clauses[0].Item1.Extent.Text.Contains('$metadata.FileVersion')
                }, $true))
                if ($guards.Count -ne 1) { throw 'Expected exactly one executable metadata guard.' }
                $guard = [scriptblock]::Create($guards[0].Extent.Text)
                $Version = '3.0.1'
                $assemblyVersion = "$Version.0"
                $commit = '0123456789012345678901234567890123456789'
                $component = @{ Name = 'Synthetic' }
                $cases = @(
                    @{ Name='matching'; File=$assemblyVersion; Product="$Version+$commit"; Reject=$false },
                    @{ Name='wrong-commit'; File=$assemblyVersion; Product="$Version+ffffffffffffffffffffffffffffffffffffffff"; Reject=$true },
                    @{ Name='missing-stamp'; File=$assemblyVersion; Product=$Version; Reject=$true },
                    @{ Name='extra-suffix'; File=$assemblyVersion; Product="$Version+$commit.extra"; Reject=$true },
                    @{ Name='wrong-file-version'; File='3.0.0.0'; Product="$Version+$commit"; Reject=$true },
                    @{ Name='wrong-product-version'; File=$assemblyVersion; Product="3.0.0+$commit"; Reject=$true },
                    @{ Name='empty-product-version'; File=$assemblyVersion; Product=''; Reject=$true }
                )
                foreach ($case in $cases) {
                    $metadata = [pscustomobject]@{ FileVersion=$case.File; ProductVersion=$case.Product }
                    $rejected = $false
                    try { & $guard }
                    catch {
                        if ($_.Exception.Message -notlike '*incorrect release version metadata.*') { throw }
                        $rejected = $true
                    }
                    if ($rejected -ne $case.Reject) { throw "Unexpected metadata admission for $($case.Name)." }
                }
                Write-Output "PASS exact source commit metadata cases=$($cases.Count)"
                """);
            var result = await RunPowerShellFileAsync(runner, [], timeoutMs: 20000);
            Assert.True(result.ExitCode == 0, result.CombinedOutput);
            Assert.Contains("PASS exact source commit metadata cases=7", result.CombinedOutput, StringComparison.Ordinal);
        }
        finally { DeleteDirectoryIfExists(sandbox); }
    }

}
