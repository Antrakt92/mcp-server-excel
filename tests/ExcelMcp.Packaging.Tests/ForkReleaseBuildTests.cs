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
}
