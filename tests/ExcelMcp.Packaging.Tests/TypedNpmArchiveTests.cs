using System.Diagnostics;
using System.Formats.Tar;
using System.IO.Compression;
using System.Text;
using System.Text.Json.Nodes;
using Xunit;
using static Sbroenne.ExcelMcp.Tests.Infrastructure.PackagingScriptTestHelper;

namespace Sbroenne.ExcelMcp.Packaging.Tests;

[Trait("RequiresExcel", "false")]
[Trait("Feature", "Packaging")]
public sealed class TypedNpmArchiveTests
{
    [Theory]
    [InlineData("../escape.txt", TarEntryType.RegularFile)]
    [InlineData("package/../../escape.txt", TarEntryType.RegularFile)]
    [InlineData("package/link", TarEntryType.SymbolicLink)]
    [InlineData("package/link", TarEntryType.HardLink)]
    public async Task ArchiveExtraction_RejectsTraversalAndLinks(string name, TarEntryType kind)
    {
        var sandbox = NewSandbox();
        try
        {
            var archive = Path.Combine(sandbox, "unsafe.tgz");
            WriteArchive(archive, [(name, kind, Encoding.UTF8.GetBytes("unsafe"))]);
            var result = await ExpandAsync(archive, Path.Combine(sandbox, "extracted"));
            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains(kind == TarEntryType.RegularFile ? "Unsafe or duplicate" : "Unsupported", result.CombinedOutput, StringComparison.Ordinal);
            Assert.False(File.Exists(Path.Combine(sandbox, "escape.txt")));
        }
        finally { Directory.Delete(sandbox, true); }
    }

    [Fact]
    public async Task ArchiveExtraction_RejectsDuplicatePathsAndMalformedArchives()
    {
        var sandbox = NewSandbox();
        try
        {
            var archive = Path.Combine(sandbox, "duplicate.tgz");
            WriteArchive(archive, [
                ("package/payload", TarEntryType.RegularFile, Encoding.UTF8.GetBytes("first")),
                ("package/PAYLOAD", TarEntryType.RegularFile, Encoding.UTF8.GetBytes("second"))
            ]);
            var duplicate = await ExpandAsync(archive, Path.Combine(sandbox, "duplicates"));
            Assert.NotEqual(0, duplicate.ExitCode);
            Assert.Contains("Unsafe or duplicate", duplicate.CombinedOutput, StringComparison.Ordinal);
            File.WriteAllText(archive, "not a gzip archive");
            var malformed = await ExpandAsync(archive, Path.Combine(sandbox, "malformed"));
            Assert.NotEqual(0, malformed.ExitCode);
        }
        finally { Directory.Delete(sandbox, true); }
    }

    [Theory]
    [InlineData("valid")]
    [InlineData("name")]
    [InlineData("architecture")]
    [InlineData("machine")]
    [InlineData("license")]
    [InlineData("launcher")]
    [InlineData("dependency")]
    [InlineData("runtime-version")]
    public async Task ArchiveVerification_RequiresMatchingMetadataArchitectureAndContents(string mutation)
    {
        var sandbox = NewSandbox();
        try
        {
            var (executable, runtimeVersion) = ReadVersionedPe(mutation == "machine" ? (ushort)0xaa64 : (ushort)0x8664);
            var version = mutation == "runtime-version" ? MismatchedVersion(runtimeVersion) : runtimeVersion;
            var runtime = JsonNode.Parse($$$"""{"name":"@sbroenne/excelcli-win32-x64","version":"{{{version}}}","main":"excelcli.exe","os":["win32"],"cpu":["x64"]}""")!.AsObject();
            var launcher = JsonNode.Parse($$$"""{"name":"@sbroenne/excelcli","version":"{{{version}}}","bin":{"excelcli":"bin/excelcli.js"},"optionalDependencies":{"@sbroenne/excelcli-win32-x64":"{{{version}}}","@sbroenne/excelcli-win32-arm64":"{{{version}}}"}}""")!.AsObject();
            if (mutation == "name") { runtime["name"] = "@sbroenne/unrelated"; }
            if (mutation == "architecture") { runtime["cpu"] = new JsonArray("arm64"); }
            if (mutation == "dependency") { launcher["optionalDependencies"]!["@sbroenne/excelcli-win32-arm64"] = MismatchedVersion(version); }
            var runtimeEntries = new List<(string, TarEntryType, byte[])> {
                ("package/package.json", TarEntryType.RegularFile, Encoding.UTF8.GetBytes(runtime.ToJsonString())),
                ("package/excelcli.exe", TarEntryType.RegularFile, executable)
            };
            if (mutation != "license") { runtimeEntries.Add(("package/LICENSE", TarEntryType.RegularFile, Encoding.UTF8.GetBytes("MIT"))); }
            var launcherEntries = new List<(string, TarEntryType, byte[])> {
                ("package/package.json", TarEntryType.RegularFile, Encoding.UTF8.GetBytes(launcher.ToJsonString())),
                ("package/LICENSE", TarEntryType.RegularFile, Encoding.UTF8.GetBytes("MIT")),
                ("package/bin/excelcli.js", TarEntryType.RegularFile, Encoding.UTF8.GetBytes("fixture"))
            };
            if (mutation != "launcher") { launcherEntries.Add(("package/lib/launcher.js", TarEntryType.RegularFile, Encoding.UTF8.GetBytes("fixture"))); }
            var runtimePath = Path.Combine(sandbox, "runtime.tgz");
            var launcherPath = Path.Combine(sandbox, "launcher.tgz");
            WriteArchive(runtimePath, runtimeEntries);
            WriteArchive(launcherPath, launcherEntries);
            var result = await RunPowerShellFileAsync(Path.Combine(RepoRoot, "scripts", "Test-NpmPackages.ps1"),
                ["-Component", "Cli", "-Architecture", "x64", "-LauncherPackage", launcherPath, "-RuntimePackage", runtimePath, "-ArchiveOnly"]);
            if (mutation == "valid")
            {
                Assert.True(result.ExitCode == 0, result.CombinedOutput);
                Assert.Contains("Archive-only validation requested", result.CombinedOutput, StringComparison.Ordinal);
            }
            else
            {
                Assert.NotEqual(0, result.ExitCode);
                var expected = mutation switch
                {
                    "name" => "Unexpected npm package name",
                    "architecture" => "metadata does not match",
                    "machine" => "does not match x64",
                    "license" => "Missing license",
                    "launcher" => "Missing launcher file",
                    "runtime-version" => $"does not match {version}",
                    _ => "matching arm64 release version"
                };
                Assert.Contains(expected, result.CombinedOutput, StringComparison.Ordinal);
            }
        }
        finally { Directory.Delete(sandbox, true); }
    }

    [Fact]
    public async Task BuildNpmPackages_RejectsMismatchedExecutableVersionBeforePacking()
    {
        var sandbox = NewSandbox();
        try
        {
            var scripts = Directory.CreateDirectory(Path.Combine(sandbox, "scripts")).FullName;
            foreach (var file in new[] { "Build-NpmPackages.ps1", "PackageHelpers.ps1" })
            {
                File.Copy(Path.Combine(RepoRoot, "scripts", file), Path.Combine(scripts, file));
            }
            var executable = Path.Combine(sandbox, "excelcli.exe");
            var (payload, runtimeVersion) = ReadVersionedPe(0x8664);
            File.WriteAllBytes(executable, payload);
            var version = MismatchedVersion(runtimeVersion);
            var output = Path.Combine(sandbox, "artifacts", "npm");

            // Deliberately omit the npm sources so reaching stage construction cannot launch npm.
            var result = await RunPowerShellFileAsync(Path.Combine(scripts, "Build-NpmPackages.ps1"),
                ["-Component", "Cli", "-Version", version, "-RuntimeExecutable", executable, "-OutputDirectory", output]);

            Assert.NotEqual(0, result.ExitCode);
            Assert.Contains($"does not match {version}", result.CombinedOutput, StringComparison.Ordinal);
            Assert.False(Directory.Exists(output));
        }
        finally { Directory.Delete(sandbox, true); }
    }

    private static (byte[] Payload, string Version) ReadVersionedPe(ushort machine)
    {
        var assembly = typeof(TypedNpmArchiveTests).Assembly.Location;
        var version = FileVersionInfo.GetVersionInfo(assembly).ProductVersion!.Split('+')[0];
        var payload = File.ReadAllBytes(assembly);
        var offset = BitConverter.ToInt32(payload, 0x3c);
        BitConverter.GetBytes(machine).CopyTo(payload, offset + 4);
        return (payload, version);
    }

    private static string MismatchedVersion(string version) => version == "0.0.0-audit" ? "0.0.1-audit" : "0.0.0-audit";

    private static string NewSandbox() =>
        Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), $"ExcelMcp.NpmArchive.{Guid.NewGuid():N}")).FullName;

    private static Task<ProcessResult> ExpandAsync(string archive, string destination) =>
        RunPowerShellCommandAsync($"""
            . '{Path.Combine(RepoRoot, "scripts", "PackageHelpers.ps1")}'
            Expand-PackageTar '{archive}' '{destination}'
            """);

    private static void WriteArchive(string path, IEnumerable<(string Name, TarEntryType Kind, byte[] Contents)> entries)
    {
        using var file = File.Create(path);
        using var compressed = new GZipStream(file, CompressionLevel.Fastest);
        using var writer = new TarWriter(compressed);
        foreach (var (name, kind, contents) in entries)
        {
            var entry = new PaxTarEntry(kind, name);
            using var data = new MemoryStream(contents);
            if (kind == TarEntryType.RegularFile) { entry.DataStream = data; }
            else { entry.LinkName = "../escape.txt"; }
            writer.WriteEntry(entry);
        }
    }
}
