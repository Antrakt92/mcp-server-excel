using System.ComponentModel;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Fact]
    public async Task SaveCopyAs_CheckedReportsNativePackageChangesAndPublishesOnlyWhenAllowed()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            string source = Path.Combine(directory, "owned.xlsx");
            string baseline = Path.Combine(directory, "baseline.xlsx");
            string target = Path.Combine(directory, "checked.xlsx");
            string session = await CreateSessionAsync(service, source);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "original marker");
            RequireSuccess(await SaveCheckedCopyAsync(service, session, baseline));
            byte[] original = await File.ReadAllBytesAsync(baseline);
            string[] baselineParts;
            using (var archive = ZipFile.OpenRead(baseline))
                baselineParts = archive.Entries.Select(entry => entry.FullName).ToArray();

            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "range.set-values",
                SessionId = session,
                Args = """{"sheetName":"Sheet1","rangeAddress":"A1:B1","overwritePolicy":"allow","values":[["changed marker",42]]}"""
            }));
            string[] noParts = [];
            var refused = await SaveCheckedCopyAsync(service, session, target, new { baselinePath = baseline, allowedParts = noParts });
            using var refusedDocument = JsonDocument.Parse(refused.Result!);
            var refusal = refusedDocument.RootElement;
            Assert.False(refusal.GetProperty("success").GetBoolean());
            Assert.Equal("Conflict", refusal.GetProperty("errorCategory").GetString());
            Assert.False(refusal.GetProperty("filePublished").GetBoolean());
            var refusedComparison = refusal.GetProperty("packageComparison");
            Assert.False(refusedComparison.GetProperty("matchesAllowlist").GetBoolean());
            var refusedWorksheet = CheckedWorksheetChange(refusedComparison);
            Assert.False(refusedWorksheet.GetProperty("allowed").GetBoolean());
            Assert.False(File.Exists(target));
            Assert.Empty(Directory.GetFiles(directory, ".excelmcp.*.tmp.xlsx"));
            Assert.Equal(original, await File.ReadAllBytesAsync(baseline));

            var accepted = await SaveCheckedCopyAsync(service, session, target, new { baselinePath = baseline, allowedParts = baselineParts });
            using var acceptedDocument = JsonDocument.Parse(accepted.Result!);
            var outcome = acceptedDocument.RootElement;
            Assert.True(outcome.GetProperty("success").GetBoolean(), accepted.Result);
            Assert.True(outcome.GetProperty("filePublished").GetBoolean());
            Assert.Equal(target, outcome.GetProperty("filePath").GetString());
            var acceptedComparison = outcome.GetProperty("packageComparison");
            Assert.True(acceptedComparison.GetProperty("matchesAllowlist").GetBoolean());
            var acceptedWorksheet = CheckedWorksheetChange(acceptedComparison);
            Assert.True(acceptedWorksheet.GetProperty("allowed").GetBoolean());
            Assert.Equal(refusedWorksheet.GetProperty("beforeSha256").GetString(), acceptedWorksheet.GetProperty("beforeSha256").GetString());
            Assert.Equal(original, await File.ReadAllBytesAsync(baseline));
            Assert.Equal("changed marker", await ReadMarkerAsync(service, session));
            Assert.Equal(source, service.SessionManager.GetSession(session)!.WorkbookPath);
            string reopened = await OpenSessionAsync(service, target);
            sessions[reopened] = 0;
            Assert.Equal("changed marker", await ReadMarkerAsync(service, reopened));
        });
    }

    [Fact]
    public async Task SaveCopyAs_CheckedRefusesDirectAndHardLinkAliasesOfBaseline()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            string source = Path.Combine(directory, "owned.xlsx");
            string baseline = Path.Combine(directory, "baseline.xlsx");
            string alias = Path.Combine(directory, "alias.xlsx");
            string session = await CreateSessionAsync(service, source);
            sessions[session] = 0;
            RequireSuccess(await SaveCheckedCopyAsync(service, session, baseline));
            byte[] original = await File.ReadAllBytesAsync(baseline);
            if (!PhysicalCreateHardLink(alias, baseline, IntPtr.Zero))
                throw new Win32Exception(Marshal.GetLastWin32Error());

            string[] allowedParts = [];
            foreach (string destination in new[] { baseline, alias })
            {
                var refused = await SaveCheckedCopyAsync(service, session, destination, new { baselinePath = baseline, allowedParts }, overwrite: true);
                Assert.False(refused.Success);
                Assert.Contains("baseline", refused.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("destination", refused.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            }
            Assert.Equal(original, await File.ReadAllBytesAsync(baseline));
            Assert.Equal(original, await File.ReadAllBytesAsync(alias));
            Assert.Empty(Directory.GetFiles(directory, ".excelmcp.*.tmp.xlsx"));
        });
    }

    [Fact]
    public async Task SaveCopyAs_CheckedConcurrentDestinationRemainsUnchanged()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            string source = Path.Combine(directory, "owned.xlsx");
            string baseline = Path.Combine(directory, "baseline.xlsx");
            string target = Path.Combine(directory, "competing.xlsx");
            string session = await CreateSessionAsync(service, source);
            sessions[session] = 0;
            RequireSuccess(await SaveCheckedCopyAsync(service, session, baseline));
            byte[] original = await File.ReadAllBytesAsync(baseline);
            string[] allowedParts;
            using (var archive = ZipFile.OpenRead(baseline))
                allowedParts = archive.Entries.Select(entry => entry.FullName).ToArray();

            const string competing = "a different publisher owns the destination";
            ServiceResponse response;
            try
            {
                WorkbookCommands.BeforeOutputCommitHookForTests = _ => File.WriteAllText(target, competing);
                response = await SaveCheckedCopyAsync(service, session, target, new { baselinePath = baseline, allowedParts });
            }
            finally
            {
                WorkbookCommands.BeforeOutputCommitHookForTests = null;
            }

            Assert.False(response.Success);
            Assert.Contains("already exists", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(competing, await File.ReadAllTextAsync(target));
            Assert.Equal(original, await File.ReadAllBytesAsync(baseline));
            Assert.Empty(Directory.GetFiles(directory, ".excelmcp.*.tmp.xlsx"));
        });
    }

    [Fact]
    public async Task SaveCopyAs_CheckedPinsCandidateAgainstWritesAndRenamesBeforeReplacingDestination()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            string source = Path.Combine(directory, "owned.xlsx");
            string baseline = Path.Combine(directory, "baseline.xlsx");
            string target = Path.Combine(directory, "existing.xlsx");
            string moved = Path.Combine(directory, "moved.xlsx");
            string session = await CreateSessionAsync(service, source);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "original marker");
            RequireSuccess(await SaveCheckedCopyAsync(service, session, baseline));
            string[] allowedParts;
            using (var archive = ZipFile.OpenRead(baseline))
                allowedParts = archive.Entries.Select(entry => entry.FullName).ToArray();
            await WriteMarkerAsync(service, session, "published from pinned file");
            await File.WriteAllTextAsync(target, "old destination");

            var tested = false;
            ServiceResponse response;
            try
            {
                WorkbookCommands.BeforeOutputCommitHookForTests = stagedPath =>
                {
                    Assert.ThrowsAny<IOException>(() => File.WriteAllText(stagedPath, "unverified replacement"));
                    Assert.ThrowsAny<IOException>(() => File.Move(stagedPath, moved));
                    tested = true;
                };
                response = await SaveCheckedCopyAsync(service, session, target, new { baselinePath = baseline, allowedParts }, overwrite: true);
            }
            finally
            {
                WorkbookCommands.BeforeOutputCommitHookForTests = null;
            }

            Assert.True(tested);
            using var result = JsonDocument.Parse(response.Result!);
            Assert.True(result.RootElement.GetProperty("success").GetBoolean(), response.Result);
            Assert.True(result.RootElement.GetProperty("filePublished").GetBoolean());
            Assert.False(File.Exists(moved));
            Assert.Empty(Directory.GetFiles(directory, ".excelmcp.*.tmp.xlsx"));
            string reopened = await OpenSessionAsync(service, target);
            sessions[reopened] = 0;
            Assert.Equal("published from pinned file", await ReadMarkerAsync(service, reopened));
        });
    }

    [Fact]
    public async Task SaveCopyAs_CheckedReportsRetainedStageWhenAnotherReaderPreventsPinning()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            string session = await CreateSessionAsync(service, Path.Combine(directory, "owned.xlsx"));
            sessions[session] = 0;
            string baseline = Path.Combine(directory, "baseline.xlsx");
            string target = Path.Combine(directory, "checked.xlsx");
            await WriteMarkerAsync(service, session, "source marker");
            RequireSuccess(await SaveCheckedCopyAsync(service, session, baseline));
            FileStream? reader = null;
            string? stage = null;
            try
            {
                WorkbookCommands.BeforeCheckedStageOpenForTests = path =>
                {
                    stage = path;
                    reader = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                };
                var response = await SaveCheckedCopyAsync(service, session, target,
                    new { baselinePath = baseline, allowedParts = Array.Empty<string>() });
                Assert.False(response.Success);
                Assert.Contains("destination was not changed", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("unpublished stage", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                Assert.NotNull(stage);
                Assert.Contains(stage, response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                Assert.True(File.Exists(stage));
                Assert.False(File.Exists(target));
            }
            finally
            {
                WorkbookCommands.BeforeCheckedStageOpenForTests = null;
                reader?.Dispose();
            }
        });
    }

    [Fact]
    public async Task SaveCopyAs_CheckedRejectsInvalidPolicyBeforeNativeStaging()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            string source = Path.Combine(directory, "owned.xlsx");
            string baseline = Path.Combine(directory, "baseline.xlsx");
            string target = Path.Combine(directory, "checked.xlsx");
            string session = await CreateSessionAsync(service, source);
            sessions[session] = 0;
            RequireSuccess(await SaveCheckedCopyAsync(service, session, baseline));
            var staged = false;
            try
            {
                WorkbookCommands.BeforeOutputCommitHookForTests = _ => staged = true;
                var pipe = await SaveCheckedCopyAsync(service, session, target,
                    new { baselinePath = @"\\.\pipe\fake.xlsx", allowedParts = Array.Empty<string>() });
                Assert.False(pipe.Success);
                Assert.Equal("InvalidInput", pipe.ErrorCategory);
                Assert.Contains("Device paths", pipe.ErrorMessage, StringComparison.OrdinalIgnoreCase);

                string[] invalidParts = ["../xl/workbook.xml"];
                var wrongPart = await SaveCheckedCopyAsync(service, session, target,
                    new { baselinePath = baseline, allowedParts = invalidParts });
                Assert.False(wrongPart.Success);
                Assert.Equal("InvalidInput", wrongPart.ErrorCategory);
                Assert.Contains("part names", wrongPart.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            }
            finally
            {
                WorkbookCommands.BeforeOutputCommitHookForTests = null;
            }

            Assert.False(staged);
            Assert.False(File.Exists(target));
            Assert.Empty(Directory.GetFiles(directory, ".excelmcp.*.tmp.xlsx"));
        });
    }

    [Fact]
    public async Task SaveCopyAs_CheckedVbaOnlyPolicyRefusesAdditionalWorksheetChange()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            string source = Path.Combine(directory, "owned.xlsm");
            string baseline = Path.Combine(directory, "baseline.xlsm");
            string target = Path.Combine(directory, "vba-only.xlsm");
            string session = await CreateSessionAsync(service, source);
            sessions[session] = 0;
            const string moduleName = "CheckedVbaSource";
            const string originalCode = "Public Function Marker() As Long\nMarker = 41\nEnd Function";
            const string replacementCode = "Public Function Marker() As Long\nMarker = 42\nEnd Function";
            var imported = await service.ProcessAsync(new ServiceRequest
            {
                Command = "vba.import",
                SessionId = session,
                Args = JsonSerializer.Serialize(new { moduleName, vbaCode = originalCode }, ServiceProtocol.JsonOptions)
            });
            RequireSuccess(imported);
            RequireSuccess(await SaveCheckedCopyAsync(service, session, baseline));
            byte[] original = await File.ReadAllBytesAsync(baseline);

            var view = await service.ProcessAsync(new ServiceRequest
            {
                Command = "vba.view",
                SessionId = session,
                Args = JsonSerializer.Serialize(new { moduleName }, ServiceProtocol.JsonOptions)
            });
            RequireSuccess(view);
            using var viewed = JsonDocument.Parse(view.Result!);
            string expectedSourceHash = Assert.IsType<string>(viewed.RootElement.GetProperty("sourceHash").GetString());
            var replaced = await service.ProcessAsync(new ServiceRequest
            {
                Command = "vba.replace-module",
                SessionId = session,
                Args = JsonSerializer.Serialize(new { moduleName, expectedSourceHash, vbaCode = replacementCode }, ServiceProtocol.JsonOptions)
            });
            RequireSuccess(replaced);
            using var replacement = JsonDocument.Parse(replaced.Result!);
            Assert.NotEqual(expectedSourceHash, replacement.RootElement.GetProperty("sourceHash").GetString());
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "range.set-values",
                SessionId = session,
                Args = """{"sheetName":"Sheet1","rangeAddress":"A1:B1","values":[["worksheet change",42]]}"""
            }));

            const string vbaPart = "xl/vbaProject.bin";
            string[] vbaOnly = [vbaPart];
            var refused = await SaveCheckedCopyAsync(service, session, target, new { baselinePath = baseline, allowedParts = vbaOnly });
            using var refusedDocument = JsonDocument.Parse(refused.Result!);
            var refusal = refusedDocument.RootElement;
            Assert.False(refusal.GetProperty("success").GetBoolean());
            Assert.Equal("Conflict", refusal.GetProperty("errorCategory").GetString());
            Assert.False(refusal.GetProperty("filePublished").GetBoolean());
            var comparison = refusal.GetProperty("packageComparison");
            Assert.False(comparison.GetProperty("matchesAllowlist").GetBoolean());
            var changedBinary = Assert.Single(comparison.GetProperty("changes").EnumerateArray(),
                part => part.GetProperty("partName").GetString() == vbaPart);
            Assert.Equal("changed", changedBinary.GetProperty("changeKind").GetString());
            Assert.True(changedBinary.GetProperty("allowed").GetBoolean());
            Assert.Matches("^[0-9a-f]{64}$", changedBinary.GetProperty("beforeSha256").GetString()!);
            Assert.Matches("^[0-9a-f]{64}$", changedBinary.GetProperty("afterSha256").GetString()!);
            Assert.NotEqual(changedBinary.GetProperty("beforeSha256").GetString(), changedBinary.GetProperty("afterSha256").GetString());
            Assert.False(CheckedWorksheetChange(comparison).GetProperty("allowed").GetBoolean());
            Assert.False(File.Exists(target));
            Assert.Empty(Directory.GetFiles(directory, ".excelmcp.*.tmp.xlsm"));
            Assert.Equal(original, await File.ReadAllBytesAsync(baseline));
            Assert.Equal("worksheet change", await ReadMarkerAsync(service, session));
        });
    }

    private static JsonElement CheckedWorksheetChange(JsonElement comparison)
    {
        var worksheet = Assert.Single(comparison.GetProperty("changes").EnumerateArray(),
            part => part.GetProperty("partName").GetString() == "xl/worksheets/sheet1.xml");
        Assert.Equal("changed", worksheet.GetProperty("changeKind").GetString());
        Assert.Matches("^[0-9a-f]{64}$", worksheet.GetProperty("beforeSha256").GetString()!);
        Assert.Matches("^[0-9a-f]{64}$", worksheet.GetProperty("afterSha256").GetString()!);
        Assert.NotEqual(worksheet.GetProperty("beforeSha256").GetString(), worksheet.GetProperty("afterSha256").GetString());
        return worksheet;
    }

    private static Task<ServiceResponse> SaveCheckedCopyAsync(
        ExcelMcpService service, string session, string targetPath, object? packageCheck = null, bool overwrite = false) =>
        service.ProcessAsync(new ServiceRequest
        {
            Command = "workbook.save-copy-as",
            SessionId = session,
            Args = JsonSerializer.Serialize(new { targetPath, overwrite, packageCheck }, ServiceProtocol.JsonOptions)
        });
}
