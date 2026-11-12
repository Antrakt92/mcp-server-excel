using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Fact]
    public async Task Create_DestinationAppearsAfterAdmission_PreservesTheOtherFile()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Combine(directory, "competing.xlsx");
            var otherFile = "another publisher owns this file"u8.ToArray();
            service.SessionManager.BeforeNewFileSessionPublishHookForTests = () => File.WriteAllBytes(path, otherFile);
            ServiceResponse response;
            try
            {
                response = await service.ProcessAsync(CreateRequest(path));
                if (response.SessionId is { } sessionId)
                    sessions[sessionId] = 0;
            }
            finally
            {
                service.SessionManager.BeforeNewFileSessionPublishHookForTests = null;
            }

            Assert.False(response.Success);
            Assert.Equal("Conflict", response.ErrorCategory);
            Assert.Equal(otherFile, await File.ReadAllBytesAsync(path));
            Assert.Null(response.SessionId);
            Assert.Empty(service.SessionManager.ActiveSessionIds);
            Assert.Equal(path, Assert.Single(Directory.GetFileSystemEntries(directory)));
        });
    }

    [Fact]
    public async Task Create_StageSaveFailsAfterAddingWorkbook_ClosesExactNewWorkbook()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Combine(directory, "failed-stage.xlsx");
            ExcelBatch.BeforeNewWorkbookSaveHookForTests = () => throw new IOException("Injected first save failure.");
            ServiceResponse response;
            try
            {
                response = await service.ProcessAsync(CreateRequest(path));
                if (response.SessionId is { } sessionId)
                    sessions[sessionId] = 0;
            }
            finally
            {
                ExcelBatch.BeforeNewWorkbookSaveHookForTests = null;
            }

            Assert.False(response.Success);
            Assert.Contains("Injected first save failure.", response.ErrorMessage, StringComparison.Ordinal);
            Assert.Null(response.SessionId);
            Assert.Empty(service.SessionManager.ActiveSessionIds);
            Assert.Empty(Directory.GetFileSystemEntries(directory));
        });
    }

    [Fact]
    public async Task Create_OpenFailsAfterPublication_RetainsWorkbookAndReportsPublication()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Combine(directory, "published.xlsx");
            ExcelBatch.BeforeWorkbookOpenHook = (opening, _) =>
            {
                if (string.Equals(opening, path, StringComparison.OrdinalIgnoreCase))
                    throw new IOException("Injected open failure after publication.");
            };
            ServiceResponse response;
            try
            {
                response = await service.ProcessAsync(CreateRequest(path));
                if (response.SessionId is { } sessionId)
                    sessions[sessionId] = 0;
            }
            finally
            {
                ExcelBatch.BeforeWorkbookOpenHook = null;
            }

            Assert.False(response.Success);
            Assert.Contains("Injected open failure after publication.", response.ErrorMessage, StringComparison.Ordinal);
            Assert.Null(response.SessionId);
            Assert.Empty(service.SessionManager.ActiveSessionIds);
            using var partialState = JsonDocument.Parse(response.PartialState!);
            Assert.True(partialState.RootElement.GetProperty("filePublished").GetBoolean());
            Assert.False(partialState.RootElement.GetProperty("sessionEstablished").GetBoolean());
            Assert.Equal(path, partialState.RootElement.GetProperty("filePath").GetString());
            Assert.Equal(path, Assert.Single(Directory.GetFileSystemEntries(directory)));

            var opened = await OpenSessionAsync(service, path);
            sessions[opened] = 0;
            Assert.Equal(path, service.SessionManager.GetSession(opened)!.WorkbookPath);
            await CloseSessionAsync(service, opened, save: false);
        });
    }

    private static ServiceRequest CreateRequest(string path) => new()
    {
        Command = "session.create",
        Args = JsonSerializer.Serialize(new { filePath = path })
    };
}
