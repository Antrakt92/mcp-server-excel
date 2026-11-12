using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Fact]
    public async Task SaveAs_DestinationCreatedAfterAdmission_RejectsAndPreservesLiveWorkbook()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var sourcePath = Path.Join(directory, "save-as-source.xlsx");
            var targetPath = Path.Join(directory, "save-as-concurrent.xlsx");
            var sessionId = await CreateSessionAsync(service, sourcePath);
            sessions[sessionId] = 0;
            const string marker = "Pending source edit";
            const string destinationContents = "Created by another writer after admission";
            await WriteMarkerAsync(service, sessionId, marker);
            var batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(sessionId));
            var originalAlerts = batch.Execute((context, _) => context.App.DisplayAlerts);
            batch.Execute((context, _) => context.App.DisplayAlerts = true);
            try
            {
                Assert.False(File.Exists(targetPath));
                WorkbookCommands.BeforeNativeSaveAsHookForTests = path =>
                {
                    Assert.Equal(targetPath, path, ignoreCase: true);
                    using var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write);
                    using var writer = new StreamWriter(output);
                    writer.Write(destinationContents);
                };
                var response = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "workbook.save-as",
                    SessionId = sessionId,
                    Args = JsonSerializer.Serialize(new { targetPath, overwrite = false }, ServiceProtocol.JsonOptions)
                });

                Assert.False(response.Success);
                Assert.Contains("already exists", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                Assert.Equal(destinationContents, await File.ReadAllTextAsync(targetPath));
                Assert.Same(batch, service.SessionManager.GetSession(sessionId));
                Assert.Equal(sourcePath, batch.Execute((context, _) => context.Book.FullName), ignoreCase: true);
                Assert.True(service.SessionManager.TryGetFilePath(sessionId, out var trackedPath));
                Assert.Equal(sourcePath, trackedPath, ignoreCase: true);
                Assert.True(batch.Execute((context, _) => context.App.DisplayAlerts));
                Assert.False(batch.Execute((context, _) => context.Book.Saved));
                Assert.Equal(marker, await ReadMarkerAsync(service, sessionId));
            }
            finally
            {
                WorkbookCommands.BeforeNativeSaveAsHookForTests = null;
                batch.Execute((context, _) => context.App.DisplayAlerts = originalAlerts);
            }
        });
    }
}
