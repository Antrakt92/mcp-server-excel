using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public async Task Save_WhenWorkbookRemainsDirty_SeparatesExcelWriteReportAndSessionState(bool attached, bool veto)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            const string baselineMarker = "Saved baseline";
            const string requestedMarker = "Changes requested for save";
            var path = Path.Join(directory, "save-evidence.xlsx");
            var owner = await CreateSessionAsync(service,
                attached ? Path.Join(directory, "attachment-host.xlsx") : path);
            sessions[owner] = 0;
            var session = owner;
            if (attached)
            {
                var ownerBatch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(owner));
                var identity = Assert.IsType<ExcelProcessIdentity>(ownerBatch.ProcessIdentity);
                var name = ownerBatch.Execute((context, _) => AddReviewWorkbook(context.App));
                session = await AttachReviewWorkbookAsync(service, await DiscoverReviewWorkbookAsync(service, name, identity));
                sessions[session] = 0;
                RequireSuccess(await service.ProcessAsync(new ServiceRequest
                {
                    Command = "workbook.save-as",
                    SessionId = session,
                    Args = JsonSerializer.Serialize(new { targetPath = path }, ServiceProtocol.JsonOptions)
                }));
            }

            var batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            batch.Execute((context, _) => { context.Book.EnableAutoRecover = false; });
            var saveRequest = new ServiceRequest { Command = "session.save", SessionId = session };
            await WriteMarkerAsync(service, session, baselineMarker);
            RequireSuccess(await service.ProcessAsync(saveRequest));
            var savedHash = await HashPhysicalWorkbookAsync(path);
            await WriteMarkerAsync(service, session, requestedMarker);

            var callbackCount = 0;
            Excel.AppEvents_WorkbookAfterSaveEventHandler afterSave = (workbook, success) =>
            {
                if (!string.Equals(workbook.FullName, path, StringComparison.OrdinalIgnoreCase) || !success)
                    return;
                callbackCount++;
                WriteLateCloseSaveMarker(workbook);
            };
            Excel.AppEvents_WorkbookBeforeSaveEventHandler beforeSave =
                (Excel.Workbook workbook, bool _, ref bool cancel) =>
                {
                    if (!string.Equals(workbook.FullName, path, StringComparison.OrdinalIgnoreCase))
                        return;
                    callbackCount++;
                    cancel = true;
                };
            batch.Execute((context, _) =>
            {
                if (veto)
                    context.App.WorkbookBeforeSave += beforeSave;
                else
                    context.App.WorkbookAfterSave += afterSave;
            });
            try
            {
                var response = await service.ProcessAsync(saveRequest);
                Assert.False(response.Success);
                Assert.Equal(1, callbackCount);
                Assert.Same(batch, service.SessionManager.GetSession(session));
                Assert.False(batch.Execute((context, _) => context.Book.Saved));
                Assert.Equal(veto ? requestedMarker : LateCloseSaveMarker, await ReadMarkerAsync(service, session));
                var afterSaveHash = await HashPhysicalWorkbookAsync(path);
                var changedOnDisk = !savedHash.AsSpan().SequenceEqual(afterSaveHash);
                Assert.Equal(!veto, changedOnDisk);
                using var state = JsonDocument.Parse(Assert.IsType<string>(response.PartialState));
                var save = state.RootElement.GetProperty("save");
                Assert.True(save.GetProperty("hasUnsavedChanges").GetBoolean());
                Assert.False(save.TryGetProperty("observationError", out _));
                var reported = save.GetProperty("excelReportedSaveSucceeded");
                if (veto)
                    Assert.True(reported.ValueKind is JsonValueKind.Null or JsonValueKind.False);
                else
                    Assert.True(reported.GetBoolean());
            }
            finally
            {
                batch.Execute((context, _) =>
                {
                    if (veto)
                        context.App.WorkbookBeforeSave -= beforeSave;
                    else
                        context.App.WorkbookAfterSave -= afterSave;
                });
            }

            await CloseSessionAsync(service, session, save: false);
            sessions.TryRemove(session, out _);
            var reopened = await OpenSessionAsync(service, path);
            sessions[reopened] = 0;
            Assert.Equal(veto ? baselineMarker : requestedMarker, await ReadMarkerAsync(service, reopened));
        });
    }
}
