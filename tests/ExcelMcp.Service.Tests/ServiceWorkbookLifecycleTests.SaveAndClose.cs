using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    private const string FirstCloseSaveMarker = "Saved by the first save";
    private const string LateCloseSaveMarker = "Edited after the first save";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CloseWithSave_PersistsEditAfterInitialSave(bool attached)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "save-and-close.xlsx");
            var owner = await CreateSessionAsync(service,
                attached ? Path.Join(directory, "attachment-host.xlsx") : path);
            sessions[owner] = 0;
            var session = owner;
            if (attached)
            {
                var ownerBatch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(owner));
                var name = ownerBatch.Execute((context, _) => AddReviewWorkbook(context.App));
                session = await AttachReviewWorkbookAsync(service, await DiscoverReviewWorkbookAsync(service, name));
                sessions[session] = 0;
                RequireSuccess(await service.ProcessAsync(new ServiceRequest
                {
                    Command = "workbook.save-as",
                    SessionId = session,
                    Args = JsonSerializer.Serialize(new { targetPath = path })
                }));
            }

            await WriteMarkerAsync(service, session, FirstCloseSaveMarker);
            var batch = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(session));
            var edited = false;
            service.SessionManager.BeforeCloseSessionHookForTests = () => batch.Execute((context, _) =>
            {
                Assert.True(context.Book.Saved);
                WriteLateCloseSaveMarker(context.Book);
                Assert.False(context.Book.Saved);
                edited = true;
            });
            try
            {
                await CloseSessionAsync(service, session, save: true);
                sessions.TryRemove(session, out _);
                Assert.True(edited);
            }
            finally
            {
                service.SessionManager.BeforeCloseSessionHookForTests = null;
            }

            var reopened = await OpenSessionAsync(service, path);
            sessions[reopened] = 0;
            Assert.Equal(LateCloseSaveMarker, await ReadMarkerAsync(service, reopened));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CloseWithSave_PreservesLateEditWhenFinalSaveOrCloseIsRefused(bool multipleWindows)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "refused-save-and-close.xlsx");
            var session = await CreateSessionAsync(service, path);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, FirstCloseSaveMarker);
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
            var saveAttempts = 0;
            var addedWindow = false;
            var subscribed = false;
            Excel.AppEvents_WorkbookBeforeSaveEventHandler cancelSave =
                (Excel.Workbook _, bool _, ref bool cancel) =>
                {
                    saveAttempts++;
                    cancel = true;
                };
            service.SessionManager.BeforeCloseSessionHookForTests = () => batch.Execute((context, _) =>
            {
                Assert.True(context.Book.Saved);
                WriteLateCloseSaveMarker(context.Book);
                Assert.False(context.Book.Saved);
                if (multipleWindows)
                {
                    Excel.Window? window = null;
                    try
                    {
                        window = context.Book.NewWindow();
                        addedWindow = true;
                    }
                    finally { ComUtilities.Release(ref window); }
                }
                else
                {
                    context.App.WorkbookBeforeSave += cancelSave;
                    subscribed = true;
                }
            });
            try
            {
                var response = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.close",
                    SessionId = session,
                    Args = """{"save":true}"""
                });
                Assert.False(response.Success);
                Assert.Same(batch, service.SessionManager.GetSession(session));
                Assert.False(batch.Execute((context, _) => context.Book.Saved));
                Assert.Equal(LateCloseSaveMarker, await ReadMarkerAsync(service, session));
                if (multipleWindows)
                {
                    Assert.Equal("Busy", response.ErrorCategory);
                    Assert.Contains("multiple open windows", response.ErrorMessage, StringComparison.Ordinal);
                    Assert.True(addedWindow);
                }
                else
                {
                    Assert.Equal(1, saveAttempts);
                    Assert.Contains("was not saved", response.ErrorMessage, StringComparison.Ordinal);
                }
            }
            finally
            {
                service.SessionManager.BeforeCloseSessionHookForTests = null;
                if (ReferenceEquals(batch, service.SessionManager.GetSessionForDiagnostics(session)))
                    batch.Execute((context, _) =>
                    {
                        if (subscribed)
                            context.App.WorkbookBeforeSave -= cancelSave;
                        if (!addedWindow)
                            return;
                        Excel.Windows? windows = null;
                        Excel.Window? window = null;
                        try
                        {
                            windows = context.Book.Windows;
                            if (windows.Count > 1)
                            {
                                window = windows.Item[2];
                                window.Close(SaveChanges: false);
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref window);
                            ComUtilities.Release(ref windows);
                        }
                    });
            }

            await CloseSessionAsync(service, session, save: false);
            sessions.TryRemove(session, out _);
            var reopened = await OpenSessionAsync(service, path);
            sessions[reopened] = 0;
            Assert.Equal(FirstCloseSaveMarker, await ReadMarkerAsync(service, reopened));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CloseWithSave_HandlesEditsFromNativeBeforeCloseEvent(bool cancelImplicitSave)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "before-close-save.xlsx");
            var session = await CreateSessionAsync(service, path);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, FirstCloseSaveMarker);
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(session));
            var beforeCloseRan = false;
            var implicitSaveAttempts = 0;
            Excel.AppEvents_WorkbookBeforeCloseEventHandler? onClose = null;
            Excel.AppEvents_WorkbookBeforeSaveEventHandler vetoSave =
                (Excel.Workbook _, bool _, ref bool cancel) =>
                {
                    implicitSaveAttempts++;
                    cancel = true;
                };
            batch.Execute((context, _) =>
            {
                onClose = (Excel.Workbook workbook, ref bool cancel) =>
                {
                    if (!string.Equals(workbook.FullName, path, StringComparison.OrdinalIgnoreCase))
                        return;

                    beforeCloseRan = true;
                    context.App.WorkbookBeforeClose -= onClose;
                    WriteLateCloseSaveMarker(workbook);
                    if (cancelImplicitSave)
                        context.App.WorkbookBeforeSave += vetoSave;
                };
                context.App.WorkbookBeforeClose += onClose;
            });

            try
            {
                var response = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.close",
                    SessionId = session,
                    Args = """{"save":true}"""
                });
                if (response.Success)
                    sessions.TryRemove(session, out _);
                Assert.True(beforeCloseRan);
                Assert.Equal(!cancelImplicitSave, response.Success);
                if (cancelImplicitSave)
                {
                    Assert.Equal(1, implicitSaveAttempts);
                    Assert.Same(batch, service.SessionManager.GetSession(session));
                    Assert.False(batch.Execute((context, _) => context.Book.Saved));
                    Assert.Equal(LateCloseSaveMarker, await ReadMarkerAsync(service, session));
                }
            }
            finally
            {
                if (ReferenceEquals(batch, service.SessionManager.GetSessionForDiagnostics(session)))
                    batch.Execute((context, _) =>
                    {
                        if (!beforeCloseRan)
                            context.App.WorkbookBeforeClose -= onClose;
                        if (cancelImplicitSave && beforeCloseRan)
                            context.App.WorkbookBeforeSave -= vetoSave;
                    });
            }

            if (cancelImplicitSave)
            {
                await CloseSessionAsync(service, session, save: false);
                sessions.TryRemove(session, out _);
            }
            var reopened = await OpenSessionAsync(service, path);
            sessions[reopened] = 0;
            Assert.Equal(cancelImplicitSave ? FirstCloseSaveMarker : LateCloseSaveMarker,
                await ReadMarkerAsync(service, reopened));
        });
    }

    private static void WriteLateCloseSaveMarker(Excel.Workbook workbook)
    {
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;
        Excel.Range? cell = null;
        try
        {
            sheets = workbook.Worksheets;
            sheet = (Excel.Worksheet)sheets[1];
            cell = sheet.Range["A1"];
            cell.Value2 = LateCloseSaveMarker;
        }
        finally
        {
            ComUtilities.Release(ref cell);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref sheets);
        }
    }
}
