using System.Security.Cryptography;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Fact]
    public async Task AuditSession_ProtectionSurvivesInteractiveSaveAsReconciliation()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var original = Path.Join(directory, "audit-original.xlsx");
            var savedAs = Path.Join(directory, "audit-user-save-as.xlsx");
            var initial = await CreateSessionAsync(service, original);
            sessions[initial] = 0;
            await WriteMarkerAsync(service, initial, "baseline");
            await CloseSessionAsync(service, initial, save: true);
            sessions.TryRemove(initial, out _);
            var id = await OpenAuditAsync(service, original);
            sessions[id] = 0;
            // Simulate the user acting in a visible Excel window, outside the agent API.
            service.SessionManager.GetSession(id)!.Execute((ctx, _) =>
                ctx.Book.SaveAs(savedAs, Excel.XlFileFormat.xlOpenXMLWorkbook));
            RequireSuccess(await service.ProcessAsync(new ServiceRequest { Command = "session.discover" }));
            Assert.True(service.SessionManager.IsAuditSession(id));
            Assert.Equal(savedAs, service.SessionManager.GetSession(id)!.WorkbookPath);
            var save = await service.ProcessAsync(new ServiceRequest { Command = "session.save", SessionId = id });
            Assert.False(save.Success);
            Assert.Contains("Audit session", save.ErrorMessage!, StringComparison.Ordinal);
            Assert.Equal("baseline", await ReadMarkerAsync(service, id));
        });
    }

    [Fact]
    public async Task AuditSession_ShutdownHandsOverInteractiveUnsavedEditsWithoutSaving()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "audit-handover.xlsx");
            var initial = await CreateSessionAsync(service, path);
            sessions[initial] = 0;
            await WriteMarkerAsync(service, initial, "saved baseline");
            await CloseSessionAsync(service, initial, save: true);
            sessions.TryRemove(initial, out _);
            var originalHash = SHA256.HashData(await File.ReadAllBytesAsync(path));
            var id = await OpenAuditAsync(service, path);
            sessions[id] = 0;
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(id));
            batch.Execute((ctx, _) =>
            {
                Excel.Sheets? sheets = null;
                Excel.Worksheet? sheet = null;
                Excel.Range? cell = null;
                try
                {
                    sheets = ctx.Book.Worksheets;
                    sheet = (Excel.Worksheet)sheets[1];
                    cell = sheet.Range["A1"];
                    cell.Value2 = "interactive unsaved edit";
                    Assert.False(ctx.Book.Saved);
                }
                finally
                {
                    ComUtilities.Release(ref cell);
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                }
            });
            var handedOver = false;
            try
            {
                RequireSuccess(await service.ProcessAsync(new ServiceRequest { Command = "service.shutdown" }));
                handedOver = true;
                sessions.TryRemove(id, out _);
                Assert.True(batch.IsExcelProcessAlive());
                using var replacement = new ExcelMcpService();
                var recovered = await AttachHandoverAsync(replacement, path);
                try { Assert.Equal("interactive unsaved edit", await ReadMarkerAsync(replacement, recovered)); }
                finally { await CloseSessionAsync(replacement, recovered, save: false); }
                Assert.Equal(originalHash, SHA256.HashData(await File.ReadAllBytesAsync(path)));
            }
            finally
            {
                if (handedOver && batch.IsExcelProcessAlive())
                {
                    using var cleanup = new ExcelMcpService();
                    var recovered = await AttachHandoverAsync(cleanup, path);
                    await CloseSessionAsync(cleanup, recovered, save: false);
                }
            }
        });
    }

    private static async Task<string> OpenAuditAsync(ExcelMcpService service, string path)
    {
        var opened = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = path, readOnly = true })
        });
        RequireSuccess(opened);
        using var document = JsonDocument.Parse(opened.Result!);
        Assert.True(document.RootElement.GetProperty("auditOnly").GetBoolean());
        return document.RootElement.GetProperty("sessionId").GetString()!;
    }

    [Fact]
    public async Task AuditSession_ReadsWithoutMacrosAndRefusesEveryMutationBeforeExcel()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "audit.xlsm");
            var initial = await CreateSessionAsync(service, path);
            sessions[initial] = 0;
            await WriteMarkerAsync(service, initial, "not-run");
            service.SessionManager.GetSession(initial)!.Execute((ctx, _) => InstallSecurityProbeMacros(ctx));
            await CloseSessionAsync(service, initial, save: true);
            sessions.TryRemove(initial, out _);
            var originalHash = SHA256.HashData(await File.ReadAllBytesAsync(path));

            var opened = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = path, readOnly = true })
            });
            RequireSuccess(opened);
            using var opening = JsonDocument.Parse(opened.Result!);
            var id = opening.RootElement.GetProperty("sessionId").GetString()!;
            sessions[id] = 0;
            Assert.True(opening.RootElement.GetProperty("auditOnly").GetBoolean());
            Assert.True(service.SessionManager.IsAuditSession(id));
            Assert.True(service.SessionManager.GetSession(id)!.Execute((ctx, _) => ctx.Book.ReadOnly));
            Assert.Equal("not-run", await ReadMarkerAsync(service, id));

            foreach (var (command, args) in new[]
            {
                ("range.set-values", """{"sheetName":"Sheet1","rangeAddress":"A1","values":[["changed"]],"overwritePolicy":"allow"}"""),
                ("session.save", "{}"),
                ("session.close", """{"save":true}"""),
                ("session.enable-macros", """{"approved":true,"allowReopen":true}"""),
                ("vba.run", """{"procedureName":"ThisWorkbook.SecurityMarker"}""")
            })
            {
                var response = await service.ProcessAsync(new ServiceRequest { Command = command, SessionId = id, Args = args });
                Assert.False(response.Success);
                Assert.Contains("Audit session", response.ErrorMessage!, StringComparison.Ordinal);
                Assert.Equal("not-run", await ReadMarkerAsync(service, id));
            }

            var listed = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
            RequireSuccess(listed);
            using var list = JsonDocument.Parse(listed.Result!);
            Assert.True(list.RootElement.GetProperty("sessions").EnumerateArray()
                .Single(entry => entry.GetProperty("sessionId").GetString() == id)
                .GetProperty("auditOnly").GetBoolean());
            await CloseSessionAsync(service, id, save: false);
            sessions.TryRemove(id, out _);
            Assert.Equal(originalHash, SHA256.HashData(await File.ReadAllBytesAsync(path)));
        });
    }
}
