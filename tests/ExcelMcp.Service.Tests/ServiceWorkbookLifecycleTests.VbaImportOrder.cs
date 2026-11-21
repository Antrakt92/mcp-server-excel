using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task VbaInvoke_ImportBeforeOrAfterMacroApproval_PersistsOneInvocation(bool importBeforeApproval)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Combine(directory, "invoke-import-order.xlsm");
            var sessionId = await CreateSessionAsync(service, path);
            sessions[sessionId] = 0;
            DisableFixtureAutoRecover();
            // Macro approval requires an existing VBA project in both cases.
            await SendAsync("vba.import", new
            {
                moduleName = "FixtureBaseline",
                vbaCode = "Public Sub BaselineMarker()\nEnd Sub"
            });
            if (importBeforeApproval)
                await ImportAsync();

            await CloseSessionAsync(service, sessionId, save: true);
            sessions.TryRemove(sessionId, out _);
            sessionId = await OpenSessionAsync(service, path);
            sessions[sessionId] = 0;
            await SendAsync("session.enable-macros", new { approved = true, allowReopen = true });
            DisableFixtureAutoRecover();

            if (!importBeforeApproval)
                await ImportAsync();

            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSessionForDiagnostics(sessionId));
            var identity = Assert.IsType<ExcelProcessIdentity>(batch.ProcessIdentity);
            var response = await SendAsync("vba.invoke", new { procedureName = "ImportOrderProbe.WriteMarker" });
            using var returned = JsonDocument.Parse(response.Result!);
            Assert.Equal("string", returned.RootElement.GetProperty("returnKind").GetString());
            Assert.Equal("invoke-marker", returned.RootElement.GetProperty("value").GetString());
            await VerifyMarkerAsync();
            Assert.Equal(identity, batch.ProcessIdentity);

            await CloseSessionAsync(service, sessionId, save: true);
            sessions.TryRemove(sessionId, out _);
            sessionId = await OpenSessionAsync(service, path);
            sessions[sessionId] = 0;
            await VerifyMarkerAsync();

            async Task VerifyMarkerAsync()
            {
                var values = await SendAsync("range.get-values", new { sheetName = "Sheet1", rangeAddress = "A1:B1" });
                using var json = JsonDocument.Parse(values.Result!);
                var row = json.RootElement.GetProperty("values")[0];
                Assert.Equal("invoke-marker", row[0].GetString());
                Assert.Equal(1d, row[1].GetDouble());
            }

            void DisableFixtureAutoRecover()
            {
                var owner = Assert.IsType<ExcelBatch>(service.SessionManager.GetSessionForDiagnostics(sessionId));
                owner.Execute((ctx, _) =>
                {
                    ctx.Book.EnableAutoRecover = false;
                    return 0;
                });
            }

            async Task ImportAsync() => await SendAsync("vba.import", new
            {
                moduleName = "ImportOrderProbe",
                vbaCode = """
                    Public Function WriteMarker() As String
                        With ThisWorkbook.Worksheets(1)
                            .Range("A1").Value2 = "invoke-marker"
                            .Range("B1").Value2 = .Range("B1").Value2 + 1
                        End With
                        WriteMarker = "invoke-marker"
                    End Function
                    """
            });

            async Task<ServiceResponse> SendAsync(string command, object args)
            {
                var result = await service.ProcessAsync(new ServiceRequest
                {
                    Command = command,
                    SessionId = sessionId,
                    Args = JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions)
                });
                Assert.True(result.Success, $"{command}: {result.ErrorCategory} {result.HResult} {result.ErrorMessage}");
                return result;
            }
        });
    }
}
