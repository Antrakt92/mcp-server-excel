using System.Diagnostics;
using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.McpServer.ServiceBridge;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "SessionLifetime")]
[Trait("RequiresExcel", "true")]
public sealed class McpSessionLifetimeTests(ITestOutputHelper output)
{
    [Fact]
    public async Task CancelledStartup_ShutdownLeavesPendingWorkbookAndReportsIt()
    {
        var path = Path.Join(Path.GetTempPath(), $"mcp-pending-startup-{Guid.NewGuid():N}.xlsx");
        using var cancellation = new CancellationTokenSource();
        using var release = new ManualResetEventSlim();
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var backend = new ObservedBackend { ObserveCreation = true };
        var bridge = new ServiceBridge.ServiceBridge(() => backend);
        ExcelBatch.AfterWorkbookOpenHookForTests = (_, workbook) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            try
            {
                sheets = ((Excel.Workbook)workbook).Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                cell = sheet.Range["A1"];
                cell.Value2 = "Retained during cancelled startup";
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
            entered.SetResult();
            Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
        };
        ExcelProcessIdentity? identity = null;
        var failure = await Record.ExceptionAsync(async () =>
        {
            var creation = bridge.SendAsync("session.create", args: new { filePath = path },
                cancellationToken: cancellation.Token);
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(20));
            var pending = Assert.Single(backend.SessionManager.GetActiveSessions());
            var batch = Assert.IsType<ExcelBatch>(backend.SessionManager.GetSession(pending.SessionId));
            identity = Assert.Single(SessionManager.GetTrackedExcelProcesses(),
                candidate => candidate.ProcessId == batch.ExcelProcessId);
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => creation);
            bridge.Dispose();
            Assert.Contains($"Excel process {batch.ExcelProcessId}", backend.ShutdownReport, StringComparison.Ordinal);
            Assert.Same(batch, backend.SessionManager.GetSession(pending.SessionId));
            Assert.True(OwnedProcessGuard.MatchesLiveProcess(identity.Value));
            Assert.Equal(path, Assert.Single(backend.SessionManager.GetActiveSessions()).FilePath);
            ExcelBatch.AfterWorkbookOpenHookForTests = null;
            release.Set();
            var completed = await backend.Finished.Task.WaitAsync(TimeSpan.FromSeconds(20));
            Assert.True(completed.Success, completed.ErrorMessage);
            Assert.Equal(pending.SessionId, await backend.Revealed.Task.WaitAsync(TimeSpan.FromSeconds(10)));
            Assert.True(SpinWait.SpinUntil(
                () => batch.GetRefreshState() == WorkbookRefreshState.Ready,
                TimeSpan.FromSeconds(10)));
            // The stopped service refuses requests; the agent saves the handed-over workbook in Excel itself.
            batch.Execute((context, _) => context.Book.Save());
            backend.SessionManager.Dispose();
            Assert.Equal(0, backend.SessionManager.ActiveSessionCount);
            Assert.True(OwnedProcessGuard.TryConfirmExited(identity.Value));

            using var reader = new ExcelMcpService();
            var opened = await reader.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = path }, ServiceProtocol.JsonOptions)
            });
            Assert.True(opened.Success, opened.ErrorMessage);
            using var session = JsonDocument.Parse(opened.Result!);
            var values = await reader.ProcessAsync(new ServiceRequest
            {
                Command = "range.get-values",
                SessionId = session.RootElement.GetProperty("sessionId").GetString(),
                Args = """{"sheetName":"Sheet1","rangeAddress":"A1"}"""
            });
            Assert.True(values.Success, values.ErrorMessage);
            using var contents = JsonDocument.Parse(values.Result!);
            Assert.Equal("Retained during cancelled startup",
                contents.RootElement.GetProperty("values")[0][0].GetString());
        });
        release.Set();
        ExcelBatch.AfterWorkbookOpenHookForTests = null;
        var cleanup = await Record.ExceptionAsync(async () =>
        {
            await backend.Finished.Task.WaitAsync(TimeSpan.FromSeconds(30));
            foreach (var sessionId in backend.SessionManager.ActiveSessionIds)
                backend.SessionManager.CloseSession(sessionId, save: false);
            bridge.Dispose();
            if (identity is { } owned)
                Assert.True(OwnedProcessGuard.TryConfirmExited(owned));
            File.Delete(path);
        });
        var failures = new[] { failure, cleanup }.OfType<Exception>().ToArray();
        if (failures.Length > 0)
            throw new AggregateException("Cancelled MCP startup or retained teardown failed.", failures);
    }

    [Fact]
    public async Task CancelledCreation_RetainsLateWorkbookWithoutClosingOtherSession()
    {
        var firstPath = Path.Combine(Path.GetTempPath(), $"mcp-survivor-{Guid.NewGuid():N}.xlsx");
        var secondPath = Path.Combine(Path.GetTempPath(), $"mcp-cancelled-{Guid.NewGuid():N}.xlsx");
        using var shutdown = new CancellationTokenSource();
        using var cancellation = new CancellationTokenSource();
        using var release = new ManualResetEventSlim();
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var input = new Pipe();
        var outputPipe = new Pipe();
        var backend = new ObservedBackend();
        var host = await ProgramTransportTestHost.StartAsync(
            input, outputPipe, shutdown.Token, "CreationCancellationClient", () => backend);
        try
        {
            var first = await CallAsync(host.Client, "file", new() { ["action"] = "create", ["file_path"] = firstPath });
            var firstId = first.GetProperty("workbook_session_id").GetString()!;
            await CallAsync(host.Client, "range", new()
            {
                ["action"] = "set-values",
                ["workbook_session_id"] = firstId,
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "A1",
                ["values"] = new object[][] { ["keep"] }
            });
            backend.ObserveCreation = true;
            ExcelBatch.AfterWorkbookOpenHookForTests = (_, _) =>
            {
                entered.SetResult();
                Assert.True(release.Wait(TimeSpan.FromSeconds(30)));
            };
            var request = new JsonRpcRequest
            {
                Id = new RequestId("cancelled-creation"),
                Method = RequestMethods.ToolsCall,
                Params = JsonSerializer.SerializeToNode(new
                {
                    name = "file",
                    arguments = new { action = "create", file_path = secondPath }
                })
            };
            var creation = host.Client.SendRequestAsync(request, cancellation.Token);
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(30));
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => creation);
            // Cancelling a client's local wait does not prove it sent the protocol notification.
            await host.Client.SendMessageAsync(new JsonRpcNotification
            {
                Method = NotificationMethods.CancelledNotification,
                Params = JsonSerializer.SerializeToNode(new { requestId = "cancelled-creation" })
            });
            await host.Client.ListToolsAsync();
            release.Set();
            var completed = await backend.Finished.Task.WaitAsync(TimeSpan.FromSeconds(150));
            Assert.True(completed.Success, completed.ErrorMessage);
            var retainedId = await backend.Revealed.Task.WaitAsync(TimeSpan.FromSeconds(30));
            Assert.NotEqual(firstId, retainedId);

            var sessions = await CallAsync(host.Client, "file_read", new() { ["action"] = "list" });
            var entries = sessions.GetProperty("sessions").EnumerateArray().ToArray();
            Assert.Equal(2, entries.Length);
            Assert.Contains(entries, session => session.GetProperty("workbook_session_id").GetString() == firstId);
            Assert.True(Assert.Single(entries, session =>
                session.GetProperty("workbook_session_id").GetString() == retainedId)
                .GetProperty("isExcelVisible").GetBoolean());
            var values = await CallAsync(host.Client, "range_read", new()
            {
                ["action"] = "get-values",
                ["workbook_session_id"] = firstId,
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "A1"
            });
            Assert.Equal("keep", values.GetProperty("values")[0][0].GetString());
        }
        finally
        {
            release.Set();
            ExcelBatch.AfterWorkbookOpenHookForTests = null;
            foreach (var session in backend.SessionManager.GetActiveSessions()
                .Where(session => session.FilePath == firstPath || session.FilePath == secondPath))
                backend.SessionManager.CloseSession(session.SessionId, save: false);
            await ProgramTransportTestHost.StopAsync(host.Client, input, outputPipe, host.ServerTask, output, shutdown);
            File.Delete(firstPath);
            File.Delete(secondPath);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Shutdown_HandsOverUnsavedChangesUnlessExplicitCloseDiscardsThem(bool discard)
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcp-persistence-{Guid.NewGuid():N}.xlsx");
        var failure = await Record.ExceptionAsync(() => CheckPersistenceAsync(path, discard));
        var deleteFailure = Record.Exception(() => File.Delete(path));
        ThrowPersistenceFailures(failure, deleteFailure);
    }

    [Fact]
    public async Task Shutdown_ReaderStartupFailureClosesOnlyTheHandedOverFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mcp-persistence-{Guid.NewGuid():N}.xlsx");
        var startupFailure = new InvalidOperationException("Simulated failure before starting the reader.");
        var failure = await Record.ExceptionAsync(() => CheckPersistenceAsync(path, discard: false, startupFailure));
        var deleteFailure = Record.Exception(() => File.Delete(path));
        Assert.Null(deleteFailure);
        Assert.Same(startupFailure, failure);
    }

    private async Task CheckPersistenceAsync(string path, bool discard, Exception? readerStartupFailure = null)
    {
        using var shutdown = new CancellationTokenSource();
        var input = new Pipe();
        var outputPipe = new Pipe();
        McpClient? writerClient = null;
        Task? writerTask = null;
        ExcelProcessIdentity? capturedExcel = null;
        var writeFailure = await Record.ExceptionAsync(async () =>
        {
            var host = await ProgramTransportTestHost.StartAsync(input, outputPipe, shutdown.Token, "PersistenceWriter",
                onServerStarted: task => writerTask = task);
            writerClient = host.Client;
            var created = await CallAsync(writerClient, "file", new() { ["action"] = "create", ["file_path"] = path });
            var id = created.GetProperty("workbook_session_id").GetString()!;
            var discovery = await CallAsync(writerClient, "file_read", new() { ["action"] = "discover" });
            var fixture = Assert.Single(discovery.GetProperty("workbooks").EnumerateArray(),
                workbook => workbook.GetProperty("location").GetString() == path);
            Assert.Equal(id, fixture.GetProperty("sessionId").GetString());
            var pid = fixture.GetProperty("processId").GetInt32();
            using var excel = Process.GetProcessById(pid);
            var identity = new ExcelProcessIdentity(pid, excel.StartTime.ToUniversalTime().ToFileTimeUtc());
            Assert.True(OwnedProcessGuard.MatchesLiveProcess(identity));
            capturedExcel = identity;
            output.WriteLine($"Test-owned Excel: {identity}.");
            await CallAsync(writerClient, "range", new()
            {
                ["action"] = "set-values",
                ["workbook_session_id"] = id,
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "A1",
                ["values"] = new object[][] { ["saved-on-shutdown"] }
            });
            if (discard)
                await CallAsync(writerClient, "file", new() { ["action"] = "close", ["workbook_session_id"] = id, ["save"] = false });
        });
        var writerStopFailure = await Record.ExceptionAsync(() =>
            ProgramTransportTestHost.StopAsync(writerClient, input, outputPipe, writerTask, output, shutdown));
        if (writeFailure is not null || writerStopFailure is not null)
        {
            var writerCleanup = capturedExcel is { } identity && writerTask is { IsCompleted: true }
                ? await Record.ExceptionAsync(() => CleanupFailedHandoverAsync(identity, path))
                : null;
            ThrowPersistenceFailures(writeFailure, writerStopFailure, writerCleanup);
        }
        var writerExcel = Assert.IsType<ExcelProcessIdentity>(capturedExcel);

        using var readerShutdown = new CancellationTokenSource();
        var readerInput = new Pipe();
        var readerOutput = new Pipe();
        McpClient? readerClient = null;
        Task? readerTask = null;
        string? attachedId = null;
        var cleanupReady = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var readerStopped = new TaskCompletionSource<(bool Stopped, bool Failed)>(TaskCreationOptions.RunContinuationsAsynchronously);
        Task? cleanup = null;
        var readFailure = await Record.ExceptionAsync(async () =>
        {
            if (readerStartupFailure is not null)
                throw readerStartupFailure;
            var reader = await ProgramTransportTestHost.StartAsync(readerInput, readerOutput, readerShutdown.Token,
                "PersistenceReader", onServerStarted: task => readerTask = task);
            readerClient = reader.Client;
            if (!discard)
            {
                var discovery = await CallAsync(reader.Client, "file_read", new() { ["action"] = "discover" });
                var selected = Assert.Single(discovery.GetProperty("workbooks").EnumerateArray(),
                    workbook => workbook.GetProperty("location").GetString() == path);
                Assert.Equal(writerExcel.ProcessId, selected.GetProperty("processId").GetInt32());
                Assert.True(OwnedProcessGuard.MatchesLiveProcess(writerExcel));
                // Keep the fixture application available after the reader closes its last workbook.
                cleanup = StartHandoverCleanup(writerExcel, path, cleanupReady, readerStopped.Task);
                await cleanupReady.Task.WaitAsync(TimeSpan.FromSeconds(10));
                var attached = await CallAsync(reader.Client, "file", new()
                {
                    ["action"] = "attach",
                    ["open_workbook_id"] = selected.GetProperty("workbookId").GetString()
                });
                attachedId = attached.GetProperty("workbook_session_id").GetString();
                var retained = await CallAsync(reader.Client, "range_read", new()
                {
                    ["action"] = "get-values",
                    ["workbook_session_id"] = attachedId,
                    ["sheet_name"] = "Sheet1",
                    ["range_address"] = "A1"
                });
                Assert.Equal("saved-on-shutdown", retained.GetProperty("values")[0][0].GetString());
                await CallAsync(reader.Client, "file", new()
                {
                    ["action"] = "close",
                    ["workbook_session_id"] = attachedId,
                    ["save"] = false
                });
                attachedId = null;
            }
            var opened = await CallAsync(reader.Client, "file", new() { ["action"] = "open", ["file_path"] = path });
            var values = await CallAsync(reader.Client, "range_read", new()
            {
                ["action"] = "get-values",
                ["workbook_session_id"] = opened.GetProperty("workbook_session_id").GetString(),
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "A1"
            });
            Assert.Equal(JsonValueKind.Null, values.GetProperty("values")[0][0].ValueKind);
        });
        var closeFailure = await Record.ExceptionAsync(async () =>
        {
            if (readerClient is not null && attachedId is not null)
                await CallAsync(readerClient, "file", new()
                {
                    ["action"] = "close",
                    ["workbook_session_id"] = attachedId,
                    ["save"] = false
                });
        });
        var stopFailure = await Record.ExceptionAsync(() =>
            ProgramTransportTestHost.StopAsync(readerClient, readerInput, readerOutput, readerTask, output, readerShutdown));
        var stopped = readerTask is null or { IsCompleted: true };
        var failed = readFailure is not null || closeFailure is not null || stopFailure is not null;
        readerStopped.SetResult((stopped, failed));
        if (cleanup is null && stopped && failed && !OwnedProcessGuard.TryConfirmExited(writerExcel))
            cleanup = CleanupFailedHandoverAsync(writerExcel, path);
        var quitFailure = cleanup is not null
            ? await Record.ExceptionAsync(() => cleanup.WaitAsync(ComInteropConstants.StaThreadJoinTimeout))
            : null;
        var exitFailure = Record.Exception(() =>
            Assert.True(SpinWait.SpinUntil(() => OwnedProcessGuard.TryConfirmExited(writerExcel),
                TimeSpan.FromSeconds(10)), $"The test-owned Excel process is still running: {writerExcel}."));
        ThrowPersistenceFailures(readFailure, closeFailure, stopFailure, quitFailure, exitFailure);
    }

    private static async Task CleanupFailedHandoverAsync(ExcelProcessIdentity identity, string path)
    {
        if (OwnedProcessGuard.TryConfirmExited(identity))
            return;

        await StartHandoverCleanup(identity, path, null, Task.FromResult((true, true)))
            .WaitAsync(ComInteropConstants.StaThreadJoinTimeout);
        Assert.True(SpinWait.SpinUntil(() => OwnedProcessGuard.TryConfirmExited(identity),
            TimeSpan.FromSeconds(10)), $"The test-owned Excel process is still running: {identity}.");
    }

    private static Task StartHandoverCleanup(ExcelProcessIdentity identity, string path,
        TaskCompletionSource? ready, Task<(bool Stopped, bool Failed)> readerStopped)
    {
        var finished = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            Excel.Application? application = null;
            Excel.Workbooks? books = null;
            Excel.Workbook? fixture = null;
            Exception? failure = null;
            try
            {
                var window = Assert.Single(RunningExcelWindows.Find().Windows, candidate => candidate.Process == identity);
                application = RunningExcelWindows.Connect(window);
                books = application.Workbooks;
                fixture = books.Item[Path.GetFileName(path)];
                Assert.Equal(path, fixture.FullName, ignoreCase: true);
                ComUtilities.Release(ref fixture);
                ComUtilities.Release(ref books);
                ready?.SetResult();

                // A normal attachment never owns Excel. Only the fixture owner may quit it when empty.
                var handover = readerStopped.GetAwaiter().GetResult();
                if (handover.Stopped && !OwnedProcessGuard.TryConfirmExited(identity))
                {
                    Assert.True(OwnedProcessGuard.MatchesLiveProcess(identity));
                    if (handover.Failed)
                    {
                        books = application.Workbooks;
                        if (books.Count != 0)
                        {
                            Assert.Equal(1, books.Count);
                            fixture = books.Item[Path.GetFileName(path)];
                            Assert.Equal(path, fixture.FullName, ignoreCase: true);
                        }
                        ComUtilities.Release(ref books);
                    }
                    var quit = ExcelShutdownService.CloseAndQuitAndReport(fixture, application, save: false,
                        filePath: path, releaseWorkbook: false);
                    application = null; // A completed close released the application reference.
                    Assert.True(quit,
                        $"The test-owned Excel could not be closed: {identity}.");
                }
            }
            catch (Exception ex)
            {
                ready?.TrySetException(ex);
                failure = ex;
            }
            finally
            {
                ComUtilities.Release(ref fixture);
                ComUtilities.Release(ref books);
                ComUtilities.Release(ref application);
            }
            if (failure is null)
                finished.SetResult();
            else
                finished.SetException(failure);
        })
        { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return finished.Task;
    }

    private static void ThrowPersistenceFailures(params Exception?[] errors)
    {
        var failures = errors.OfType<Exception>().ToArray();
        if (failures.Length == 1)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failures[0]).Throw();
        if (failures.Length > 0)
            throw new AggregateException("MCP persistence verification or owned Excel cleanup failed.", failures);
    }

    private static async Task<JsonElement> CallAsync(McpClient client, string tool, Dictionary<string, object?> args)
    {
        var result = await client.CallToolAsync(tool, args);
        Assert.NotEqual(true, result.IsError);
        Assert.NotNull(result.StructuredContent);
        Assert.True(result.StructuredContent.Value.GetProperty("success").GetBoolean());
        return result.StructuredContent.Value;
    }

    private sealed class ObservedBackend : IServiceBridgeBackend
    {
        private readonly ExcelMcpService _service = new();
        internal SessionManager SessionManager => _service.SessionManager;
        internal bool ObserveCreation { get; set; }
        internal TaskCompletionSource<ServiceResponse> Finished { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource<string> Revealed { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal string? ShutdownReport { get; private set; }

        public string? RequestShutdown() => ShutdownReport = _service.RequestShutdown();

        public async Task<ServiceResponse> ProcessAsync(ServiceRequest request)
        {
            var observe = ObserveCreation && request.Command == "session.create";
            var response = await _service.ProcessAsync(request);
            if (observe)
                Finished.TrySetResult(response);
            return response;
        }

        public void RevealSession(string sessionId)
        {
            _service.SessionManager.RevealForRecovery(sessionId, waitExpired: true);
            Revealed.TrySetResult(sessionId);
        }

        public void Dispose() => _service.Dispose();
    }
}
