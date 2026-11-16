using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service.Rpc;
using Sbroenne.ExcelMcp.Service.Workers;
using StreamJsonRpc;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ExcelControllerTests
{
    [Fact(Timeout = 15000)]
    public Task Controller_DiscoveryExposesRetainedSessionOnlyAfterAttaching() =>
        WithDiscoveryWorkerAsync(retained: true, async (controller, worker) =>
        {
            var discovered = await ReadDiscoveryAsync(controller);
            Assert.True(discovered.GetProperty("complete").GetBoolean());
            var workbook = discovered.GetProperty("workbooks")[0];
            Assert.False(workbook.TryGetProperty("sessionId", out _));

            var attached = await AttachDiscoverySelectionAsync(controller);
            Assert.True(attached.Success, attached.ErrorMessage);
            Assert.Equal(DiscoveryWorker.SessionId, attached.SessionId);
            Assert.Equal(1, worker.AttachCalls);
            var rechecked = await ReadDiscoveryAsync(controller);
            Assert.Equal(DiscoveryWorker.SessionId,
                rechecked.GetProperty("workbooks")[0].GetProperty("sessionId").GetString());
        });

    [Fact(Timeout = 15000)]
    public Task Controller_DiscoveryRetiresAbsentSelectionOnlyAfterCompleteInspection() =>
        WithDiscoveryWorkerAsync(retained: false, async (controller, worker) =>
        {
            Assert.Equal(1, (await ReadDiscoveryAsync(controller)).GetProperty("workbooks").GetArrayLength());
            worker.IncludeWorkbook = false;
            worker.Complete = false;
            Assert.False((await ReadDiscoveryAsync(controller)).GetProperty("complete").GetBoolean());
            var partial = await AttachDiscoverySelectionAsync(controller);
            Assert.False(partial.Success);
            Assert.Equal("Selection is no longer open.", partial.ErrorMessage);
            Assert.Equal(1, worker.AttachCalls);

            worker.Complete = true;
            var complete = await ReadDiscoveryAsync(controller);
            Assert.True(complete.GetProperty("complete").GetBoolean());
            Assert.Empty(complete.GetProperty("workbooks").EnumerateArray());
            var expired = await AttachDiscoverySelectionAsync(controller);
            Assert.False(expired.Success);
            Assert.Equal("Unknown open-workbook selection. Discover open workbooks first.", expired.ErrorMessage);
            Assert.Equal(1, worker.AttachCalls);
        });

    [Fact(Timeout = 15000)]
    public Task Controller_DiscoveryDoesNotReplaceNewerAttachmentWithOlderReply() =>
        WithDiscoveryWorkerAsync(retained: false, async (controller, worker) =>
        {
            var first = await ReadDiscoveryAsync(controller);
            Assert.False(first.GetProperty("workbooks")[0].TryGetProperty("sessionId", out _));
            var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            worker.BeforeDiscoveryReply = async () =>
            {
                entered.TrySetResult();
                await release.Task;
            };
            var discovery = ReadDiscoveryAsync(controller);
            try
            {
                await entered.Task.WaitAsync(TimeSpan.FromSeconds(3));
                var attached = await AttachDiscoverySelectionAsync(controller);
                Assert.True(attached.Success, attached.ErrorMessage);
                Assert.Equal(DiscoveryWorker.SessionId, attached.SessionId);
            }
            finally
            {
                release.TrySetResult();
            }

            var newer = await discovery;
            Assert.Equal(DiscoveryWorker.SessionId,
                newer.GetProperty("workbooks")[0].GetProperty("sessionId").GetString());
        });

    private static async Task<JsonElement> ReadDiscoveryAsync(ExcelController controller)
    {
        var result = await controller.ProcessAsync(new ServiceRequest { Command = "session.discover" });
        Assert.True(result.Success, result.ErrorMessage);
        using var json = JsonDocument.Parse(result.Result!);
        return json.RootElement.Clone();
    }

    private static Task<ServiceResponse> AttachDiscoverySelectionAsync(ExcelController controller) =>
        controller.ProcessAsync(new ServiceRequest
        {
            Command = "session.attach",
            Args = """{"workbookId":"selection"}"""
        });

    private static async Task WithDiscoveryWorkerAsync(bool retained, Func<ExcelController, DiscoveryWorker, Task> run)
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-discovery-selection-");
        var registry = new WorkerRegistry(directory.FullName);
        var process = WorkerProcess.Current;
        var native = new ExcelProcessIdentity(process.ProcessId, process.StartedAtUtcFileTime);
        var previous = process with { StartedAtUtcFileTime = process.StartedAtUtcFileTime - 1 };
        var registration = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", process, previous,
            "previous", WorkerRegistry.Build, native, OwnsExcel: true);
        var target = new DiscoveryWorker(registration, Path.Combine(directory.FullName, "book.xlsx"), retained);
        registry.Write(registration);
        using var controller = new ExcelController(launcher: Launcher, registry: registry);
        await StopNativeObservationPollingAsync(controller);
        using var host = new DaemonHost(registration.PipeName, pipe => new JsonRpc(pipe, pipe, target),
            maximumConnections: 8);
        var listener = host.RunAsync(registration.PipeName);
        RunningExcelWindows.FindHookForTests = () => new RunningExcelWindows.Discovery(
            [new RunningExcelWindows.ApplicationWindow(nint.Zero, native)], []);
        try
        {
            await run(controller, target);
        }
        finally
        {
            RunningExcelWindows.FindHookForTests = null;
            await controller.ShutdownAsync();
            host.RequestShutdown();
            await listener.WaitAsync(TimeSpan.FromSeconds(3));
            registry.Remove(registration.PipeName);
            directory.Delete();
        }
    }

    private sealed class DiscoveryWorker(WorkerRegistration registration, string path, bool attached) : IExcelWorkerRpc
    {
        internal const string SessionId = "worker-session";
        private long _revision;
        internal bool IncludeWorkbook { get; set; } = true;
        internal bool Complete { get; set; } = true;
        internal Func<Task>? BeforeDiscoveryReply { get; set; }
        internal int AttachCalls { get; private set; }

        public Task<WorkerSnapshot> ClaimAsync(WorkerClaim claim, CancellationToken cancellationToken = default)
        {
            registration = registration with { Controller = claim.Controller, Generation = claim.Generation };
            return Task.FromResult(Snapshot());
        }

        public Task<WorkerSnapshot> SnapshotAsync(string generation, CancellationToken cancellationToken = default) =>
            Task.FromResult(Snapshot());

        public async Task<WorkerExecution> ExecuteAsync(WorkerCall call, CancellationToken cancellationToken = default)
        {
            if (call.Request.Command == "session.discover")
            {
                OpenWorkbookInfo[] workbooks = IncludeWorkbook
                    ? [new("selection", "book.xlsx", attached ? path : null, registration.Worker.ProcessId, true, false,
                        attached ? SessionId : null)] : [];
                WorkbookDiscoveryError[] errors = Complete ? [] : [new(registration.Worker.ProcessId, "Incomplete fixture inspection.")];
                var execution = new WorkerExecution(new ServiceResponse
                {
                    Success = true,
                    Result = JsonSerializer.Serialize(new { workbooks, errors, complete = Complete }, ServiceProtocol.JsonOptions)
                }, true, Snapshot());
                if (BeforeDiscoveryReply is { } wait)
                    await wait();
                return execution;
            }

            Assert.Equal("session.attach", call.Request.Command);
            AttachCalls++;
            if (!IncludeWorkbook)
                return new(ExcelWorkerHost.Failure(call.Request, "InvalidInput", "Selection is no longer open."), false, Snapshot());
            attached = true;
            return new(new ServiceResponse { Success = true, SessionId = SessionId }, true, Snapshot());
        }

        public Task<ServiceResponse?> ObserveLateCompletionAsync(string generation, string operationId,
            CancellationToken cancellationToken = default) => Task.FromResult<ServiceResponse?>(null);

        public Task<WorkerRelease> ReleaseAsync(string generation, CancellationToken cancellationToken = default)
        {
            attached = false;
            registration = registration with { ReleaseConfirmed = true };
            return Task.FromResult(new WorkerRelease(new ServiceResponse { Success = true }, Snapshot()));
        }

        private WorkerSnapshot Snapshot()
        {
            SessionDescriptor[] sessions = attached ? [new(SessionId, path)] : [];
            var list = new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { sessions }, ServiceProtocol.JsonOptions)
            };
            var owners = new Dictionary<string, IReadOnlyList<string>>();
            if (attached)
                owners.Add(SessionId, [path]);
            return new WorkerSnapshot(registration, sessions, list, null, null, owners,
                Revision: Interlocked.Increment(ref _revision));
        }
    }
}
