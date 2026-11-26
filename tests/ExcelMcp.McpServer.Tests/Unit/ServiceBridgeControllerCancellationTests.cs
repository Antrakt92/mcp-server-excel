using System.Reflection;
using System.Text.Json;
using Sbroenne.ExcelMcp.McpServer.ServiceBridge;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Service.Rpc;
using Sbroenne.ExcelMcp.Service.Workers;
using Xunit;
using Bridge = Sbroenne.ExcelMcp.McpServer.ServiceBridge.ServiceBridge;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Layer", "McpServer")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceBridge")]
[Trait("Speed", "Medium")]
[Trait("RequiresExcel", "false")]
public sealed class ServiceBridgeControllerCancellationTests
{
    [Theory(Timeout = 15000)]
    [InlineData(false)]
    [InlineData(true)]
    public async Task TrackedSave_RecoversOriginalOutcomeWithoutAnotherDispatch(bool nativeWaitExpired)
    {
        const string session = "owned-session";
        var operationId = Guid.NewGuid().ToString();
        var pipeName = $"excelmcp-worker-{Guid.NewGuid():N}";
        var target = new DeferredWorker(pipeName, nativeWaitExpired);
        using var host = new DaemonHost(pipeName,
            pipe => new StreamJsonRpc.JsonRpc(pipe, pipe, target), maximumConnections: 8);
        var run = host.RunAsync(pipeName);
        var registry = Directory.CreateTempSubdirectory("excelmcp-receipt-test-");
        using var controller = new ExcelController(launcher: new("unused.exe", []),
            registry: new WorkerRegistry(registry.FullName));
        var worker = new ExcelWorkerClient(new(pipeName, WorkerProcess.Current, WorkerProcess.Current,
            "generation", WorkerRegistry.Build, null), WorkerProcess.Current, "generation");
        typeof(ExcelController).GetField("_utility", BindingFlags.NonPublic | BindingFlags.Instance)!
            .SetValue(controller, Task.FromResult(worker));
        using var bridge = new Bridge(() => new ExcelControllerBackend(controller), maxOutstandingDispatches: 2);
        using var caller = new CancellationTokenSource();
        var save = bridge.SendAsync("session.save", session, cancellationToken: caller.Token, operationId: operationId);
        try
        {
            await target.Started.Task.WaitAsync(TimeSpan.FromSeconds(3));
            if (nativeWaitExpired)
            {
                Assert.Equal("Timeout", (await save.WaitAsync(TimeSpan.FromSeconds(3))).ErrorCategory);
                await target.Observing.Task.WaitAsync(TimeSpan.FromSeconds(3));
            }
            else
            {
                caller.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => save);
                await target.Cancelled.Task.WaitAsync(TimeSpan.FromSeconds(3));
            }

            using var pending = JsonDocument.Parse((await bridge.SendAsync(OperationReceipts.LookupCommand,
                session, new { operationId })).Result!);
            Assert.Equal("pending", pending.RootElement.GetProperty("state").GetString());
            var duplicate = await bridge.SendAsync("session.save", session, operationId: operationId);
            Assert.Equal("Conflict", duplicate.ErrorCategory);
            Assert.Equal(1, target.Calls);

            target.Finish.TrySetResult();
            using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(3));
            while (true)
            {
                var response = await bridge.SendAsync(OperationReceipts.LookupCommand,
                    session, new { operationId }, cancellationToken: deadline.Token);
                using var receipt = JsonDocument.Parse(response.Result!);
                var root = receipt.RootElement;
                if (root.GetProperty("state").GetString() == "pending")
                {
                    await Task.Delay(10, deadline.Token);
                    continue;
                }
                Assert.Equal("completed", root.GetProperty("state").GetString());
                Assert.Equal(operationId, root.GetProperty("operationId").GetString());
                Assert.True(root.GetProperty("result").GetProperty("saved").GetBoolean());
                break;
            }
            Assert.Equal(1, target.Calls);
        }
        finally
        {
            target.Finish.TrySetResult();
            host.RequestShutdown();
            await run.WaitAsync(TimeSpan.FromSeconds(3));
            registry.Delete();
        }
    }

    [Fact(Timeout = 15000)]
    public async Task CancelledCallerReturnsWhileControllerKeepsWorkerDispatchAdmitted()
    {
        var pipeName = $"excelmcp-worker-{Guid.NewGuid():N}";
        var target = new DeferredWorker(pipeName);
        using var host = new DaemonHost(pipeName,
            pipe => new StreamJsonRpc.JsonRpc(pipe, pipe, target), maximumConnections: 8);
        var run = host.RunAsync(pipeName);
        var registry = Directory.CreateTempSubdirectory("excelmcp-dispatch-test-");
        using var controller = new ExcelController(launcher: new("unused.exe", []),
            registry: new WorkerRegistry(registry.FullName));
        var worker = new ExcelWorkerClient(new(pipeName, WorkerProcess.Current, WorkerProcess.Current,
            "generation", WorkerRegistry.Build, null), WorkerProcess.Current, "generation");
        // Route through a real controller and worker RPC without launching Excel or a child process.
        typeof(ExcelController).GetField("_utility", BindingFlags.NonPublic | BindingFlags.Instance)!
            .SetValue(controller, Task.FromResult(worker));
        using var bridge = new Bridge(() => new ExcelControllerBackend(controller), maxOutstandingDispatches: 1);
        using var caller = new CancellationTokenSource();
        var cancelled = bridge.SendAsync("diag.echo", args: new { message = "hold" }, cancellationToken: caller.Token);
        try
        {
            await target.Started.Task.WaitAsync(TimeSpan.FromSeconds(3));
            caller.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => cancelled.WaitAsync(TimeSpan.FromSeconds(3)));
            await target.Cancelled.Task.WaitAsync(TimeSpan.FromSeconds(3));
            // Give an early controller cancellation facade time to reach the bridge's completion observer.
            await Task.Delay(250);
            Assert.False(target.Finish.Task.IsCompleted);
            var busy = await bridge.SendAsync("diag.echo", args: new { message = "probe" });
            Assert.Equal("ServiceBusy", busy.ErrorCategory);
            Assert.Equal(1, target.Calls);

            target.Finish.TrySetResult();
            using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(3));
            while (worker.Snapshot?.Revision != 1)
                await Task.Delay(10, deadline.Token);
            while (true)
            {
                var probe = await bridge.SendAsync("diag.echo", args: new { message = "probe" },
                    cancellationToken: deadline.Token);
                if (probe.Success)
                    break;
                Assert.Equal("ServiceBusy", probe.ErrorCategory);
                await Task.Delay(10, deadline.Token);
            }
            Assert.Equal(2, target.Calls);
            Assert.Equal(2, worker.Snapshot!.Revision);
        }
        finally
        {
            target.Finish.TrySetResult();
            host.RequestShutdown();
            await run.WaitAsync(TimeSpan.FromSeconds(3));
            registry.Delete();
        }
    }

    private sealed class DeferredWorker(string pipeName, bool nativeWaitExpired = false) : IExcelWorkerRpc
    {
        private int _calls;
        internal int Calls => Volatile.Read(ref _calls);
        internal TaskCompletionSource Started { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource Cancelled { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource Finish { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource Observing { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public async Task<WorkerExecution> ExecuteAsync(WorkerCall call, CancellationToken cancellationToken = default)
        {
            var revision = Interlocked.Increment(ref _calls);
            if (revision == 1)
            {
                using var registration = cancellationToken.Register(() => Cancelled.TrySetResult());
                Started.TrySetResult();
                if (!nativeWaitExpired)
                    await Finish.Task;
            }
            var snapshot = new WorkerSnapshot(
                new(pipeName, WorkerProcess.Current, WorkerProcess.Current, call.Generation, WorkerRegistry.Build, null),
                [], new ServiceResponse { Success = true, Result = """{"sessions":[]}""" },
                null, null, new Dictionary<string, IReadOnlyList<string>>(), Revision: revision);
            var response = nativeWaitExpired
                ? new ServiceResponse { Success = false, ErrorCategory = "Timeout" }
                : call.Request.Command == "session.save" ? Saved() : new ServiceResponse { Success = true, Result = "{}" };
            return new(response, Accepted: true, snapshot, HasLateNativeCompletion: nativeWaitExpired);
        }

        public Task<WorkerSnapshot> ClaimAsync(WorkerClaim claim, CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();

        public Task<WorkerSnapshot> SnapshotAsync(string generation, CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();

        public async Task<ServiceResponse?> ObserveLateCompletionAsync(string generation, string operationId,
            CancellationToken cancellationToken = default)
        {
            Observing.TrySetResult();
            await Finish.Task.WaitAsync(cancellationToken);
            return Saved();
        }

        private static ServiceResponse Saved() => new() { Success = true, Result = """{"saved":true}""" };

        public Task<WorkerRelease> ReleaseAsync(string generation, bool onlyIfEmpty, CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();
    }
}
