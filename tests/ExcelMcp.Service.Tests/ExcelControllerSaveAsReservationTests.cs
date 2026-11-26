using System.Collections.Concurrent;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service.Rpc;
using Sbroenne.ExcelMcp.Service.Workers;
using StreamJsonRpc;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ExcelControllerTests
{
    [Theory(Timeout = 15000)]
    [InlineData("Timeout", false, false, false, true)]
    [InlineData("Timeout", true, false, false, true)]
    [InlineData("Timeout", false, false, true, true)]
    [InlineData("Cancelled", false, false, true, false)]
    [InlineData("Cancelled", false, false, true, true)]
    [InlineData("Busy", false, true, false, false)]
    [InlineData("Busy", false, true, true, false)]
    [InlineData("Busy", false, true, false, true)]
    [InlineData("InvalidInput", true, false, false, false)]
    [InlineData("InvalidInput", true, false, false, true)]
    [InlineData("ExcelDisconnected", false, false, false, true)]
    public async Task Controller_ReconcilesSaveAsDestinationWithoutLosingUncertainOutcomes(
        string errorCategory, bool statusUnavailable, bool dialogOpen, bool unsaved, bool includeWaitDiagnostics)
    {
        const string sessionId = "save-as-session";
        var directory = Directory.CreateTempSubdirectory("excelmcp-save-as-reservation-");
        try
        {
            var registry = new WorkerRegistry(directory.FullName);
            using var controller = new ExcelController(launcher: Launcher, registry: registry);
            await StopNativeObservationPollingAsync(controller);
            var type = typeof(ExcelController);
            var generation = Assert.IsType<string>(
                type.GetField("_generation", NativeObservationHidden)!.GetValue(controller));
            var current = WorkerProcess.Current;
            var registration = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", current, current,
                generation, WorkerRegistry.Build, Excel: null);
            var original = unsaved ? null : Path.Join(directory.FullName, "original.xlsx");
            var destination = Path.Join(directory.FullName, "destination.xlsx");

            WorkerSnapshot Snapshot(string? path, long revision, string? statusError = null, bool operationPending = false) =>
                new(registration, [new SessionDescriptor(sessionId, path)], new ServiceResponse
                {
                    Success = true,
                    Result = JsonSerializer.Serialize(new { sessions = new[] { new { sessionId, filePath = path, operationPending } } },
                        ServiceProtocol.JsonOptions)
                }, null, null, new Dictionary<string, IReadOnlyList<string>> { [sessionId] = path is null ? [] : [path] },
                    StatusError: statusError, Revision: revision);

            var before = Snapshot(original, 1);
            var failure = Snapshot(original, 3, statusUnavailable ? "Excel status could not be refreshed." : null, dialogOpen);
            var completed = Snapshot(destination, 5);
            const string settledWait = """{"operation":{"state":"completed","waitExpired":true,"executionPending":false}}""";
            var partialState = errorCategory switch
            {
                "ExcelDisconnected" => """{"outcome":"unknown","operation":{"state":"completed","waitExpired":true,"executionPending":false}}""",
                "Timeout" or "Cancelled" when includeWaitDiagnostics =>
                    """{"operation":{"command":"workbook.save-as","state":"running","waitExpired":true,"executionPending":true}}""",
                _ => includeWaitDiagnostics ? settledWait : null
            };
            var target = new SaveAsReservationWorker(errorCategory, partialState, failure, completed);
            using var host = new DaemonHost(registration.PipeName, pipe => new JsonRpc(pipe, pipe, target),
                maximumConnections: 8);
            var listener = host.RunAsync(registration.PipeName);
            try
            {
                var worker = new ExcelWorkerClient(registration, current, generation);
                worker.RecordSnapshot(before);
                var workers = Assert.IsType<ConcurrentDictionary<string, ExcelWorkerClient>>(
                    type.GetField("_workers", NativeObservationHidden)!.GetValue(controller));
                Assert.True(workers.TryAdd(registration.PipeName, worker));
                var sessions = Assert.IsType<ConcurrentDictionary<string, ExcelWorkerClient>>(
                    type.GetField("_sessions", NativeObservationHidden)!.GetValue(controller));
                Assert.True(sessions.TryAdd(sessionId, worker));
                var apply = type.GetMethod("ApplyWorkerSnapshot", NativeObservationHidden)!;

                Task<ServiceResponse> SaveAsAsync() => controller.ProcessAsync(new ServiceRequest
                {
                    Command = "workbook.save-as",
                    SessionId = sessionId,
                    Args = JsonSerializer.Serialize(new { targetPath = destination })
                });

                var save = SaveAsAsync();
                try
                {
                    await target.Entered.Task.WaitAsync(TimeSpan.FromSeconds(3));
                    worker.RecordSnapshot(Snapshot(original, 2));
                    apply.Invoke(controller, [worker]);
                    Assert.Equal("Conflict", (await SaveAsAsync()).ErrorCategory);
                    Assert.Equal(1, target.Calls);
                }
                finally
                {
                    target.Release.TrySetResult();
                }

                var failed = await save;
                Assert.False(failed.Success);
                Assert.Equal(errorCategory, failed.ErrorCategory);
                Assert.Equal(partialState, failed.PartialState);
                if (errorCategory == "ExcelDisconnected")
                {
                    worker.RecordSnapshot(Snapshot(original, 4));
                    apply.Invoke(controller, [worker]);
                    Assert.Equal("Conflict", (await SaveAsAsync()).ErrorCategory);
                    Assert.Equal(1, target.Calls);
                    return;
                }
                if (unsaved && includeWaitDiagnostics && errorCategory is "Timeout" or "Cancelled")
                {
                    worker.RecordSnapshot(Snapshot(original, 4));
                    apply.Invoke(controller, [worker]);
                    Assert.Equal("Conflict", (await SaveAsAsync()).ErrorCategory);
                    Assert.Equal(1, target.Calls);

                    // The late Save As can still publish its confirmed path. A request
                    // for the same destination then belongs to this session again.
                    worker.RecordSnapshot(completed);
                    apply.Invoke(controller, [worker]);
                }
                if (statusUnavailable || dialogOpen)
                {
                    Assert.Equal("Conflict", (await SaveAsAsync()).ErrorCategory);
                    Assert.Equal(1, target.Calls);
                    worker.RecordSnapshot(Snapshot(original, 4));
                    apply.Invoke(controller, [worker]);
                }

                var retried = await SaveAsAsync();
                Assert.True(retried.Success, retried.ErrorMessage);
                Assert.Equal(2, target.Calls);
                var list = await controller.ProcessAsync(new ServiceRequest { Command = "session.list" });
                Assert.True(list.Success, list.ErrorMessage);
                using var result = JsonDocument.Parse(list.Result!);
                Assert.Equal(destination, result.RootElement.GetProperty("sessions")[0].GetProperty("filePath").GetString());
            }
            finally
            {
                await controller.ShutdownAsync();
                host.RequestShutdown();
                await listener.WaitAsync(TimeSpan.FromSeconds(4));
            }
        }
        finally
        {
            directory.Delete(recursive: true);
        }
    }

    private sealed class SaveAsReservationWorker(
        string errorCategory, string? partialState, WorkerSnapshot failure, WorkerSnapshot completed) : IExcelWorkerRpc
    {
        internal TaskCompletionSource Entered { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource Release { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _calls;
        internal int Calls => Volatile.Read(ref _calls);

        public async Task<WorkerExecution> ExecuteAsync(WorkerCall call, CancellationToken cancellationToken = default)
        {
            Assert.Equal("workbook.save-as", call.Request.Command);
            if (Interlocked.Increment(ref _calls) == 1)
            {
                Entered.TrySetResult();
                await Release.Task.WaitAsync(cancellationToken);
                return new(new ServiceResponse
                {
                    Success = false,
                    ErrorCategory = errorCategory,
                    PartialState = partialState
                }, Accepted: true, failure);
            }
            return new(new ServiceResponse { Success = true }, Accepted: true, completed);
        }

        public Task<WorkerSnapshot> ClaimAsync(WorkerClaim claim, CancellationToken cancellationToken = default) =>
            Task.FromResult(failure);

        public Task<WorkerSnapshot> SnapshotAsync(string generation, CancellationToken cancellationToken = default) =>
            Task.FromResult(completed);

        public Task<ServiceResponse?> ObserveLateCompletionAsync(string generation, string operationId,
            CancellationToken cancellationToken = default) => Task.FromResult<ServiceResponse?>(null);

        public Task<WorkerRelease> ReleaseAsync(string generation, bool onlyIfEmpty, CancellationToken cancellationToken = default) =>
            Task.FromResult(new WorkerRelease(new ServiceResponse { Success = true }, completed with
            {
                Registration = completed.Registration with { ReleaseConfirmed = true },
                Revision = completed.Revision + 1
            }));
    }
}
