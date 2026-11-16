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
    [InlineData(false)]
    [InlineData(true)]
    public async Task Controller_ReleasesUnchangedSaveAsDestinationAfterWorkerSettles(bool statusUnavailable)
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
            var original = Path.Join(directory.FullName, "original.xlsx");
            var destination = Path.Join(directory.FullName, "destination.xlsx");

            WorkerSnapshot Snapshot(string path, long revision, string? statusError = null) =>
                new(registration, [new SessionDescriptor(sessionId, path)], new ServiceResponse
                {
                    Success = true,
                    Result = JsonSerializer.Serialize(new { sessions = new[] { new { sessionId, filePath = path } } },
                        ServiceProtocol.JsonOptions)
                }, null, null, new Dictionary<string, IReadOnlyList<string>> { [sessionId] = [path] },
                    StatusError: statusError, Revision: revision);

            var before = Snapshot(original, 1);
            var timeout = Snapshot(original, 3, statusUnavailable ? "Excel status could not be refreshed." : null);
            var completed = Snapshot(destination, 5);
            var target = new SaveAsReservationWorker(timeout, completed);
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

                var expired = await save;
                Assert.False(expired.Success);
                Assert.Equal("Timeout", expired.ErrorCategory);
                if (statusUnavailable)
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

    private sealed class SaveAsReservationWorker(WorkerSnapshot timeout, WorkerSnapshot completed) : IExcelWorkerRpc
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
                    ErrorCategory = "Timeout",
                    PartialState = """{"operation":{"waitExpired":true}}"""
                }, Accepted: true, timeout);
            }
            return new(new ServiceResponse { Success = true }, Accepted: true, completed);
        }

        public Task<WorkerSnapshot> ClaimAsync(WorkerClaim claim, CancellationToken cancellationToken = default) =>
            Task.FromResult(timeout);

        public Task<WorkerSnapshot> SnapshotAsync(string generation, CancellationToken cancellationToken = default) =>
            Task.FromResult(completed);

        public Task<ServiceResponse?> ObserveLateCompletionAsync(string generation, string operationId,
            CancellationToken cancellationToken = default) => Task.FromResult<ServiceResponse?>(null);

        public Task<WorkerRelease> ReleaseAsync(string generation, CancellationToken cancellationToken = default) =>
            Task.FromResult(new WorkerRelease(new ServiceResponse { Success = true }, completed with
            {
                Registration = completed.Registration with { ReleaseConfirmed = true },
                Revision = completed.Revision + 1
            }));
    }
}
