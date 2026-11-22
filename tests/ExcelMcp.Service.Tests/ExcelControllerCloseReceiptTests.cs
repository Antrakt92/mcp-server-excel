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
    [InlineData("no-save")]
    [InlineData("timing-only")]
    [InlineData("saved")]
    [InlineData("saved-with-recording-error")]
    [InlineData("failed-teardown")]
    [InlineData("snapshot-only")]
    [InlineData("release-only")]
    public async Task Controller_CloseReplayRequiresExplicitSaveReceipt(string scenario)
    {
        const string sessionId = "closed-with-or-without-evidence";
        var directory = Path.Combine(Path.GetTempPath(), $"excelmcp-close-receipt-{Guid.NewGuid():N}");
        var registry = new WorkerRegistry(directory);
        try
        {
            using var controller = new ExcelController(launcher: Launcher, registry: registry);
            await StopNativeObservationPollingAsync(controller);
            var generation = Assert.IsType<string>(typeof(ExcelController)
                .GetField("_generation", NativeObservationHidden)!.GetValue(controller));
            var current = WorkerProcess.Current;
            var registration = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", current,
                current, generation, WorkerRegistry.Build, Excel: null);
            var saveConfirmed = scenario is "saved" or "saved-with-recording-error";
            var recordingFailed = scenario == "saved-with-recording-error";
            var response = scenario switch
            {
                "saved" or "saved-with-recording-error" => new ServiceResponse { Success = true, Result = """{"saved":true}""" },
                "failed-teardown" => new ServiceResponse
                {
                    Success = false,
                    ErrorCategory = "InternalError",
                    ErrorMessage = "Teardown failed."
                },
                "timing-only" => new ServiceResponse { Success = true, Result = """{"closeTimingsMs":{"save":5}}""" },
                _ => new ServiceResponse { Success = true, Result = """{"saved":false}""" }
            };
            if (recordingFailed)
                response = response.WithCompletionFailure(
                    new ServiceRequest { Command = "session.close", SessionId = sessionId },
                    new IOException("Worker close publication failed."));
            var settled = new WorkerSnapshot(registration with { ReleaseConfirmed = scenario == "release-only" },
                [], new ServiceResponse { Success = true, Result = """{"success":true,"sessions":[],"count":0}""" },
                null, null, new Dictionary<string, IReadOnlyList<string>>(), Revision: 2);
            var filePath = Path.Combine(directory, "claimed.xlsx");
            var active = settled with
            {
                Registration = registration,
                Sessions = [new SessionDescriptor(sessionId, filePath)],
                SessionList = new ServiceResponse
                {
                    Success = true,
                    Result = JsonSerializer.Serialize(new { sessions = new[] { new { sessionId, filePath } } })
                },
                OwnershipKeys = new Dictionary<string, IReadOnlyList<string>> { [sessionId] = [filePath] },
                Revision = 1
            };
            var target = new CloseReceiptWorker(response, recordingFailed
                ? active with { StatusError = "Worker close publication failed.", Revision = 2 }
                : settled);
            using var host = new DaemonHost(registration.PipeName,
                pipe => new JsonRpc(pipe, pipe, target), maximumConnections: 8);
            var run = host.RunAsync(registration.PipeName);
            try
            {
                var worker = new ExcelWorkerClient(registration, current, generation);
                worker.RecordSnapshot(active);
                var workers = Assert.IsType<System.Collections.Concurrent.ConcurrentDictionary<string, ExcelWorkerClient>>(
                    typeof(ExcelController).GetField("_workers", NativeObservationHidden)!.GetValue(controller));
                Assert.True(workers.TryAdd(registration.PipeName, worker));
                var sessions = Assert.IsType<System.Collections.Concurrent.ConcurrentDictionary<string, ExcelWorkerClient>>(
                    typeof(ExcelController).GetField("_sessions", NativeObservationHidden)!.GetValue(controller));
                Assert.True(sessions.TryAdd(sessionId, worker));
                var applySnapshot = typeof(ExcelController).GetMethod("ApplyWorkerSnapshot", NativeObservationHidden)!;
                var paths = Assert.IsAssignableFrom<System.Collections.IEnumerable>(
                    typeof(ExcelController).GetField("_paths", NativeObservationHidden)!.GetValue(controller));
                applySnapshot.Invoke(controller, [worker]);
                Assert.Single(paths.Cast<object>());

                if (scenario is "snapshot-only" or "release-only")
                {
                    worker.RecordSnapshot(settled);
                    applySnapshot.Invoke(controller, [worker]);
                }
                else
                {
                    var first = await controller.ProcessAsync(new ServiceRequest
                    {
                        Command = "session.close",
                        SessionId = sessionId,
                        Args = scenario == "no-save" ? """{"save":false}""" : """{"save":true}"""
                    });
                    Assert.Equal(response.Success, first.Success);
                    if (recordingFailed)
                    {
                        Assert.Equal("ServiceUnavailable", first.ErrorCategory);
                        Assert.Equal(response.ErrorMessage, first.ErrorMessage);
                        Assert.Equal(response.PartialState, first.PartialState);
                        Assert.Same(worker, sessions[sessionId]);
                        Assert.Single(paths.Cast<object>());
                        Assert.Equal(0, target.ReleaseCalls);
                        worker.RecordSnapshot(settled with { Revision = 3 });
                        applySnapshot.Invoke(controller, [worker]);
                        Assert.False(sessions.ContainsKey(sessionId));
                        Assert.Empty(paths.Cast<object>());
                    }
                }

                var replay = await controller.ProcessAsync(new ServiceRequest
                {
                    Command = "session.close",
                    SessionId = sessionId,
                    Args = """{"save":true}"""
                });
                Assert.Equal(saveConfirmed, replay.Success);
                if (saveConfirmed)
                {
                    using var receipt = JsonDocument.Parse(replay.Result!);
                    Assert.True(receipt.RootElement.GetProperty("saved").GetBoolean());
                }
                else
                    Assert.Equal("SessionUnavailable", replay.ErrorCategory);
                Assert.Equal(scenario is "snapshot-only" or "release-only" ? 0 : 1, target.Calls);
                Assert.True((await controller.ProcessAsync(new ServiceRequest
                {
                    Command = "session.close",
                    SessionId = sessionId
                })).Success);
            }
            finally
            {
                await controller.ShutdownAsync();
                host.RequestShutdown();
                await run.WaitAsync(TimeSpan.FromSeconds(4));
            }
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    private sealed class CloseReceiptWorker(ServiceResponse response, WorkerSnapshot snapshot) : IExcelWorkerRpc
    {
        internal int Calls { get; private set; }
        internal int ReleaseCalls { get; private set; }

        public Task<WorkerExecution> ExecuteAsync(WorkerCall call, CancellationToken cancellationToken = default)
        {
            Calls++;
            return Task.FromResult(new WorkerExecution(response, Accepted: true, snapshot));
        }

        public Task<WorkerSnapshot> ClaimAsync(WorkerClaim claim, CancellationToken cancellationToken = default) =>
            Task.FromResult(snapshot);

        public Task<WorkerSnapshot> SnapshotAsync(string generation, CancellationToken cancellationToken = default) =>
            Task.FromResult(snapshot);

        public Task<ServiceResponse?> ObserveLateCompletionAsync(string generation, string operationId,
            CancellationToken cancellationToken = default) => Task.FromResult<ServiceResponse?>(null);

        public Task<WorkerRelease> ReleaseAsync(string generation, CancellationToken cancellationToken = default)
        {
            ReleaseCalls++;
            return Task.FromResult(new WorkerRelease(new ServiceResponse { Success = true }, snapshot));
        }
    }
}
