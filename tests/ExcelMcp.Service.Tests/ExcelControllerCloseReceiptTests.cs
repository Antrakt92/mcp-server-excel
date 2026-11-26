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
                        await ExerciseClosedSessionRetentionAsync(controller, registration, settled);
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

    private static async Task ExerciseClosedSessionRetentionAsync(ExcelController controller,
        WorkerRegistration registration, WorkerSnapshot template)
    {
        var other = registration with { PipeName = $"excelmcp-worker-{Guid.NewGuid():N}" };
        var worker = new ExcelWorkerClient(other, other.Controller, other.Generation);
        var workers = Assert.IsType<System.Collections.Concurrent.ConcurrentDictionary<string, ExcelWorkerClient>>(
            typeof(ExcelController).GetField("_workers", NativeObservationHidden)!.GetValue(controller));
        var sessions = Assert.IsType<System.Collections.Concurrent.ConcurrentDictionary<string, ExcelWorkerClient>>(
            typeof(ExcelController).GetField("_sessions", NativeObservationHidden)!.GetValue(controller));
        var published = Assert.IsType<System.Collections.Concurrent.ConcurrentDictionary<string, WorkerSnapshot>>(
            typeof(ExcelController).GetField("_publishedSnapshots", NativeObservationHidden)!.GetValue(controller));
        var history = Assert.IsType<CloseHistory>(
            typeof(ExcelController).GetField("_closeHistory", NativeObservationHidden)!.GetValue(controller));
        var apply = typeof(ExcelController).GetMethod("ApplyWorkerSnapshot", NativeObservationHidden)!;
        Assert.True(workers.TryAdd(other.PipeName, worker));
        try
        {
            var empty = template with { Registration = other };
            void ObserveCloses(string[] ids, long revision)
            {
                foreach (var id in ids)
                    Assert.True(sessions.TryAdd(id, worker));
                worker.RecordSnapshot(empty with
                {
                    Sessions = ids.Select(id => new SessionDescriptor(id, null)).ToArray(),
                    OwnershipKeys = ids.ToDictionary(id => id, _ => (IReadOnlyList<string>)[]),
                    SessionList = new ServiceResponse
                    {
                        Success = true,
                        Result = JsonSerializer.Serialize(new { sessions = ids.Select(sessionId => new { sessionId }) })
                    },
                    Revision = revision
                });
                apply.Invoke(controller, [worker]);
                worker.RecordSnapshot(empty with { Revision = revision + 1 });
                apply.Invoke(controller, [worker]);
                Assert.All(ids, id => Assert.False(sessions.ContainsKey(id)));
            }

            const string oldest = "oldest-closed-session";
            ObserveCloses([oldest], 1);
            var retained = Enumerable.Range(0, ExcelController.SettledCloseRetention + 16)
                .Select(i => $"settled-closed-session-{i}").ToArray();
            // Turn over the diagnostic history as well so the early save
            // receipt survives only if its unsettled worker protects it.
            foreach (var id in retained.Take(16))
                history.Record(new(id, null), new ServiceResponse { Success = true });
            ObserveCloses(retained, 3);

            var expired = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = oldest,
                Args = """{"save":true}"""
            });
            Assert.False(expired.Success);
            Assert.Equal("SessionNotFound", expired.ErrorCategory);
            Assert.Contains("No close or save was started", expired.ErrorMessage, StringComparison.Ordinal);
            Assert.Null(typeof(ExcelController).GetField("_utility", NativeObservationHidden)!.GetValue(controller));

            // Simulate a retry recording its lost reply after the entry has expired.
            typeof(ExcelController).GetMethod("RecordLostSessionCloseAttempt", NativeObservationHidden)!
                .Invoke(controller, [oldest, null, new ServiceResponse { Success = false, ErrorCategory = "WorkerUnavailable" }]);
            Assert.DoesNotContain(oldest, history.RetainedSessionIds());

            var recent = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = retained[0]
            });
            Assert.True(recent.Success, recent.ErrorMessage);
        }
        finally
        {
            workers.TryRemove(other.PipeName, out _);
            published.TryRemove(other.PipeName, out _);
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

        public Task<WorkerRelease> ReleaseAsync(string generation, bool onlyIfEmpty, CancellationToken cancellationToken = default)
        {
            ReleaseCalls++;
            return Task.FromResult(new WorkerRelease(new ServiceResponse { Success = true }, snapshot));
        }
    }
}
