using System.Collections.Concurrent;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service.Workers;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ExcelControllerTests
{
    [Theory(Timeout = 30000)]
    [InlineData("released-exited", true, true)]
    [InlineData("released-live", true, false)]
    [InlineData("released-live", false, true)]
    [InlineData("no-identity", true, false)]
    [InlineData("active-live", true, false)]
    public async Task Controller_ExitedWorkerReconcilesStaleRowsOnlyAfterExactExcelExit(
        string scenario, bool staleSession, bool shouldRetire)
    {
        var directory = Path.Combine(Path.GetTempPath(), $"excelmcp-stale-release-{Guid.NewGuid():N}");
        try
        {
            var registry = new WorkerRegistry(directory);
            using var controller = new ExcelController(launcher: Launcher, registry: registry);
            await StopNativeObservationPollingAsync(controller);
            using var exited = new ProductProcessLauncher(DotNetHost, ["--list-sdks"]).Start([]);
            _ = exited.Handle;
            var process = WorkerProcess.From(exited);
            await exited.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
            Assert.Equal(0, exited.ExitCode);

            var type = typeof(ExcelController);
            var generation = Assert.IsType<string>(type.GetField("_generation", NativeObservationHidden)!.GetValue(controller));
            var current = WorkerProcess.Current;
            var nativeExit = new ExcelProcessIdentity(process.ProcessId, process.StartedAtUtcFileTime);
            var live = new ExcelProcessIdentity(current.ProcessId, current.StartedAtUtcFileTime);
            var registration = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", process, current, generation,
                WorkerRegistry.Build, scenario == "active-live" ? live : null, OwnsExcel: scenario == "active-live",
                ReleasedExcel: scenario switch
                {
                    "released-exited" or "active-live" => nativeExit,
                    "released-live" => live,
                    _ => null
                });
            var worker = new ExcelWorkerClient(registration, current, generation);
            var workers = Assert.IsType<ConcurrentDictionary<string, ExcelWorkerClient>>(
                type.GetField("_workers", NativeObservationHidden)!.GetValue(controller));
            var sessions = Assert.IsType<ConcurrentDictionary<string, ExcelWorkerClient>>(
                type.GetField("_sessions", NativeObservationHidden)!.GetValue(controller));
            var claims = Assert.IsAssignableFrom<System.Collections.IEnumerable>(
                type.GetField("_paths", NativeObservationHidden)!.GetValue(controller));
            var sessionId = Guid.NewGuid().ToString("N");
            var filePath = Path.Combine(directory, "retained.xlsx");
            SessionDescriptor[] snapshotRows = staleSession ? [new(sessionId, filePath)] : [];
            worker.RecordSnapshot(new(registration, snapshotRows,
                new ServiceResponse
                {
                    Success = true,
                    Result = JsonSerializer.Serialize(new { sessions = snapshotRows }, ServiceProtocol.JsonOptions)
                }, null, null,
                staleSession ? new Dictionary<string, IReadOnlyList<string>> { [sessionId] = [filePath] }
                    : new Dictionary<string, IReadOnlyList<string>>(), Revision: 1));
            Assert.True(workers.TryAdd(registration.PipeName, worker));
            try
            {
                if (staleSession)
                    Assert.True(sessions.TryAdd(sessionId, worker));
                type.GetMethod("ApplyWorkerSnapshot", NativeObservationHidden)!.Invoke(controller, [worker]);
                Assert.Equal(staleSession, claims.Cast<object>().Any());
                registry.Write(scenario == "released-exited"
                    ? registration with { LastExcelObservation = new(nativeExit, ExcelNativeProcessState.Exited, 0, null) }
                    : registration);

                type.GetMethod("RetireExitedWorker", NativeObservationHidden)!.Invoke(controller, [registration, false]);

                Assert.Equal(!shouldRetire, workers.ContainsKey(registration.PipeName));
                Assert.Equal(!shouldRetire, registry.Read(registration.PipeName) is not null);
                Assert.Equal(staleSession && !shouldRetire, sessions.ContainsKey(sessionId));
                Assert.Equal(staleSession && !shouldRetire, claims.Cast<object>().Any());
                if (scenario == "released-exited")
                {
                    var retry = await controller.ProcessAsync(new ServiceRequest { Command = "session.close", SessionId = sessionId });
                    Assert.False(retry.Success);
                    using var partial = JsonDocument.Parse(retry.PartialState!);
                    Assert.Equal("unknown", partial.RootElement.GetProperty("outcome").GetString());
                    var exit = partial.RootElement.GetProperty("diagnostics").GetProperty("nativeProcess");
                    Assert.Equal(nativeExit.ProcessId, exit.GetProperty("processId").GetInt32());
                    Assert.Equal(nativeExit.StartedAtUtcFileTime.ToString(System.Globalization.CultureInfo.InvariantCulture),
                        exit.GetProperty("startedAtUtcFileTime").GetString());
                    Assert.Equal(0u, exit.GetProperty("exitCode").GetUInt32());
                }
            }
            finally
            {
                workers.TryRemove(registration.PipeName, out _);
            }
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, recursive: true);
        }
    }

    [Theory(Timeout = 30000)]
    [InlineData("matching")]
    [InlineData("released")]
    [InlineData("worker-creation")]
    [InlineData("controller-creation")]
    [InlineData("generation")]
    [InlineData("build")]
    [InlineData("pipe")]
    [InlineData("excel-creation")]
    [InlineData("observation-creation")]
    [InlineData("alive")]
    [InlineData("confirmed-forged-excel")]
    public async Task Controller_DeadWorkerRecoversOnlyMatchingDurableExitBeforeRetirement(string scenario)
    {
        var directory = Path.Combine(Path.GetTempPath(), $"excelmcp-durable-exit-{Guid.NewGuid():N}");
        try
        {
            var registry = new WorkerRegistry(directory);
            using var controller = new ExcelController(launcher: Launcher, registry: registry);
            await StopNativeObservationPollingAsync(controller);
            using var exited = new ProductProcessLauncher(DotNetHost, ["--list-sdks"]).Start([]);
            _ = exited.Handle;
            var process = WorkerProcess.From(exited);
            await exited.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
            Assert.Equal(0, exited.ExitCode);
            var identity = new ExcelProcessIdentity(process.ProcessId, process.StartedAtUtcFileTime);
            var generation = Assert.IsType<string>(typeof(ExcelController)
                .GetField("_generation", NativeObservationHidden)!.GetValue(controller));
            var current = WorkerProcess.Current;
            var registration = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", process, current,
                generation, WorkerRegistry.Build, identity, OwnsExcel: true);
            var worker = new ExcelWorkerClient(registration, current, generation);
            worker.RecordSnapshot(new(registration, [],
                new ServiceResponse { Success = true, Result = """{"success":true,"sessions":[],"count":0}""" },
                null, null, new Dictionary<string, IReadOnlyList<string>>(), Revision: 1));
            var workers = Assert.IsType<ConcurrentDictionary<string, ExcelWorkerClient>>(
                typeof(ExcelController).GetField("_workers", NativeObservationHidden)!.GetValue(controller));
            Assert.True(workers.TryAdd(registration.PipeName, worker));
            var recorded = registration with
            {
                LastExcelObservation = new(identity, ExcelNativeProcessState.Exited, 0, null)
            };
            recorded = scenario switch
            {
                "released" => recorded with { Excel = null, OwnsExcel = false, ReleasedExcel = identity },
                "worker-creation" => recorded with { Worker = process with { StartedAtUtcFileTime = process.StartedAtUtcFileTime + 1 } },
                "controller-creation" => recorded with { Controller = current with { StartedAtUtcFileTime = current.StartedAtUtcFileTime + 1 } },
                "generation" => recorded with { Generation = "stale-generation" },
                "build" => recorded with { Build = "stale-build" },
                "pipe" => recorded with { PipeName = $"excelmcp-worker-{Guid.NewGuid():N}" },
                "excel-creation" => recorded with
                {
                    Excel = identity with { StartedAtUtcFileTime = identity.StartedAtUtcFileTime + 1 },
                    LastExcelObservation = recorded.LastExcelObservation! with
                    { Identity = identity with { StartedAtUtcFileTime = identity.StartedAtUtcFileTime + 1 } }
                },
                "observation-creation" => recorded with
                {
                    LastExcelObservation = recorded.LastExcelObservation! with
                    { Identity = identity with { StartedAtUtcFileTime = identity.StartedAtUtcFileTime + 1 } }
                },
                "confirmed-forged-excel" => recorded with
                {
                    Excel = null,
                    OwnsExcel = false,
                    ReleaseConfirmed = true,
                    ReleasedExcel = identity with { StartedAtUtcFileTime = identity.StartedAtUtcFileTime + 1 },
                    LastExcelObservation = recorded.LastExcelObservation! with
                    { Identity = identity with { StartedAtUtcFileTime = identity.StartedAtUtcFileTime + 1 } }
                },
                "alive" => recorded with
                { LastExcelObservation = new(identity, ExcelNativeProcessState.Alive, null, null) },
                _ => recorded
            };
            // Model a durable newer record not delivered over RPC before the worker exits.
            var path = Path.Combine(directory, registration.PipeName + ".json");
            string recordJson = JsonSerializer.Serialize(recorded, ServiceProtocol.JsonOptions);
            File.WriteAllText(path, recordJson);
            Assert.Equal(scenario == "confirmed-forged-excel", recorded.ReleaseConfirmed);
            var ownershipEvents = new List<string>();
            controller.ExcelProcessObserved += _ => ownershipEvents.Add("observed");
            controller.ExcelProcessReleased += _ => ownershipEvents.Add("released");
            try
            {
                var refresh = (Task<WorkerSnapshot?>)typeof(ExcelController)
                    .GetMethod("RefreshWorkerAsync", NativeObservationHidden)!
                    .Invoke(controller, [worker, CancellationToken.None])!;
                Assert.Null(await refresh);
                Assert.NotNull(worker.StatusError);
                Assert.Equal(registration, worker.Registration);
                Assert.Empty(ownershipEvents);
                typeof(ExcelController).GetMethod("RetireExitedWorker", NativeObservationHidden)!
                    .Invoke(controller, [registration, false]);
                if (scenario == "pipe")
                {
                    Assert.True(workers.TryGetValue(registration.PipeName, out var retained));
                    Assert.Same(worker, retained);
                    Assert.Equal(registration, worker.Registration);
                    Assert.Contains("ownership is unconfirmed", worker.StatusError, StringComparison.Ordinal);
                    Assert.Equal(recordJson, File.ReadAllText(path));
                    Assert.Empty(ownershipEvents);
                    return;
                }
                Assert.False(workers.ContainsKey(registration.PipeName));
                registry.Remove(registration.PipeName);

                using var result = NativeObservationSessionList(controller);
                var root = result.RootElement;
                Assert.Equal(0, root.GetProperty("count").GetInt32());
                Assert.Empty(root.GetProperty("sessions").EnumerateArray());
                var history = root.GetProperty("recentExcelExits").EnumerateArray().ToArray();
                if (scenario is "matching" or "released")
                {
                    var observation = Assert.Single(history);
                    Assert.Equal(identity.ProcessId, observation.GetProperty("processId").GetInt32());
                    Assert.Equal(identity.StartedAtUtcFileTime.ToString(System.Globalization.CultureInfo.InvariantCulture),
                        observation.GetProperty("startedAtUtcFileTime").GetString());
                    Assert.Equal("Exited", observation.GetProperty("state").GetString());
                    Assert.Equal(0u, observation.GetProperty("exitCode").GetUInt32());
                    Assert.Equal(JsonValueKind.Null, observation.GetProperty("nativeError").ValueKind);
                }
                else
                    Assert.Empty(history);
                Assert.Empty(ownershipEvents);
            }
            finally
            {
                workers.TryRemove(registration.PipeName, out _);
            }
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, recursive: true);
        }
    }
}
