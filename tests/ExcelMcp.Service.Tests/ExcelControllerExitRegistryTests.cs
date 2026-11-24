using System.Collections.Concurrent;
using System.ComponentModel;
using System.Diagnostics;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service.Workers;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ExcelControllerTests
{
    [Theory(Timeout = 15000)]
    [InlineData("missing", false, true)]
    [InlineData("missing", true, true)]
    [InlineData("stale-zero", false, true)]
    [InlineData("matching-zero", false, false)]
    [InlineData("previous-zero", false, false)]
    [InlineData("live", false, false)]
    public async Task Controller_SettledMissingSessionChecksExactProcessWhenExitRecordIsUnavailable(
        string evidence, bool releasedIdentity, bool expectedUnconfirmed)
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-missing-exit-record-");
        try
        {
            using var controller = new ExcelController(launcher: Launcher, registry: new WorkerRegistry(directory.FullName));
            await StopNativeObservationPollingAsync(controller);
            var current = WorkerProcess.Current;
            var excel = new ExcelProcessIdentity(current.ProcessId, current.StartedAtUtcFileTime);
            if (evidence != "live")
            {
                using var child = new ProductProcessLauncher(DotNetHost, ["--list-sdks"]).Start([]);
                _ = child.Handle;
                var process = WorkerProcess.From(child);
                excel = new ExcelProcessIdentity(process.ProcessId, process.StartedAtUtcFileTime);
                await child.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(4));
            }

            const string sessionId = "retired-before-exit-record";
            var generation = ExitRegistryControllerGeneration(controller);
            var registration = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", current,
                current, generation, WorkerRegistry.Build, excel);
            var worker = new ExcelWorkerClient(registration, current, generation);
            var workers = ExitRegistryWorkers(controller);
            Assert.True(workers.TryAdd(registration.PipeName, worker));
            try
            {
                var path = Path.Join(directory.FullName, "claimed.xlsx");
                var sessions = Assert.IsType<ConcurrentDictionary<string, ExcelWorkerClient>>(
                    typeof(ExcelController).GetField("_sessions", NativeObservationHidden)!.GetValue(controller));
                Assert.True(sessions.TryAdd(sessionId, worker));
                var apply = typeof(ExcelController).GetMethod("ApplyWorkerSnapshot", NativeObservationHidden)!;
                worker.RecordSnapshot(new(registration, [new SessionDescriptor(sessionId, path)],
                    new ServiceResponse { Success = true, Result = """{"success":true,"sessions":[{"sessionId":"retired-before-exit-record"}],"count":1}""" },
                    null, null, new Dictionary<string, IReadOnlyList<string>> { [sessionId] = [path] }, Revision: 1));
                apply.Invoke(controller, [worker]);

                var observation = evidence switch
                {
                    "stale-zero" => new ExcelNativeProcessObservation(
                        excel with { StartedAtUtcFileTime = excel.StartedAtUtcFileTime - 1 },
                        ExcelNativeProcessState.Exited, 0, null),
                    "matching-zero" => new ExcelNativeProcessObservation(excel, ExcelNativeProcessState.Exited, 0, null),
                    _ => null
                };
                var settled = registration with
                {
                    Excel = releasedIdentity ? null : excel,
                    ReleasedExcel = releasedIdentity ? excel : null,
                    LastExcelObservation = observation
                };
                if (evidence == "previous-zero")
                    PublishNativeObservation(controller, excel, ExcelNativeProcessState.Exited, 0, null);
                worker.RecordSnapshot(new(settled, [],
                    new ServiceResponse { Success = true, Result = """{"success":true,"sessions":[],"count":0}""" },
                    null, null, new Dictionary<string, IReadOnlyList<string>>(), Revision: 2));
                apply.Invoke(controller, [worker]);
                Assert.DoesNotContain(sessionId, sessions.Keys);

                if (releasedIdentity)
                {
                    var history = Assert.IsType<CloseHistory>(typeof(ExcelController)
                        .GetField("_closeHistory", NativeObservationHidden)!.GetValue(controller));
                    history.Record(new(sessionId, excel), new ServiceResponse
                    {
                        ErrorCategory = "Busy",
                        ExceptionType = "ExcelBusyException",
                        HResult = "0x800706BA"
                    });
                }

                var close = await controller.ProcessAsync(new ServiceRequest { Command = "session.close", SessionId = sessionId });
                Assert.Equal(!expectedUnconfirmed, close.Success);
                if (expectedUnconfirmed)
                {
                    Assert.Equal("ServiceUnavailable", close.ErrorCategory);
                    Assert.Contains("Excel process exited", close.ErrorMessage, StringComparison.Ordinal);
                    using var state = JsonDocument.Parse(close.PartialState!);
                    Assert.Equal("unknown", state.RootElement.GetProperty("outcome").GetString());
                    var native = state.RootElement.GetProperty("diagnostics").GetProperty("nativeProcess");
                    Assert.Equal(excel.ProcessId, native.GetProperty("processId").GetInt32());
                    Assert.Equal(excel.StartedAtUtcFileTime, long.Parse(native.GetProperty("startedAtUtcFileTime").GetString()!,
                        System.Globalization.CultureInfo.InvariantCulture));
                    Assert.Equal("Exited", native.GetProperty("state").GetString());
                    Assert.Equal(JsonValueKind.Null, native.GetProperty("exitCode").ValueKind);
                    var list = await controller.ProcessAsync(new ServiceRequest { Command = "session.list" });
                    Assert.True(list.Success, list.ErrorMessage);
                    using var listed = JsonDocument.Parse(list.Result!);
                    var attempt = Assert.Single(listed.RootElement.GetProperty("recentCloseAttempts").EnumerateArray());
                    Assert.Equal(sessionId, attempt.GetProperty("sessionId").GetString());
                    Assert.True(attempt.GetProperty("sessionReleased").GetBoolean());
                    Assert.False(attempt.GetProperty("closeConfirmed").GetBoolean());
                    Assert.False(attempt.GetProperty("saveConfirmed").GetBoolean());
                    Assert.Equal(native.GetRawText(), attempt.GetProperty("nativeExit").GetRawText());
                    var failure = attempt.GetProperty("lastFailure");
                    Assert.Equal(releasedIdentity ? "Busy" : "ServiceUnavailable", failure.GetProperty("category").GetString());
                    if (releasedIdentity)
                        Assert.Equal("0x800706BA", failure.GetProperty("hresult").GetString());

                    var operationId = Guid.NewGuid().ToString("D");
                    var invoke = await controller.ProcessAsync(new ServiceRequest
                    {
                        Command = "vba.invoke",
                        SessionId = sessionId,
                        OperationId = operationId,
                        Args = """{"procedureName":"Probe.Check"}"""
                    });
                    Assert.False(invoke.Success);
                    using var invokeState = JsonDocument.Parse(invoke.PartialState!);
                    Assert.Equal("Probe.Check", invokeState.RootElement.GetProperty("diagnostics")
                        .GetProperty("operationEvidence").GetProperty("requestedProcedure").GetString());
                    var lookup = await controller.ProcessAsync(new ServiceRequest
                    {
                        Command = OperationReceipts.LookupCommand,
                        SessionId = sessionId,
                        Args = JsonSerializer.Serialize(new { operationId })
                    });
                    Assert.True(lookup.Success, lookup.ErrorMessage);
                    using var receipt = JsonDocument.Parse(lookup.Result!);
                    Assert.Equal("Probe.Check", receipt.RootElement.GetProperty("operationEvidence")
                        .GetProperty("requestedProcedure").GetString());
                }
                else
                {
                    using var result = JsonDocument.Parse(close.Result!);
                    Assert.False(result.RootElement.GetProperty("saved").GetBoolean());
                }
            }
            finally
            {
                workers.TryRemove(registration.PipeName, out _);
            }
        }
        finally
        {
            directory.Delete();
        }
    }

    [Theory(Timeout = 15000)]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Registry_DiscoveryRetainsExcelRegisteredDuringTheRead(bool controllerAlive)
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-registration-race-");
        var registry = new WorkerRegistry(directory.FullName);
        using var exited = new ProductProcessLauncher(DotNetHost, ["--list-sdks"]).Start([]);
        var stopped = WorkerProcess.From(exited);
        await exited.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(4));
        var current = WorkerProcess.Current;
        var record = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", stopped,
            controllerAlive ? current : stopped, "generation", WorkerRegistry.Build, null);
        var registered = record with
        {
            Excel = new ExcelProcessIdentity(current.ProcessId, current.StartedAtUtcFileTime),
            OwnsExcel = true
        };
        try
        {
            registry.Write(record);
            registry.AfterRecordOpenForTests = () =>
            {
                registry.AfterRecordOpenForTests = null;
                registry.Write(registered);
            };

            Assert.Equal(registered, Assert.Single(registry.Read()));
            Assert.Equal(registered, registry.Read(record.PipeName));
        }
        finally
        {
            registry.Remove(record.PipeName);
            directory.Delete();
        }
    }

    [Fact(Timeout = 15000)]
    public async Task Registry_DiscoveryDoesNotPruneACompletedRecordWhileItsWorkerCanWrite()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-registration-release-");
        var registry = new WorkerRegistry(directory.FullName);
        using var exited = new ProductProcessLauncher(DotNetHost, ["--list-sdks"]).Start([]);
        var stopped = WorkerProcess.From(exited);
        await exited.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(4));
        var record = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", WorkerProcess.Current,
            stopped, "generation", WorkerRegistry.Build, null, ReleaseConfirmed: true);
        try
        {
            registry.Write(record);
            Assert.Empty(registry.Read());
            Assert.Equal(record, registry.Read(record.PipeName));
        }
        finally
        {
            registry.Remove(record.PipeName);
            directory.Delete();
        }
    }

    [Theory(Timeout = 15000)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    public async Task Registry_DiscoveryLeavesExitForLiveOwnerToAcknowledge(bool controllerAlive, bool replaceDuringRead)
    {
        var directory = Path.Combine(Path.GetTempPath(), $"excelmcp-exit-discovery-{Guid.NewGuid():N}");
        var registry = new WorkerRegistry(directory);
        using var exited = new ProductProcessLauncher(DotNetHost, ["--list-sdks"]).Start([]);
        var stopped = WorkerProcess.From(exited);
        await exited.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(4));
        var excel = new ExcelProcessIdentity(stopped.ProcessId, stopped.StartedAtUtcFileTime);
        var record = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", stopped,
            controllerAlive ? WorkerProcess.Current : stopped, "generation", WorkerRegistry.Build, excel,
            LastExcelObservation: new(excel, ExcelNativeProcessState.Exited, 0, null));
        try
        {
            registry.Write(replaceDuringRead ? record with { LastExcelObservation = null } : record);
            if (replaceDuringRead)
                registry.AfterRecordOpenForTests = () =>
                {
                    registry.AfterRecordOpenForTests = null;
                    registry.Write(record);
                };
            Assert.Empty(registry.Read());
            Assert.Equal(controllerAlive ? record : null, registry.Read(record.PipeName));
        }
        finally
        {
            registry.Remove(record.PipeName);
            Directory.Delete(directory);
        }
    }

    [Theory(Timeout = 15000)]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Controller_PollRecoversExitWhenWorkerDiesDuringSnapshotRead(bool confirmedRelease)
    {
        var directory = Path.Combine(Path.GetTempPath(), $"excelmcp-exit-snapshot-{Guid.NewGuid():N}");
        var registry = new WorkerRegistry(directory);
        try
        {
            using var controller = new ExcelController(launcher: Launcher, registry: registry);
            using var exited = new ProductProcessLauncher(DotNetHost, ["--list-sdks"]).Start([]);
            var excelProcess = WorkerProcess.From(exited);
            await exited.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(4));
            var excel = new ExcelProcessIdentity(excelProcess.ProcessId, excelProcess.StartedAtUtcFileTime);
            using var child = Process.Start(new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "cmd.exe"),
                "/d /q /c set /p \"unused=\"")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardInput = true
            })!;
            var generation = ExitRegistryControllerGeneration(controller);
            var registration = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", WorkerProcess.From(child),
                WorkerProcess.Current, generation, WorkerRegistry.Build, excel, OwnsExcel: true);
            var worker = new ExcelWorkerClient(registration, registration.Controller, generation);
            var workers = ExitRegistryWorkers(controller);
            var observation = new ExcelNativeProcessObservation(excel, ExcelNativeProcessState.Exited, 0, null);
            var receipt = registration with { LastExcelObservation = observation };
            if (confirmedRelease)
                receipt = receipt with { ReleaseConfirmed = true, Excel = null, ReleasedExcel = excel, OwnsExcel = false };
            controller.BeforeWorkerSnapshotReadForTests = async candidate =>
            {
                if (candidate != worker || child.HasExited)
                    return;
                registry.Write(receipt);
                child.StandardInput.Close();
                await child.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(4));
            };
            try
            {
                Assert.True(workers.TryAdd(registration.PipeName, worker));
                await WaitForExitRegistryPollAsync(controller, () => !workers.ContainsKey(registration.PipeName));
                Assert.Null(registry.Read(registration.PipeName));
                using var result = NativeObservationSessionList(controller);
                var observed = Assert.Single(result.RootElement.GetProperty("recentExcelExits").EnumerateArray());
                Assert.Equal(excel.ProcessId, observed.GetProperty("processId").GetInt32());
                Assert.Equal(0u, observed.GetProperty("exitCode").GetUInt32());
            }
            finally
            {
                workers.TryRemove(registration.PipeName, out _);
                child.StandardInput.Close();
                await child.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(4));
            }
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact(Timeout = 15000)]
    public async Task Controller_PollContinuesForOtherWorkersWhenAProcessProbeIsDenied()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-poll-denied-");
        var registry = new WorkerRegistry(directory.FullName);
        try
        {
            using var controller = new ExcelController(launcher: Launcher, registry: registry);
            var healthy = await controller.StartWorkerAsync(null, CancellationToken.None);
            using var child = Process.Start(new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "cmd.exe"),
                "/d /q /c set /p \"unused=\"")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardInput = true
            })!;
            var generation = ExitRegistryControllerGeneration(controller);
            var registration = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", WorkerProcess.From(child),
                WorkerProcess.Current, generation, WorkerRegistry.Build, null);
            var denied = new ExcelWorkerClient(registration, registration.Controller, generation);
            var workers = ExitRegistryWorkers(controller);
            var healthyReads = 0;
            controller.BeforeWorkerSnapshotReadForTests = candidate =>
            {
                if (candidate == healthy)
                    Interlocked.Increment(ref healthyReads);
                return Task.CompletedTask;
            };
            registry.Write(registration);
            WorkerProcess.BeforeIdentityReadForTests = process =>
            {
                if (process.Id == child.Id)
                    throw new Win32Exception(5);
            };
            try
            {
                Assert.True(workers.TryAdd(registration.PipeName, denied));
                await WaitForExitRegistryPollAsync(controller, () =>
                    denied.StatusError is not null && Volatile.Read(ref healthyReads) >= 3);
                Assert.Same(denied, workers[registration.PipeName]);
                Assert.Equal(registration, registry.Read(registration.PipeName));
                Assert.Null(healthy.StatusError);
            }
            finally
            {
                WorkerProcess.BeforeIdentityReadForTests = null;
                workers.TryRemove(registration.PipeName, out _);
                child.StandardInput.Close();
                await child.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(4));
                registry.Remove(registration.PipeName);
            }
        }
        finally
        {
            directory.Delete(recursive: true);
        }
    }

    [Fact(Timeout = 15000)]
    public async Task Controller_PollRetriesLockedExitedWorkerReceipt()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"excelmcp-exit-retry-{Guid.NewGuid():N}");
        var registry = new WorkerRegistry(directory);
        var controller = new ExcelController(launcher: Launcher, registry: registry);
        try
        {
            using var exited = new ProductProcessLauncher(DotNetHost, ["--list-sdks"]).Start([]);
            var stopped = WorkerProcess.From(exited);
            await exited.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(4));
            var excel = new ExcelProcessIdentity(stopped.ProcessId, stopped.StartedAtUtcFileTime);
            var generation = ExitRegistryControllerGeneration(controller);
            var registration = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", stopped,
                WorkerProcess.Current, generation, WorkerRegistry.Build, excel, OwnsExcel: true);
            var worker = new ExcelWorkerClient(registration, registration.Controller, generation);
            var workers = ExitRegistryWorkers(controller);
            registry.Write(registration with
            {
                LastExcelObservation = new(excel, ExcelNativeProcessState.Exited, 0, null)
            });
            var path = Path.Combine(directory, registration.PipeName + ".json");
            using (new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                Assert.True(workers.TryAdd(registration.PipeName, worker));
                await WaitForExitRegistryPollAsync(controller, () =>
                    !workers.ContainsKey(registration.PipeName)
                    || worker.StatusError?.Contains("record could not be removed", StringComparison.Ordinal) == true);
                Assert.Same(worker, workers[registration.PipeName]);
                Assert.True(File.Exists(path));
            }
            await WaitForExitRegistryPollAsync(controller, () => !workers.ContainsKey(registration.PipeName));
            Assert.False(File.Exists(path));
            using var result = NativeObservationSessionList(controller);
            Assert.Single(result.RootElement.GetProperty("recentExcelExits").EnumerateArray());
        }
        finally
        {
            try
            {
                controller.Dispose();
            }
            finally
            {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    private static ConcurrentDictionary<string, ExcelWorkerClient> ExitRegistryWorkers(ExcelController controller) =>
        Assert.IsType<ConcurrentDictionary<string, ExcelWorkerClient>>(typeof(ExcelController)
            .GetField("_workers", NativeObservationHidden)!.GetValue(controller));

    private static string ExitRegistryControllerGeneration(ExcelController controller) =>
        Assert.IsType<string>(typeof(ExcelController)
            .GetField("_generation", NativeObservationHidden)!.GetValue(controller));

    private static async Task WaitForExitRegistryPollAsync(ExcelController controller, Func<bool> complete)
    {
        var poll = Assert.IsAssignableFrom<Task>(typeof(ExcelController)
            .GetField("_poll", NativeObservationHidden)!.GetValue(controller));
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(7));
        while (!complete())
        {
            Assert.False(poll.IsCompleted, poll.Exception?.ToString());
            await Task.Delay(25, timeout.Token);
        }
    }
}
