using System.Collections.Concurrent;
using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service.Workers;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ExcelControllerTests
{
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
