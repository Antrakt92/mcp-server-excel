using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service.Workers;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ExcelControllerTests
{
    [Theory(Timeout = 15000)]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Controller_PollRetiresDeadUtilityOnlyWhenNoApplicationRemains(bool retainedLiveApplication)
    {
        var directory = Path.Combine(Path.GetTempPath(), $"excelmcp-dead-utility-{Guid.NewGuid():N}");
        var registry = new WorkerRegistry(directory);
        try
        {
            using var controller = new ExcelController(launcher: Launcher, registry: registry);
            using var exited = new ProductProcessLauncher(DotNetHost, ["--list-sdks"]).Start([]);
            var stopped = WorkerProcess.From(exited);
            await exited.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(4));
            var generation = ExitRegistryControllerGeneration(controller);
            var owner = WorkerProcess.Current;
            var registration = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", stopped,
                owner, generation, WorkerRegistry.Build, Excel: null);
            var worker = new ExcelWorkerClient(registration, owner, generation);
            worker.RecordSnapshot(new(registration, [],
                new ServiceResponse { Success = true, Result = """{"success":true,"sessions":[],"count":0}""" },
                null, null, new Dictionary<string, IReadOnlyList<string>>(), Revision: 1));
            // Model an application identity durably published immediately before the worker exited.
            var recorded = retainedLiveApplication
                ? registration with { Excel = new ExcelProcessIdentity(owner.ProcessId, owner.StartedAtUtcFileTime), OwnsExcel = true }
                : registration;
            registry.Write(recorded);
            var utility = typeof(ExcelController).GetField("_utility", NativeObservationHidden)!;
            utility.SetValue(controller, Task.FromResult(worker));
            var workers = ExitRegistryWorkers(controller);
            var statusReads = 0;
            registry.BeforeRecordReadForTests = _ => Interlocked.Increment(ref statusReads);
            Assert.True(workers.TryAdd(registration.PipeName, worker));
            if (retainedLiveApplication)
            {
                // A second polling pass ensures the first retirement attempt has already finished.
                await WaitForExitRegistryPollAsync(controller, () => Volatile.Read(ref statusReads) >= 3);
                Assert.True(workers.ContainsKey(registration.PipeName));
                Assert.Equal(recorded, registry.Read(registration.PipeName));
            }
            else
            {
                await WaitForExitRegistryPollAsync(controller, () => utility.GetValue(controller) is null);
                Assert.False(workers.ContainsKey(registration.PipeName));
                Assert.Null(registry.Read(registration.PipeName));
                using var result = NativeObservationSessionList(controller);
                Assert.Equal(0, result.RootElement.GetProperty("count").GetInt32());
            }
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }
}
