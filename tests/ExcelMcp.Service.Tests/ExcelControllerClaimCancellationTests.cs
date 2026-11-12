using System.Collections.Concurrent;
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
    public async Task Controller_CancelledDiscoveryTracksOnlyAClaimThatCouldHaveBeenSent(bool afterCommit)
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-claim-cancellation-");
        var registry = new WorkerRegistry(directory.FullName);
        var current = WorkerProcess.Current;
        var previousController = current with { StartedAtUtcFileTime = current.StartedAtUtcFileTime - 1 };
        var excel = new ExcelProcessIdentity(current.ProcessId, current.StartedAtUtcFileTime);
        var record = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", current, previousController,
            "previous", WorkerRegistry.Build, excel, OwnsExcel: true);
        registry.Write(record);

        using var controller = new ExcelController(launcher: Launcher, registry: registry);
        await StopNativeObservationPollingAsync(controller);
        using var cancellation = new CancellationTokenSource();
        var target = new DeferredClaimWorker(record, registry);
        using var host = new DaemonHost(record.PipeName, pipe => new JsonRpc(pipe, pipe, target)
        {
            CancelLocallyInvokedMethodsWhenConnectionIsClosed = true
        }, maximumConnections: 8);
        Task? listener = afterCommit ? host.RunAsync(record.PipeName) : null;
        RunningExcelWindows.FindHookForTests = () => new RunningExcelWindows.Discovery(
            [new RunningExcelWindows.ApplicationWindow(nint.Zero, excel)], []);
        try
        {
            var discovery = controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.discover",
                CallerCancellation = cancellation.Token
            });
            // Without a listening pipe, the call has reached the cancellable connection wait.
            Assert.False(discovery.IsCompleted);
            if (afterCommit)
                await target.Committed.Task.WaitAsync(TimeSpan.FromSeconds(3));
            cancellation.Cancel();
            if (afterCommit)
                await Task.WhenAny(discovery, Task.Delay(100));
            target.Finish.TrySetResult();

            var result = await discovery.WaitAsync(TimeSpan.FromSeconds(3));
            Assert.False(result.Success);
            Assert.Equal("Cancelled", result.ErrorCategory);
            Assert.Equal(0, target.DiscoverCalls);
            Assert.Equal(0, target.ReleaseCalls);
            var workers = Assert.IsType<ConcurrentDictionary<string, ExcelWorkerClient>>(
                typeof(ExcelController).GetField("_workers", NativeObservationHidden)!.GetValue(controller));
            if (afterCommit)
            {
                var claimed = Assert.Single(workers).Value;
                var snapshot = Assert.IsType<WorkerSnapshot>(claimed.Snapshot);
                Assert.Equal(snapshot.Registration, registry.Read(record.PipeName));
                Assert.Equal(current, snapshot.Registration.Controller);
                Assert.Equal(1, target.ClaimCalls);
            }
            else
            {
                Assert.Empty(workers);
                Assert.Equal(record, registry.Read(record.PipeName));
                Assert.Equal(0, target.ClaimCalls);
            }
        }
        finally
        {
            RunningExcelWindows.FindHookForTests = null;
            target.Finish.TrySetResult();
            // The fixture owns no Excel objects. Keep shutdown out of the claim assertions.
            await controller.ShutdownAsync();
            host.RequestShutdown();
            if (listener is not null)
                await listener.WaitAsync(TimeSpan.FromSeconds(3));
            registry.Remove(record.PipeName);
            directory.Delete();
        }
    }

    private sealed class DeferredClaimWorker(WorkerRegistration record, WorkerRegistry registry) : IExcelWorkerRpc
    {
        internal TaskCompletionSource Committed { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource Finish { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal int ClaimCalls { get; private set; }
        internal int DiscoverCalls { get; private set; }
        internal int ReleaseCalls { get; private set; }

        public async Task<WorkerSnapshot> ClaimAsync(WorkerClaim claim, CancellationToken cancellationToken = default)
        {
            ClaimCalls++;
            record = record with { Controller = claim.Controller, Generation = claim.Generation };
            registry.Write(record);
            Committed.TrySetResult();
            await Finish.Task;
            return Snapshot();
        }

        public Task<WorkerSnapshot> SnapshotAsync(string generation, CancellationToken cancellationToken = default) =>
            generation == record.Generation
                ? Task.FromResult(Snapshot())
                : throw new InvalidOperationException("This request is not from the owning Excel controller.");

        public Task<WorkerExecution> ExecuteAsync(WorkerCall call, CancellationToken cancellationToken = default)
        {
            DiscoverCalls++;
            return Task.FromResult(new WorkerExecution(new ServiceResponse { Success = false }, false, Snapshot()));
        }

        public Task<WorkerRelease> ReleaseAsync(string generation, CancellationToken cancellationToken = default)
        {
            ReleaseCalls++;
            return Task.FromResult(new WorkerRelease(new ServiceResponse { Success = false }, Snapshot()));
        }

        private WorkerSnapshot Snapshot() => new(record, [],
            new ServiceResponse { Success = true, Result = """{"success":true,"sessions":[],"count":0}""" },
            null, null, new Dictionary<string, IReadOnlyList<string>>(), Revision: 1);
    }
}
