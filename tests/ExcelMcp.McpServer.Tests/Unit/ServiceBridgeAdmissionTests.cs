using System.Collections.Concurrent;
using System.Text.Json;
using Sbroenne.ExcelMcp.McpServer.ServiceBridge;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Bridge = Sbroenne.ExcelMcp.McpServer.ServiceBridge.ServiceBridge;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "ServiceBridge")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class ServiceBridgeAdmissionTests
{
    [Fact]
    public async Task SaturationRejectsWithoutSerializationAndCancellationKeepsPendingWorkOwned()
    {
        var backend = new Backend();
        using var bridge = new Bridge(() => backend, maxOutstandingDispatches: 2);
        using var cancellation = new CancellationTokenSource();
        var first = bridge.SendAsync("sheet.list", "first", cancellationToken: cancellation.Token);
        var second = bridge.SendAsync("sheet.list", "second");
        try
        {
            await Started(backend);
            await Started(backend);
            var busy = await bridge.SendAsync("sheet.list", "rejected", new Poison());
            Assert.False(busy.Success);
            Assert.Equal("ServiceBusy", busy.ErrorCategory);
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => first);
            Assert.True(backend.Requests["first"].CallerCancellation.IsCancellationRequested);
            Assert.Equal(["first"], backend.Revealed.ToArray());
            Assert.False(backend.Pending["first"].Task.IsCompleted);
            Assert.Equal("ServiceBusy", (await bridge.SendAsync("sheet.list", "still-rejected")).ErrorCategory);
            Assert.Equal(2, backend.Calls);
            Assert.Equal(0, backend.ShutdownCalls);
            Assert.False(backend.Disposed);
            backend.Complete("first");
            backend.Fast = true;
            Assert.True((await WhenAdmitted(bridge)).Success);
            Assert.Equal(3, backend.Calls);
        }
        finally
        {
            backend.Complete("first");
            backend.Complete("second");
            await second.WaitAsync(TimeSpan.FromSeconds(3));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BlockedRevealRetainsCapacityAfterDispatchCompletes(bool lateCreation)
    {
        var backend = new Backend { BlockReveal = true, CompleteOnReveal = !lateCreation };
        using var bridge = new Bridge(() => backend, maxOutstandingDispatches: 2);
        using var firstCancellation = new CancellationTokenSource();
        using var secondCancellation = new CancellationTokenSource();
        var first = bridge.SendAsync(lateCreation ? "session.create" : "sheet.list",
            lateCreation ? null : "first", cancellationToken: firstCancellation.Token);
        Task<ServiceResponse>? second = null;
        Task? cancelFirst = null;
        Task? cancelSecond = null;
        try
        {
            await Started(backend);
            if (lateCreation)
            {
                firstCancellation.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => first);
            }
            else
            {
                cancelFirst = Task.Run(firstCancellation.Cancel);
                Assert.True(await backend.RevealEntered.WaitAsync(TimeSpan.FromSeconds(3)));
            }
            second = bridge.SendAsync(lateCreation ? "session.create" : "sheet.list",
                lateCreation ? null : "second", cancellationToken: secondCancellation.Token);
            await Started(backend);
            if (lateCreation)
            {
                secondCancellation.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => second);
                backend.CompleteCreation("new1");
                backend.CompleteCreation("new2");
                Assert.True(await backend.RevealEntered.WaitAsync(TimeSpan.FromSeconds(3)));
                Assert.True(await backend.RevealEntered.WaitAsync(TimeSpan.FromSeconds(3)));
            }
            else
            {
                cancelSecond = Task.Run(secondCancellation.Cancel);
                Assert.True(await backend.RevealEntered.WaitAsync(TimeSpan.FromSeconds(3)));
            }
            await Task.Delay(100);
            var busy = await bridge.SendAsync("sheet.list", "third").WaitAsync(TimeSpan.FromSeconds(1));
            Assert.Equal("ServiceBusy", busy.ErrorCategory);
            Assert.Equal(2, backend.Calls);
            Assert.Equal(2, backend.Revealed.Count);
            Assert.Equal(0, backend.ShutdownCalls);
            Assert.False(backend.Disposed);
        }
        finally
        {
            backend.ReleaseReveal.Set();
            backend.Complete("first");
            backend.Complete("second");
            backend.CompleteCreation("new1");
            backend.CompleteCreation("new2");
            if (cancelFirst is not null) await cancelFirst.WaitAsync(TimeSpan.FromSeconds(3));
            if (cancelSecond is not null) await cancelSecond.WaitAsync(TimeSpan.FromSeconds(3));
            if (firstCancellation.IsCancellationRequested)
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => first);
            if (second is not null && secondCancellation.IsCancellationRequested)
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => second);
        }
        backend.Fast = true;
        Assert.True((await WhenAdmitted(bridge)).Success);
        Assert.Equal(3, backend.Calls);
    }

    [Fact]
    public async Task CanceledFactoryWaitNeverDispatchesOrRevealsAnExistingSession()
    {
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var backend = new Backend { Fast = true };
        using var bridge = new Bridge(() =>
        {
            entered.Set();
            if (!release.Wait(TimeSpan.FromSeconds(5))) throw new TimeoutException("Factory not released.");
            return backend;
        }, maxOutstandingDispatches: 1);
        using var cancel = new CancellationTokenSource();
        var request = Task.Run(() => bridge.SendAsync("sheet.list", "untouched", new Poison(), cancellationToken: cancel.Token));
        try
        {
            Assert.True(entered.Wait(TimeSpan.FromSeconds(3)));
            cancel.Cancel();
        }
        finally { release.Set(); }
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => request);
        Assert.Equal(0, backend.Calls);
        Assert.Empty(backend.Revealed);
        Assert.Equal(0, backend.ShutdownCalls);
        Assert.True((await bridge.SendAsync("sheet.list", "usable")).Success);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(4_294_968)]
    [InlineData(int.MaxValue)]
    public async Task InvalidTimeoutCannotInitializeOrSerialize(int seconds)
    {
        var factories = 0;
        var backend = new Backend { Fast = true };
        using var bridge = new Bridge(() => { factories++; return backend; }, maxOutstandingDispatches: 1);
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() =>
            bridge.SendAsync("sheet.list", "unused", new Poison(), timeoutSeconds: seconds));
        Assert.Equal(0, factories);
        Assert.Equal(0, backend.Calls);
        Assert.True((await bridge.SendAsync("sheet.list", "usable")).Success);
        Assert.Equal(1, factories);
    }

    [Fact]
    public async Task ZeroDeadlineCannotStartOrRevealWork()
    {
        var backend = new Backend { Fast = true };
        using var bridge = new Bridge(() => backend);
        var response = await bridge.SendAsync("sheet.list", "untouched", timeoutSeconds: 0);
        Assert.False(response.Success);
        Assert.Equal("Timeout", response.ErrorCategory);
        Assert.Equal(0, backend.Calls);
        Assert.Empty(backend.Revealed);
        Assert.Equal(0, backend.ShutdownCalls);
    }

    [Theory]
    [InlineData("session.open")]
    [InlineData("session.create")]
    [InlineData("session.test")]
    public async Task StartupOperationsKeepTheirExistingServiceOwnedDeadline(string command)
    {
        var backend = new Backend { Fast = true };
        using var bridge = new Bridge(() => backend);
        Assert.True((await bridge.SendAsync(command, "startup", timeoutSeconds: 0)).Success);
        Assert.False(backend.Requests["startup"].CallerCancellation.IsCancellationRequested);
        Assert.Equal(1, backend.Calls);
    }

    [Fact]
    public async Task SerializationFailureDoesNotConsumeCapacity()
    {
        var backend = new Backend { Fast = true };
        using var bridge = new Bridge(() => backend, maxOutstandingDispatches: 1);
        await Assert.ThrowsAnyAsync<Exception>(() => bridge.SendAsync("sheet.list", "unused", new Poison()));
        Assert.Equal(0, backend.Calls);
        Assert.Empty(backend.Revealed);
        Assert.True((await bridge.SendAsync("sheet.list", "usable")).Success);
    }

    [Fact]
    public async Task FactoryResultRemainsOwnedWhenConcurrentShutdownDisposalFails()
    {
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var backend = new Backend { ThrowOnDispose = true };
        using var bridge = new Bridge(() =>
        {
            entered.Set();
            if (!release.Wait(TimeSpan.FromSeconds(5))) throw new TimeoutException("Factory not released.");
            return backend;
        });
        var request = Task.Run(() => bridge.SendAsync("session.list"));
        try
        {
            Assert.True(entered.Wait(TimeSpan.FromSeconds(3)));
            bridge.Dispose();
        }
        finally { release.Set(); }
        try
        {
            var failure = await request;
            Assert.False(failure.Success);
            Assert.Equal("ServiceStartup", failure.ErrorCategory);
            Assert.Equal(1, backend.DisposeCalls);
            Assert.Equal(0, backend.Calls);
            backend.ThrowOnDispose = false;
            bridge.Dispose();
            Assert.True(backend.Disposed);
            Assert.Equal(2, backend.DisposeCalls);
            Assert.Equal(2, backend.ShutdownCalls);
        }
        finally
        {
            backend.ThrowOnDispose = false;
        }
    }

    private static async Task Started(Backend backend) =>
        Assert.True(await backend.Started.WaitAsync(TimeSpan.FromSeconds(3)));

    private static async Task<ServiceResponse> WhenAdmitted(Bridge bridge)
    {
        for (var attempt = 0; attempt < 100; attempt++)
        {
            var response = await bridge.SendAsync("sheet.list", "after-recovery");
            if (response.Success) return response;
            Assert.Equal("ServiceBusy", response.ErrorCategory);
            await Task.Delay(10);
        }
        throw new TimeoutException("Recovery did not release admission.");
    }

    private sealed class Poison
    {
        private readonly string _message = "Rejected work must not be serialized.";
        public string Value => throw new InvalidOperationException(_message);
    }

    private sealed class Backend : IServiceBridgeBackend
    {
        private int _calls;
        internal int Calls => Volatile.Read(ref _calls);
        internal int DisposeCalls { get; private set; }
        internal int ShutdownCalls { get; private set; }
        internal bool Disposed { get; private set; }
        internal bool Fast { get; set; }
        internal bool BlockReveal { get; init; }
        internal bool CompleteOnReveal { get; init; }
        internal bool ThrowOnDispose { get; set; }
        internal ConcurrentDictionary<string, TaskCompletionSource<ServiceResponse>> Pending { get; } = new();
        internal ConcurrentDictionary<string, ServiceRequest> Requests { get; } = new();
        internal ConcurrentQueue<string> Revealed { get; } = new();
        internal SemaphoreSlim Started { get; } = new(0);
        internal SemaphoreSlim RevealEntered { get; } = new(0);
        internal ManualResetEventSlim ReleaseReveal { get; } = new();
        public Task<ServiceResponse> ProcessAsync(ServiceRequest request)
        {
            var call = Interlocked.Increment(ref _calls);
            var key = request.SessionId ?? $"new{call}";
            Requests[key] = request;
            var pending = Pending.GetOrAdd(key, _ => new(TaskCreationOptions.RunContinuationsAsynchronously));
            Started.Release();
            if (Fast) pending.TrySetResult(new ServiceResponse { Success = true });
            return pending.Task;
        }
        public void RevealSession(string sessionId)
        {
            Revealed.Enqueue(sessionId);
            if (CompleteOnReveal) Complete(sessionId);
            RevealEntered.Release();
            if (BlockReveal && !ReleaseReveal.Wait(TimeSpan.FromSeconds(5)))
                throw new TimeoutException("Reveal was not released.");
        }
        internal void Complete(string id)
        {
            if (Pending.TryGetValue(id, out var pending))
                pending.TrySetResult(new ServiceResponse { Success = true });
        }
        internal void CompleteCreation(string id)
        {
            if (Pending.TryGetValue(id, out var pending))
                pending.TrySetResult(new ServiceResponse
                {
                    Success = true,
                    Result = JsonSerializer.Serialize(new { sessionId = "late-" + id })
                });
        }
        public string? RequestShutdown()
        {
            ShutdownCalls++;
            return null;
        }
        public void Dispose()
        {
            DisposeCalls++;
            if (ThrowOnDispose) throw new InvalidOperationException("Synthetic disposal failure.");
            Disposed = true;
            ReleaseReveal.Set();
            foreach (var pending in Pending.Values)
                pending.TrySetResult(new ServiceResponse { Success = false });
        }
    }
}
