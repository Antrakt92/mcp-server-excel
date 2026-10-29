using System.Diagnostics;
using Microsoft.Extensions.Hosting;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "ParentProcessMonitor")]
[Trait("Speed", "Fast")]
[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
public sealed class ParentProcessMonitorTests
{
    private static Process Child() => Process.Start(new ProcessStartInfo
    {
        FileName = Environment.GetEnvironmentVariable("COMSPEC") ?? "cmd.exe",
        Arguments = "/c ping -n 30 127.0.0.1 >nul",
        UseShellExecute = false,
        CreateNoWindow = true
    }) ?? throw new InvalidOperationException("Child process did not start.");

    [Fact]
    public async Task LiveParentIsNotStoppedAndExitSignalsExactlyOnce()
    {
        using var child = Child();
        var lifetime = new FakeLifetime();
        try
        {
            using var monitor = ParentProcessMonitor.StartForProcess(lifetime, (uint)child.Id,
                TimeSpan.FromMilliseconds(10));
            Assert.NotNull(monitor);
            await Task.Delay(100);
            Assert.Equal(0, lifetime.Stops);
            child.Kill(entireProcessTree: true);
            await child.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(3));
            await lifetime.Stopped.Task.WaitAsync(TimeSpan.FromSeconds(3));
            await Task.Delay(100);
            Assert.Equal(1, lifetime.Stops);
        }
        finally { if (!child.HasExited) child.Kill(entireProcessTree: true); }
    }

    [Fact]
    public async Task DisposeDoesNotSignalLaterParentExit()
    {
        using var child = Child();
        var lifetime = new FakeLifetime();
        try
        {
            var monitor = ParentProcessMonitor.StartForProcess(lifetime, (uint)child.Id,
                TimeSpan.FromMilliseconds(10));
            Assert.NotNull(monitor);
            monitor.Dispose(); monitor.Dispose();
            child.Kill(entireProcessTree: true);
            await child.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(3));
            await Task.Delay(100);
            Assert.Equal(0, lifetime.Stops);
        }
        finally { if (!child.HasExited) child.Kill(entireProcessTree: true); }
    }

    [Fact]
    public void ReusedPidCannotBeMistakenForParent()
    {
        var lifetime = new FakeLifetime();
        using var self = Process.GetCurrentProcess();
        using var monitor = ParentProcessMonitor.StartForProcess(lifetime, (uint)self.Id,
            notCreatedAfter: self.StartTime.ToUniversalTime().ToFileTimeUtc() - 1);
        Assert.Null(monitor);
        Assert.Equal(1, lifetime.Stops);
    }

    [Fact]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "CA1835",
        Justification = "Exercises the required legacy Stream.ReadAsync override, not the Memory overload.")]
    public async Task EofReportsOncePreservesDataAndIgnoresEmptyRead()
    {
        var stops = 0;
        using var stream = new TransportLifetimeStream(new MemoryStream([10, 20]), () => stops++);
        Assert.Equal(0, stream.Read([], 0, 0));
        Assert.Equal(0, stops);
        var buffer = new byte[4];
        Assert.Equal(2, await stream.ReadAsync(buffer.AsMemory()));
        Assert.Equal(new byte[] { 10, 20 }, buffer[..2]);
        Assert.Equal(0, stops);
        Assert.Equal(0, await stream.ReadAsync(buffer, 0, buffer.Length));
        Assert.Equal(0, stream.Read(buffer));
        Assert.Equal(1, stops);
    }

    private sealed class FakeLifetime : IHostApplicationLifetime
    {
        private int _stops;
        public int Stops => Volatile.Read(ref _stops);
        internal TaskCompletionSource Stopped { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public CancellationToken ApplicationStarted => CancellationToken.None;
        public CancellationToken ApplicationStopping => CancellationToken.None;
        public CancellationToken ApplicationStopped => CancellationToken.None;
        public void StopApplication() { Interlocked.Increment(ref _stops); Stopped.TrySetResult(); }
    }
}
