using System.Diagnostics;
using Sbroenne.ExcelMcp.Service.Rpc;
using Xunit;
using ClientPipeSecurity = Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceSecurity;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceDaemon")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class TerminalSessionPipeTests
{
    [Fact]
    public void DefaultCliPipeIncludesTheCurrentWindowsSessionInBothHelpers()
    {
        using var process = Process.GetCurrentProcess();
        var pipe = ServiceSecurity.GetCliPipeName();
        Assert.EndsWith($"-session-{process.SessionId}", pipe, StringComparison.Ordinal);
        Assert.Equal(pipe, ClientPipeSecurity.GetCliPipeName());
    }

    [Theory(Timeout = 15000)]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExplicitPipeRejectsUnverifiedSessionThenAcceptsNativeSameSessionPeer(bool unknown)
    {
        var pipeName = $"excelmcp-terminal-session-{Guid.NewGuid():N}";
        int requests = 0;
        int probes = 0;
        using var current = Process.GetCurrentProcess();
        var differentSession = (uint)current.SessionId + 1;
        using var host = new DaemonHost(_ =>
        {
            Interlocked.Increment(ref requests);
            return Task.FromResult(new ServiceResponse { Success = true });
        });
        host.ReadClientSessionIdForTests = _ =>
        {
            Interlocked.Increment(ref probes);
            return unknown ? null : differentSession;
        };
        var run = host.RunAsync(pipeName);
        try
        {
            using var rejected = new ServiceClient(pipeName);
            var rejectedResponse = await rejected.SendAsync(
                new ServiceRequest { Command = "service.ping" }, TimeSpan.FromSeconds(3));
            Assert.False(rejectedResponse.Success);
            Assert.Equal(0, Volatile.Read(ref requests));
            Assert.Equal(1, Volatile.Read(ref probes));

            host.ReadClientSessionIdForTests = null;
            using var accepted = new ServiceClient(pipeName);
            var acceptedResponse = await accepted.SendAsync(
                new ServiceRequest { Command = "service.ping" }, TimeSpan.FromSeconds(3));
            Assert.True(acceptedResponse.Success, acceptedResponse.ErrorMessage);
            Assert.Equal(1, Volatile.Read(ref requests));
        }
        finally
        {
            host.RequestShutdown();
            await run.WaitAsync(TimeSpan.FromSeconds(4));
        }
    }
}
