using System.Text;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "TransportLifetime")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class StdioShutdownDiagnosticsTests
{
    [Fact]
    public async Task DisabledDiagnosticsWriteNothing()
    {
        var diagnostics = new StdioShutdownDiagnostics(enabled: false);
        diagnostics.Record(StdioStopReason.HostFailure);
        bool opened = false;
        Assert.False(await diagnostics.ReportAsync(() =>
        {
            opened = true;
            return TextWriter.Null;
        }, 1));
        Assert.False(opened);
    }

    [Fact]
    public async Task ReportsFirstReasonAndFinalExitCode()
    {
        var diagnostics = new StdioShutdownDiagnostics(enabled: true);
        diagnostics.Record(StdioStopReason.StdoutFailure);
        diagnostics.Record(StdioStopReason.ParentExit);
        using var output = new StringWriter();
        Assert.True(await diagnostics.ReportAsync(() => output, 1));
        Assert.Equal("ExcelMCP stdio stopped: reason=stdout_failure exit=1" + Environment.NewLine,
            output.ToString());
    }

    [Fact]
    public async Task ClosedStderrDoesNotEscapeDiagnosticBoundary()
    {
        var diagnostics = new StdioShutdownDiagnostics(enabled: true);
        using var output = new StringWriter();
        output.Close();
        Assert.False(await diagnostics.ReportAsync(() => output, 0));
    }

    [Fact]
    public async Task FailureOpeningStderrDoesNotEscapeDiagnosticBoundary()
    {
        var diagnostics = new StdioShutdownDiagnostics(enabled: true);
        Assert.False(await diagnostics.ReportAsync(
            () => throw new InvalidOperationException("Controlled stderr initialization failure"), 1));
    }

    [Fact]
    public async Task FailureWritingStderrDoesNotEscapeDiagnosticBoundary()
    {
        var diagnostics = new StdioShutdownDiagnostics(enabled: true);
        using var output = new FailingWriter();
        Assert.False(await diagnostics.ReportAsync(() => output, 1));
    }

    [Fact]
    public async Task StderrThatIsNotDrainedDoesNotHoldProcessExit()
    {
        using var output = new BlockedWriter();
        var diagnostics = new StdioShutdownDiagnostics(enabled: true);
        try
        {
            var report = diagnostics.ReportAsync(() => output, 0);
            await output.Entered.Task.WaitAsync(TimeSpan.FromSeconds(3));
            Assert.False(await report.WaitAsync(TimeSpan.FromSeconds(3)));
            Assert.False(output.Finished.Task.IsCompleted);
        }
        finally
        {
            output.Release.TrySetResult();
            await output.Finished.Task.WaitAsync(TimeSpan.FromSeconds(3));
        }
    }

    private sealed class FailingWriter : TextWriter
    {
        public override Encoding Encoding => Encoding.UTF8;
        public override void WriteLine(string? value) => throw new InvalidOperationException("Controlled stderr write failure");
    }

    private sealed class BlockedWriter : TextWriter
    {
        internal TaskCompletionSource Entered { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource Release { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource Finished { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public override Encoding Encoding => Encoding.UTF8;

        public override void WriteLine(string? value)
        {
            Entered.TrySetResult();
            Release.Task.GetAwaiter().GetResult();
            Finished.TrySetResult();
        }
    }
}
