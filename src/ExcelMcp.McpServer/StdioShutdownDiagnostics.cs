namespace Sbroenne.ExcelMcp.McpServer;

internal enum StdioStopReason
{
    HostShutdown,
    StdinEof,
    StdinFailure,
    StdoutFailure,
    ParentExit,
    Canceled,
    HostFailure
}

/// <summary>Records the first observed stop signal without writing on transport or timer callbacks.</summary>
internal sealed class StdioShutdownDiagnostics(bool enabled)
{
    private int _reason;

    internal void Record(StdioStopReason reason)
    {
        if (enabled)
            Interlocked.CompareExchange(ref _reason, (int)reason, (int)StdioStopReason.HostShutdown);
    }

    internal async Task<bool> ReportAsync(Func<TextWriter> openStderr, int exitCode)
    {
        if (!enabled)
            return false;

        string reason = (StdioStopReason)Volatile.Read(ref _reason) switch
        {
            StdioStopReason.StdinEof => "stdin_eof",
            StdioStopReason.StdinFailure => "stdin_failure",
            StdioStopReason.StdoutFailure => "stdout_failure",
            StdioStopReason.ParentExit => "parent_exit",
            StdioStopReason.Canceled => "canceled",
            StdioStopReason.HostFailure => "host_failure",
            _ => "host_shutdown"
        };

        // A terminating client can also stop draining stderr. Diagnostics are
        // best effort and run only after the host and workbook cleanup returned.
        var write = Task.Run(() => TryWrite(openStderr, $"ExcelMCP stdio stopped: reason={reason} exit={exitCode}"));
        var completed = await Task.WhenAny(write, Task.Delay(250)).ConfigureAwait(false);
        return completed == write && await write.ConfigureAwait(false);
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "A diagnostic writer is external to the host. If opening or writing stderr fails, return false and preserve the real process exit status.")]
    private static bool TryWrite(Func<TextWriter> openStderr, string message)
    {
        try
        {
            openStderr().WriteLine(message);
            return true;
        }
#pragma warning disable CA1031 // A stderr failure must never replace the host outcome.
        catch (Exception)
        {
            return false;
        }
#pragma warning restore CA1031
    }
}
