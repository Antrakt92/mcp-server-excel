using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Utilities;

internal static class ConnectionRefreshWait
{
    internal static void Wait(
        Func<bool> isRefreshing,
        Action cancelRefresh,
        CancellationToken cancellationToken,
        int pollIntervalMs)
    {
        var elapsed = Stopwatch.StartNew();
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!isRefreshing())
                return;

            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                // Avoid pumping STA COM messages while waiting. Excel callbacks can wake
                // managed sleeps early and turn a cross-process status poll into a spin loop.
                ComUtilities.KernelSleep(pollIntervalMs);
                cancellationToken.ThrowIfCancellationRequested();
                if (elapsed.ElapsedMilliseconds < pollIntervalMs)
                    continue;
                elapsed.Restart();
                if (!isRefreshing())
                    return;
            }
        }
        catch (OperationCanceledException ex)
        {
            try
            {
                cancelRefresh();
            }
            catch (Exception cancelFailure) when (cancelFailure is COMException or RuntimeBinderException)
            {
                throw new OperationCanceledException(
                    "Refresh was cancelled or timed out, but Excel could not confirm cancellation. " +
                    "The query may still be running; inspect refresh status before saving or closing.",
                    new AggregateException(ex, cancelFailure), cancellationToken);
            }
            throw;
        }
    }
}
