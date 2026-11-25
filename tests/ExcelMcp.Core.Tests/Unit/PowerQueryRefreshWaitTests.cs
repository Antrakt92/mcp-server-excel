using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

[Trait("RequiresExcel", "false")]
public sealed class PowerQueryRefreshWaitTests
{
    [Fact]
    public void RefreshWait_WithPreCancelledToken_DoesNotReportIdleAsSuccess()
    {
        using var cts = new CancellationTokenSource();
        cts.Cancel();
        var cancellationRequested = false;

        var exception = Assert.Throws<TargetInvocationException>(() => GetRefreshWaitMethod().Invoke(null,
        [
            (Func<bool>)(() => false),
            (Action)(() => cancellationRequested = true),
            cts.Token
        ]));

        var cancellation = Assert.IsType<OperationCanceledException>(exception.InnerException);
        Assert.Equal(cts.Token, cancellation.CancellationToken);
        Assert.True(cancellationRequested);
    }

    [Fact]
    public void RefreshWait_WhenProviderCannotCancel_PreservesBothFailures()
    {
        using var cts = new CancellationTokenSource();
        cts.Cancel();
#pragma warning disable CA2201 // Synthetic provider cancellation failure.
        var providerFailure = new COMException("Provider did not cancel", unchecked((int)0x80004005));
#pragma warning restore CA2201

        var exception = Assert.Throws<TargetInvocationException>(() => GetRefreshWaitMethod().Invoke(null,
        [
            (Func<bool>)(() => true),
            (Action)(() => throw providerFailure),
            cts.Token
        ]));

        var cancellation = Assert.IsType<OperationCanceledException>(exception.InnerException);
        Assert.Equal(cts.Token, cancellation.CancellationToken);
        Assert.Contains("may still be running", cancellation.Message, StringComparison.Ordinal);
        var causes = Assert.IsType<AggregateException>(cancellation.InnerException);
        Assert.Equal(cts.Token, Assert.IsType<OperationCanceledException>(causes.InnerExceptions[0]).CancellationToken);
        Assert.Same(providerFailure, causes.InnerExceptions[1]);
    }

    [Fact]
    public void RefreshWait_WhenCancellationRequested_InvokesCancelActionAndThrows()
    {
        MethodInfo waitMethod = GetRefreshWaitMethod();

        bool cancelCalled = false;
        using var cts = new CancellationTokenSource();

        var cancellationThread = new Thread(() =>
        {
            Thread.Sleep(50);
            cts.Cancel();
        });
        cancellationThread.Start();

        try
        {
            var exception = Assert.Throws<TargetInvocationException>(() =>
                waitMethod.Invoke(null,
                [
                    (Func<bool>)(() => true),
                    (Action)(() => cancelCalled = true),
                    cts.Token
                ]));

            Assert.IsType<OperationCanceledException>(exception.InnerException);
            Assert.True(cancelCalled);
        }
        finally
        {
            cancellationThread.Join();
        }
    }

    [Fact]
    public void RefreshWait_WhenRefreshCompletes_DoesNotInvokeCancelAction()
    {
        MethodInfo waitMethod = GetRefreshWaitMethod();

        int pollCount = 0;
        bool cancelCalled = false;

        waitMethod.Invoke(null,
        [
            (Func<bool>)(() => Interlocked.Increment(ref pollCount) == 1),
            (Action)(() => cancelCalled = true),
            CancellationToken.None
        ]);

        Assert.True(pollCount >= 2);
        Assert.False(cancelCalled);
    }

    private static MethodInfo GetRefreshWaitMethod()
    {
        var waitMethod = typeof(PowerQueryCommands).GetMethod(
            "WaitForRefreshCompletion",
            BindingFlags.NonPublic | BindingFlags.Static);

        return waitMethod ?? throw new InvalidOperationException(
            "Expected private method WaitForRefreshCompletion was not found.");
    }
}
