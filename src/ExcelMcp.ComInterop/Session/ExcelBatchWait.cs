namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal static class ExcelBatchWait
{
    internal static T Execute<T>(
        IExcelBatch batch,
        TimeSpan timeout,
        Func<ExcelContext, CancellationToken, T> operation) =>
        Execute(batch.Execute, timeout, operation);

    internal static T Execute<T>(
        Func<Func<ExcelContext, CancellationToken, T>, CancellationToken, T> execute,
        TimeSpan timeout,
        Func<ExcelContext, CancellationToken, T> operation)
    {
        using var deadline = new CancellationTokenSource(timeout);
        try
        {
            return execute(operation, deadline.Token);
        }
        catch (OperationCanceledException ex) when (
            deadline.IsCancellationRequested && ex.CancellationToken == deadline.Token)
        {
            throw new TimeoutException(
                $"The Excel wait expired after {timeout.TotalSeconds:F0} seconds. " +
                "Inspect session status: execution may still be pending, and the wait does not undo workbook changes.", ex);
        }
    }
}
