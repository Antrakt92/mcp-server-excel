using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Theory]
    [InlineData("returned")]
    [InlineData("failed")]
    [InlineData("cancelled")]
    public async Task ExecutionDiagnostics_SnapshotFailurePreservesCallbackCauseAndRetainedEdit(string callbackOutcome)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var sessionId = await CreateSessionAsync(service, Path.Join(directory, "execution-diagnostics.xlsx"));
            sessions[sessionId] = 0;
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(sessionId));
            var identity = Assert.IsType<ExcelProcessIdentity>(batch.ProcessIdentity);
            using var cancellation = new CancellationTokenSource();
            Exception? operationFailure = callbackOutcome switch
            {
                "returned" => null,
                "failed" => new InvalidOperationException("Synthetic callback failure after the edit."),
                "cancelled" => new OperationCanceledException("Synthetic callback cancellation after the edit.", cancellation.Token),
                _ => throw new ArgumentException("Unknown callback outcome.", nameof(callbackOutcome))
            };
#pragma warning disable CA2201 // Injected publication failure; the callback itself uses real Excel.
            var snapshotFailure = new COMException("Synthetic visibility publication failure.", unchecked((int)0x80020009));
#pragma warning restore CA2201
            var marker = $"Retained edit after callback {callbackOutcome}";
            var callbackCalls = 0;
            Exception failure;
            batch.BeforeVisibilitySnapshotPublicationHookForTests = () => throw snapshotFailure;
            try
            {
                // Exercise the internal execution boundary directly so Service does not
                // classify or serialize the original exception before these assertions.
                failure = Assert.ThrowsAny<Exception>(() => batch.Execute((context, _) =>
                {
                    callbackCalls++;
                    WriteExecutionDiagnosticsMarker(context, marker);
                    if (callbackOutcome == "cancelled")
                        cancellation.Cancel();
                    if (operationFailure is not null)
                        throw operationFailure;
                    return 7;
                }));
            }
            finally
            {
                batch.BeforeVisibilitySnapshotPublicationHookForTests = null;
            }

            Assert.Equal(1, callbackCalls);
            if (operationFailure is null)
            {
                Assert.Same(snapshotFailure, failure);
            }
            else
            {
                AggregateException aggregate;
                if (operationFailure is OperationCanceledException originalCancellation)
                {
                    var cancelled = Assert.IsType<OperationCanceledException>(failure);
                    Assert.Equal(originalCancellation.CancellationToken, cancelled.CancellationToken);
                    Assert.True(cancelled.CancellationToken.IsCancellationRequested);
                    Assert.Equal(originalCancellation.Message, cancelled.Message);
                    aggregate = Assert.IsType<AggregateException>(cancelled.InnerException);
                }
                else
                {
                    aggregate = Assert.IsType<AggregateException>(failure);
                }
                Assert.Collection(aggregate.InnerExceptions,
                    cause => Assert.Same(operationFailure, cause),
                    cause => Assert.Same(snapshotFailure, cause));
            }

            Assert.True(ExcelExecutionFailureContext.TryFind(failure, out var failureContext, out var complete));
            Assert.True(complete);
            Assert.NotNull(failureContext);
            Assert.Equal(operationFailure is null, failureContext.CallbackReturned);
            Assert.Equal(operationFailure?.GetType().Name, failureContext.OperationExceptionType);
            Assert.Equal(operationFailure?.HResult, failureContext.OperationHResult);
            Assert.Equal(nameof(COMException), failureContext.SnapshotExceptionType);
            Assert.Equal(unchecked((int)0x80020009), failureContext.SnapshotHResult);

            Assert.Same(batch, service.SessionManager.GetSession(sessionId));
            Assert.Equal(identity, service.SessionManager.GetSessionProcessIdentity(sessionId));
            Assert.Contains(identity, SessionManager.GetTrackedExcelProcesses());
            Assert.True(OwnedProcessGuard.MatchesLiveProcess(identity));
            Assert.True(batch.IsExcelProcessAlive());
            Assert.False(batch.HasTimedOutOperation);
            Assert.Equal(marker, await ReadMarkerAsync(service, sessionId));
            Assert.False(batch.Execute((context, _) => context.Book.Saved));
        });
    }

    private static void WriteExecutionDiagnosticsMarker(ExcelContext context, string marker)
    {
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;
        Excel.Range? cell = null;
        try
        {
            sheets = context.Book.Worksheets;
            sheet = (Excel.Worksheet)sheets["Sheet1"];
            cell = sheet.Range["A1"];
            cell.Value2 = marker;
        }
        finally
        {
            ComUtilities.Release(ref cell);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref sheets);
        }
    }
}
