using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging.Abstractions;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "SessionLifecycle")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class SaveBusyDiagnosticTests
{
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ObservedSuccessfulSave_PreservesNativeFailureWithoutRetrying(bool saveCallFails)
    {
        Action<bool>? afterSave = null;
        var nativeFailure = Assert.IsType<COMException>(Marshal.GetExceptionForHR(unchecked((int)0x800AC472)));
        int saveAttempts = 0;
        int readAttempts = 0;

        var failure = Assert.Throws<WorkbookSaveIncompleteException>(() => ExcelShutdownService.SaveAndObserve(
            handler => afterSave = handler,
            () => afterSave = null,
            () =>
            {
                saveAttempts++;
                afterSave!(true);
                if (saveCallFails)
                    throw nativeFailure;
            },
            () =>
            {
                readAttempts++;
                throw nativeFailure;
            }, "fixture.xlsx", NullLogger.Instance));

        Assert.Same(nativeFailure, failure.InnerException);
        Assert.True(failure.ExcelReportedSaveSucceeded);
        Assert.Null(failure.HasUnsavedChanges);
        Assert.Equal(1, saveAttempts);
        Assert.Equal(saveCallFails ? 0 : 1, readAttempts);
        Assert.Null(afterSave);
    }

    [Fact]
    public void UnobservedSaveFailure_PropagatesAsNativeFailureForExistingRetryPolicy()
    {
        var nativeFailure = Assert.IsType<COMException>(Marshal.GetExceptionForHR(unchecked((int)0x800AC472)));

        var failure = Assert.Throws<COMException>(() => ExcelShutdownService.SaveAndObserve(
            _ => { }, () => { }, () => throw nativeFailure,
            () => throw new InvalidOperationException("Saved must not be read after a failed save."),
            "fixture.xlsx", NullLogger.Instance));

        Assert.Same(nativeFailure, failure);
    }

    [Fact]
    public void FinalSaveBusy_PreservesTypedRefusalAndInnerComHResult()
    {
        var comFailure = Assert.IsType<COMException>(Marshal.GetExceptionForHR(unchecked((int)0x800AC472)));
        var failure = ExcelShutdownService.CreateSaveFailureException(comFailure, "fixture.xlsx");

        var busy = Assert.IsType<ExcelBusyException>(failure);
        Assert.Same(comFailure, busy.InnerException);
        Assert.Equal(unchecked((int)0x800AC472), Assert.IsType<COMException>(busy.InnerException).HResult);
        Assert.Contains("not been saved", busy.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(unchecked((int)0x800A03EC))]
    [InlineData(ResiliencePipelines.RPC_E_DISCONNECTED)]
    public void OtherSaveFailures_RetainTheirExistingCategoryAndComCause(int hresult)
    {
        var comFailure = Assert.IsType<COMException>(Marshal.GetExceptionForHR(hresult));
        var failure = ExcelShutdownService.CreateSaveFailureException(comFailure, "fixture.xlsx");

        Assert.IsType<InvalidOperationException>(failure);
        Assert.Same(comFailure, failure.InnerException);
    }

    [Fact]
    public void SaveBusy_DoesNotInventAnExternalLock_AndExplainsUnsavedState()
    {
        var message = ExcelShutdownService.CreateSaveFailureMessage(
            unchecked((int)0x800AC472), "fixture.xlsx", "Excel declined automation.");

        Assert.Contains("Excel is busy", message, StringComparison.Ordinal);
        Assert.Contains("not been saved", message, StringComparison.Ordinal);
        Assert.Contains("open workbook", message, StringComparison.Ordinal);
        Assert.DoesNotContain("another user", message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("dialog", message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("locked", message, StringComparison.OrdinalIgnoreCase);
    }
}
