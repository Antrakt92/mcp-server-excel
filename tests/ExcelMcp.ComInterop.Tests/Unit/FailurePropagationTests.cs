using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Formatting;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ErrorHandling")]
[Trait("RequiresExcel", "false")]
[System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2201",
    Justification = "Synthetic COM failures prove the exact exception crosses internal boundaries without calling desktop Excel.")]
public sealed class FailurePropagationTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Formatting_CallerCancellation_Propagates(bool dax)
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => dax
            ? DaxFormatter.FormatAsync("1 + 1", cancellation.Token)
            : MCodeFormatter.FormatAsync("let a = 1 in a", cancellation.Token));
    }

    [Fact]
    public void NumberFormatLocale_UnexpectedComFailure_PreservesOriginalError()
    {
        var error = new COMException("Cannot read locale", unchecked((int)0x80070005));
        var application = DispatchProxy.Create<Excel.Application, ApplicationProxy>();
        Assert.Equal("General", (object?)application.International[Excel.XlApplicationInternational.xlGeneralFormatName]);
        ((ApplicationProxy)(object)application).Error = error;

        Assert.Same(error, Assert.Throws<COMException>(() => new NumberFormatTranslator(application)));
    }

    [Fact]
    public void RefreshProbe_UnexpectedComFailure_PreservesOriginalError()
    {
        var error = new COMException("Cannot read readiness", unchecked((int)0x80070005));
        var application = DispatchProxy.Create<Excel.Application, ApplicationProxy>();
        Assert.True(application.Ready);
        ((ApplicationProxy)(object)application).Error = error;

        Assert.Same(error, Assert.Throws<COMException>(() =>
            WorkbookRefreshProbe.Read(application, null!)));
    }

    [Fact]
    public void MacroReopenClose_RefusesUnsavedWorkbookBeforeCallingClose()
    {
        var application = DispatchProxy.Create<Excel.Application, ApplicationProxy>();
        var workbook = DispatchProxy.Create<Excel.Workbook, CloseWorkbookProxy>();

        var response = Assert.Throws<InvalidOperationException>(() =>
            ExcelShutdownService.CloseSavedWorkbookForReopen(application, workbook));

        Assert.Contains("unsaved changes", response.Message, StringComparison.Ordinal);
        Assert.Equal(0, ((CloseWorkbookProxy)(object)workbook).CloseCalls);
        Assert.Equal(0, ((ApplicationProxy)(object)application).WorkbooksReads);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MacroReopenClose_ReportsUncertaintyAfterNativeDispatch(bool closeReturns)
    {
        var cause = new COMException("Synthetic close or readback failure.", unchecked((int)0x800A03EC));
        var application = DispatchProxy.Create<Excel.Application, ApplicationProxy>();
        var applicationProxy = (ApplicationProxy)(object)application;
        var workbook = DispatchProxy.Create<Excel.Workbook, CloseWorkbookProxy>();
        var workbookProxy = (CloseWorkbookProxy)(object)workbook;
        workbookProxy.Saved = true;
        if (closeReturns)
            applicationProxy.Error = cause;
        else
            workbookProxy.CloseError = cause;

        var response = Assert.Throws<WorkbookSecurityCompletionUnconfirmedException>(() =>
            ExcelShutdownService.CloseSavedWorkbookForReopen(application, workbook));

        Assert.Same(cause, response.InnerException);
        Assert.Equal(1, workbookProxy.CloseCalls);
        Assert.Equal(closeReturns ? 1 : 0, applicationProxy.WorkbooksReads);
    }

    [Fact]
    public void MacroReopenClose_ConfirmedStillOpenIsAnOrdinaryRefusal()
    {
        var application = DispatchProxy.Create<Excel.Application, ApplicationProxy>();
        var workbook = DispatchProxy.Create<Excel.Workbook, CloseWorkbookProxy>();
        var workbookProxy = (CloseWorkbookProxy)(object)workbook;
        workbookProxy.Saved = true;
        var workbooks = DispatchProxy.Create<Excel.Workbooks, WorkbooksProxy>();
        ((WorkbooksProxy)(object)workbooks).Workbook = workbook;
        Assert.Same(workbook, workbooks[1]);
        ((ApplicationProxy)(object)application).Workbooks = workbooks;

        var response = Assert.Throws<InvalidOperationException>(() =>
            ExcelShutdownService.CloseSavedWorkbookForReopen(application, workbook));

        Assert.Contains("did not close", response.Message, StringComparison.Ordinal);
        Assert.Equal(1, workbookProxy.CloseCalls);
        Assert.Equal(1, ((ApplicationProxy)(object)application).WorkbooksReads);
    }

    [Theory]
    [InlineData(false, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    [InlineData(false, true, false)]
    public void OpeningSettingsRestoration_PreservesEveryFailureAndMarksStateUnconfirmed(
        bool primaryFailed, bool restoreAlerts, bool failSecurity)
    {
        var app = DispatchProxy.Create<Excel.Application, OpeningSettingsProxy>();
        var proxy = (OpeningSettingsProxy)(object)app;
        var securityFailure = new COMException("Synthetic AutomationSecurity restoration failure.");
        var alertsFailure = new COMException("Synthetic DisplayAlerts restoration failure.");
        var primary = primaryFailed ? new InvalidOperationException("Synthetic prior operation failure.") : null;
        if (failSecurity)
            proxy.SecurityFailure = securityFailure;
        proxy.AlertsFailure = alertsFailure;

        var response = Assert.Throws<WorkbookSecurityCompletionUnconfirmedException>(() =>
            ExcelBatch.RestoreOpeningSettings(app, Office.MsoAutomationSecurity.msoAutomationSecurityByUI,
                restoreAlerts ? false : null, primary));

        var failures = Assert.IsType<AggregateException>(response.InnerException).InnerExceptions;
        List<Exception> expected = [];
        if (primary is not null)
            expected.Add(primary);
        if (failSecurity)
            expected.Add(securityFailure);
        if (restoreAlerts)
            expected.Add(alertsFailure);
        Assert.Equal(expected, failures);
        Assert.Equal(1, proxy.SecurityWrites);
        Assert.Equal(restoreAlerts ? 1 : 0, proxy.AlertsWrites);
    }

    [Fact]
    public void OpeningSettingsRestoration_SuccessDoesNotRaiseCompletionFailure()
    {
        var app = DispatchProxy.Create<Excel.Application, OpeningSettingsProxy>();
        var proxy = (OpeningSettingsProxy)(object)app;

        ExcelBatch.RestoreOpeningSettings(app, Office.MsoAutomationSecurity.msoAutomationSecurityByUI, false, null);

        Assert.Equal(Office.MsoAutomationSecurity.msoAutomationSecurityByUI, proxy.Security);
        Assert.False(proxy.Alerts);
    }

    public class OpeningSettingsProxy : DispatchProxy
    {
        public Exception? SecurityFailure { get; set; }
        public Exception? AlertsFailure { get; set; }
        public int SecurityWrites { get; private set; }
        public int AlertsWrites { get; private set; }
        public Office.MsoAutomationSecurity? Security { get; private set; }
        public bool Alerts { get; private set; } = true;

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            switch (targetMethod?.Name)
            {
                case "set_AutomationSecurity":
                    SecurityWrites++;
                    if (SecurityFailure is not null)
                        throw SecurityFailure;
                    Security = (Office.MsoAutomationSecurity)args![0]!;
                    return null;
                case "set_DisplayAlerts":
                    AlertsWrites++;
                    if (AlertsFailure is not null)
                        throw AlertsFailure;
                    Alerts = (bool)args![0]!;
                    return null;
                default:
                    throw new NotSupportedException(targetMethod?.Name);
            }
        }
    }

    [Fact]
    public void WriteGuard_FailedCapture_PropagatesAndDoesNotLeakNesting()
    {
        var error = new COMException("Cannot read ScreenUpdating", unchecked((int)0x80070005));
        var application = DispatchProxy.Create<Excel.Application, ApplicationProxy>();
        var proxy = (ApplicationProxy)(object)application;
        proxy.Error = error;

        Assert.Same(error, Assert.Throws<COMException>(() => new ExcelWriteGuard(application)));
        proxy.Error = null;
        proxy.Workbooks = DispatchProxy.Create<Excel.Workbooks, WorkbooksProxy>();
        Assert.Equal(0, proxy.Workbooks.Count);
        using (var guard = new ExcelWriteGuard(application))
            Assert.False(proxy.ScreenUpdating);
        Assert.True(proxy.ScreenUpdating);
    }

    public class ApplicationProxy : DispatchProxy
    {
        public Exception? Error { get; set; }
        public bool ScreenUpdating { get; private set; } = true;
        public Excel.Workbooks? Workbooks { get; set; }
        public int WorkbooksReads { get; private set; }

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            if (targetMethod?.Name == "get_Workbooks")
                WorkbooksReads++;
            if (Error != null)
                throw Error;
            return targetMethod?.Name switch
            {
                "get_ScreenUpdating" => ScreenUpdating,
                "get_Ready" => true,
                "get_International" => "General",
                "get_Workbooks" => Workbooks,
                "set_ScreenUpdating" => SetScreenUpdating((bool)args![0]!),
                _ => throw new NotSupportedException(targetMethod?.Name)
            };
        }

        private object? SetScreenUpdating(bool value)
        {
            ScreenUpdating = value;
            return null;
        }
    }

    public class CloseWorkbookProxy : DispatchProxy
    {
        public bool Saved { get; set; }
        public Exception? CloseError { get; set; }
        public int CloseCalls { get; private set; }

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            if (targetMethod?.Name == "get_Saved")
                return Saved;
            if (targetMethod?.Name != "Close")
                throw new NotSupportedException(targetMethod?.Name);
            CloseCalls++;
            if (CloseError != null)
                throw CloseError;
            return null;
        }
    }

    public class WorkbooksProxy : DispatchProxy
    {
        public Excel.Workbook? Workbook { get; set; }

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args) => targetMethod?.Name switch
        {
            "get_Count" => Workbook is null ? 0 : 1,
            "get__Default" when args is [1] => Workbook,
            _ => throw new NotSupportedException(targetMethod?.Name)
        };
    }
}
