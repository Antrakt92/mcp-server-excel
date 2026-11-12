using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PivotTables")]
public sealed class PivotTableRefreshTimeoutTests
{
    [Fact]
    public void Refresh_WithExplicitTimeout_ExpiresTheBatchWait()
    {
        using var batch = ObserveBatchToken(token =>
        {
            Assert.True(token.CanBeCanceled);
            Assert.True(token.WaitHandle.WaitOne(TimeSpan.FromSeconds(5)));
            token.ThrowIfCancellationRequested();
        });

        var failure = Assert.Throws<TimeoutException>(() =>
            new PivotTableCommands().Refresh(batch, "Report", TimeSpan.FromMilliseconds(20)));

        Assert.IsType<OperationCanceledException>(failure.InnerException);
        Assert.Contains("execution may still be pending", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Refresh_WithoutTimeout_UsesTheSessionDefault()
    {
        using var batch = ObserveBatchToken(token => Assert.False(token.CanBeCanceled));

        Assert.Throws<ProbeComplete>(() => new PivotTableCommands().Refresh(batch, "Report"));
    }

    [Fact]
    public void Refresh_WhenNativeRefreshReturnsFalse_ReportsFailure()
    {
        int refreshCalls = 0;
        var cache = ExcelInterfaceProxy.Create<Excel.PivotCache>((method, _) => method.Name == "get_RecordCount"
            ? 7 : throw new NotSupportedException(method.Name));
        var pivot = ExcelInterfaceProxy.Create<Excel.PivotTable>((method, _) =>
        {
            if (method.Name == "PivotCache")
                return cache;
            if (method.Name != "RefreshTable")
                throw new NotSupportedException(method.Name);
            refreshCalls++;
            return false;
        });
        var pivots = ExcelInterfaceProxy.Create<Excel.PivotTables>((method, _) => method.Name == "Item"
            ? pivot : throw new NotSupportedException(method.Name));
        var sheet = ExcelInterfaceProxy.Create<Excel.Worksheet>((method, _) => method.Name == "PivotTables"
            ? pivots : throw new NotSupportedException(method.Name));
        var sheets = ExcelInterfaceProxy.Create<Excel.Sheets>((method, _) => method.Name == "get__Default"
            ? sheet : throw new NotSupportedException(method.Name), typeof(ExcelContext).Assembly);
        var workbook = ExcelInterfaceProxy.Create<Excel.Workbook>((method, _) => method.Name == "get_Worksheets"
            ? sheets : throw new NotSupportedException(method.Name), typeof(ExcelContext).Assembly);
        var application = ExcelInterfaceProxy.Create<Excel.Application>((method, _) => method.Name == "get_International"
            ? "General" : throw new NotSupportedException(method.Name), typeof(ExcelContext).Assembly);
        var context = new ExcelContext("fixture.xlsx", application, workbook);
        using var batch = ExcelInterfaceProxy.Create<IExcelBatch>((method, args) => method.Name switch
        {
            nameof(IDisposable.Dispose) => null,
            "get_WorkbookPath" => context.WorkbookPath,
            nameof(IExcelBatch.Execute) => Assert.IsType<Func<ExcelContext, CancellationToken, PivotTableRefreshResult>>(
                args![0])(context, Assert.IsType<CancellationToken>(args[1])),
            _ => throw new NotSupportedException(method.Name)
        });

        var failure = Assert.Throws<InvalidOperationException>(() =>
            new PivotTableCommands().Refresh(batch, "Report", sheetName: "Sheet"));

        Assert.Equal(1, refreshCalls);
        Assert.Contains("did not refresh", failure.Message, StringComparison.Ordinal);
    }

    private static IExcelBatch ObserveBatchToken(Action<CancellationToken> observe) =>
        ExcelInterfaceProxy.Create<IExcelBatch>((method, args) =>
        {
            if (method.Name == nameof(IDisposable.Dispose))
                return null;
            if (method.Name != nameof(IExcelBatch.Execute))
                throw new NotSupportedException(method.Name);
            observe(Assert.IsType<CancellationToken>(args![1]));
            throw new ProbeComplete();
        });

    private sealed class ProbeComplete : Exception;
}
