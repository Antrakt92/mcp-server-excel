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
public sealed class PivotTableRequiredRefreshTests
{
    private const string WorkbookPath = "fixture.xlsx";

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Grouping_StopsBeforeLookingUpField_WhenRequiredRefreshFails(bool groupDates)
    {
        var fieldLookups = 0;
        var pivot = ExcelInterfaceProxy.Create<Excel.PivotTable>((method, _) =>
        {
            if (method.Name == "RefreshTable")
                return false;
            fieldLookups++;
            throw new NotSupportedException(method.Name);
        });
        var strategy = new RegularPivotTableFieldStrategy();

        var result = groupDates
            ? strategy.GroupByDate(pivot, "Field", DateGroupingInterval.Months, WorkbookPath)
            : strategy.GroupByNumeric(pivot, "Field", null, null, 10, WorkbookPath);

        Assert.False(result.Success);
        Assert.Contains("did not refresh", result.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(0, fieldLookups);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(0)]
    [InlineData(3)]
    public void CreateCalculatedField_DistinguishesFailedPreflightFromPartialMutation(int failedRefresh)
    {
        var refreshCalls = 0;
        var fieldAdds = 0;
        var refreshFailure = new InvalidOperationException("Native refresh could not be read.");
        var field = ExcelInterfaceProxy.Create<Excel.PivotField>((method, _) => method.Name == "get_Caption"
            ? "Margin" : throw new NotSupportedException(method.Name));
        var fields = ExcelInterfaceProxy.Create<Excel.CalculatedFields>((method, _) =>
        {
            Assert.Equal("Add", method.Name);
            fieldAdds++;
            return field;
        });
        var pivot = ExcelInterfaceProxy.Create<Excel.PivotTable>((method, _) => method.Name switch
        {
            "RefreshTable" => Refresh(),
            "CalculatedFields" => fields,
            _ => throw new NotSupportedException(method.Name)
        });

        bool Refresh()
        {
            refreshCalls++;
            if (failedRefresh == 3 && refreshCalls == 2)
                throw refreshFailure;
            return refreshCalls != failedRefresh;
        }

        var strategy = new RegularPivotTableFieldStrategy();
        if (failedRefresh is 2 or 3)
        {
            var error = Assert.Throws<OperationCompletionUnconfirmedException>(() =>
                strategy.CreateCalculatedField(pivot, "Margin", "=Sales-Cost", WorkbookPath));
            Assert.Contains("list-calculated-fields before retrying", error.Message, StringComparison.Ordinal);
            Assert.Same(failedRefresh == 3 ? refreshFailure : null, error.InnerException);
        }
        else
        {
            var result = strategy.CreateCalculatedField(pivot, "Margin", "=Sales-Cost", WorkbookPath);
            Assert.Equal(failedRefresh == 0, result.Success);
            if (failedRefresh == 1)
                Assert.Contains("did not refresh before creating", result.ErrorMessage, StringComparison.Ordinal);
        }
        Assert.Equal(failedRefresh == 1 ? 1 : 2, refreshCalls);
        Assert.Equal(failedRefresh == 1 ? 0 : 1, fieldAdds);
    }

    [Fact]
    public void DeleteCalculatedField_ReportsCompletedDeletionAndFailedRefresh()
    {
        var deleted = false;
        var field = ExcelInterfaceProxy.Create<Excel.PivotField>((method, _) =>
        {
            if (method.Name == "get_Name")
                return "Margin";
            Assert.Equal("Delete", method.Name);
            deleted = true;
            return null;
        });
        var fields = ExcelInterfaceProxy.Create<Excel.CalculatedFields>((method, _) => method.Name switch
        {
            "get_Count" => 1,
            "Item" => field,
            _ => throw new NotSupportedException(method.Name)
        });
        var cache = ExcelInterfaceProxy.Create<Excel.PivotCache>((method, _) => method.Name == "get_OLAP"
            ? false : throw new NotSupportedException(method.Name));
        var pivot = ExcelInterfaceProxy.Create<Excel.PivotTable>((method, _) => method.Name switch
        {
            "get_Name" => "Report",
            "PivotCache" => cache,
            "CalculatedFields" => fields,
            "RefreshTable" => false,
            _ => throw new NotSupportedException(method.Name)
        });
        var pivots = ExcelInterfaceProxy.Create<Excel.PivotTables>((method, _) => method.Name switch
        {
            "get_Count" => 1,
            "Item" => pivot,
            _ => throw new NotSupportedException(method.Name)
        });
        var sheet = ExcelInterfaceProxy.Create<Excel.Worksheet>((method, _) => method.Name switch
        {
            "get_Name" => "Sheet",
            "PivotTables" => pivots,
            _ => throw new NotSupportedException(method.Name)
        });
        var sheets = ExcelInterfaceProxy.Create<Excel.Sheets>((method, _) => method.Name switch
        {
            "get_Count" => 1,
            "get__Default" => sheet,
            _ => throw new NotSupportedException(method.Name)
        }, typeof(ExcelContext).Assembly);
        var workbook = ExcelInterfaceProxy.Create<Excel.Workbook>((method, _) => method.Name == "get_Worksheets"
            ? sheets : throw new NotSupportedException(method.Name), typeof(ExcelContext).Assembly);
        var app = ExcelInterfaceProxy.Create<Excel.Application>((method, _) => method.Name == "get_International"
            ? "General" : throw new NotSupportedException(method.Name), typeof(ExcelContext).Assembly);
        var context = new ExcelContext(WorkbookPath, app, workbook);
        using var batch = ExcelInterfaceProxy.Create<IExcelBatch>((method, args) => method.Name switch
        {
            nameof(IDisposable.Dispose) => null,
            nameof(IExcelBatch.Execute) => Assert.IsType<Func<ExcelContext, CancellationToken, OperationResult>>(
                args![0])(context, Assert.IsType<CancellationToken>(args[1])),
            _ => throw new NotSupportedException(method.Name)
        });

        var failure = Assert.Throws<OperationCompletionUnconfirmedException>(() =>
            new PivotTableCommands().DeleteCalculatedField(batch, "Report", "Margin"));

        Assert.True(deleted);
        Assert.Contains("completed deletion", failure.Message, StringComparison.Ordinal);
        Assert.Contains("list-calculated-fields before retrying", failure.Message, StringComparison.Ordinal);
    }
}
