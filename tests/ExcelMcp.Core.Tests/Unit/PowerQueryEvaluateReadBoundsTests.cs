using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PowerQuery")]
public sealed class PowerQueryEvaluateReadBoundsTests
{
    [Theory]
    [InlineData(100_000, 1, 0, "4 MiB response byte budget", true)]
    [InlineData(100_001, 10, 0, "cell read limit", true)]
    [InlineData(1, 1, 240_000, "4 MiB response byte budget", false)]
    [InlineData(0, 0, 0, null, true)]
    public void Evaluate_ChecksResultLimitsAndCleansTemporaryObjects(
        int rowCount, int columnCount, int extraSourceLength, string? limit, bool refreshExpected)
    {
        var events = new List<string>();
        object? Record(string action)
        {
            events.Add(action);
            return null;
        }

        var rows = ExcelInterfaceProxy.Create<Excel.Range>((method, _) =>
            method.Name == "get_Count" ? rowCount : throw new NotSupportedException(method.Name));
        var columns = ExcelInterfaceProxy.Create<Excel.Range>((method, _) =>
            method.Name == "get_Count" ? columnCount : throw new NotSupportedException(method.Name));
        var dataRange = ExcelInterfaceProxy.Create<Excel.Range>((method, _) => method.Name switch
        {
            "get_Rows" => rows,
            "get_Columns" => columns,
            _ => throw new InvalidOperationException($"Unexpected native read {method.Name}.")
        });
        var queryTable = ExcelInterfaceProxy.Create<Excel.QueryTable>((method, _) =>
        {
            if (method.Name == "Refresh")
            {
                events.Add("refresh");
                return true;
            }
            return method.Name.StartsWith("set_", StringComparison.Ordinal)
                ? null : throw new NotSupportedException(method.Name);
        });
        var table = ExcelInterfaceProxy.Create<Excel.ListObject>((method, _) => method.Name switch
        {
            "get_QueryTable" => queryTable,
            "get_HeaderRowRange" => null,
            "get_DataBodyRange" => rowCount == 0 ? null : dataRange,
            "Delete" => Record("delete-table"),
            _ => throw new NotSupportedException(method.Name)
        });
        var tables = ExcelInterfaceProxy.Create<Excel.ListObjects>((method, _) =>
            method.Name == "Add" ? table : throw new NotSupportedException(method.Name));
        var sheet = ExcelInterfaceProxy.Create<Excel.Worksheet>((method, _) => method.Name switch
        {
            "set_Name" => null,
            "get_Range" => dataRange,
            "get_ListObjects" => tables,
            "Delete" => Record("delete-sheet"),
            _ => throw new NotSupportedException(method.Name)
        });
        var sheets = ExcelInterfaceProxy.Create<Excel.Sheets>((method, _) =>
            method.Name == "Add" ? sheet : throw new NotSupportedException(method.Name));
        var query = ExcelInterfaceProxy.Create<Excel.WorkbookQuery>((method, _) =>
            method.Name == "Delete" ? Record("delete-query") : throw new NotSupportedException(method.Name));
        var queries = ExcelInterfaceProxy.Create<Excel.Queries>((method, _) =>
            method.Name == "Add" ? query : throw new NotSupportedException(method.Name));
        var connections = ExcelInterfaceProxy.Create<Excel.Connections>((method, _) =>
            method.Name == "get_Count" ? 0 : throw new NotSupportedException(method.Name),
            typeof(PowerQueryCommands).Assembly);
        var app = ExcelInterfaceProxy.Create<Excel.Application>((method, _) => method.Name switch
        {
            "get_International" => "General",
            "get_DisplayAlerts" => true,
            "set_DisplayAlerts" => null,
            _ => throw new NotSupportedException(method.Name)
        }, typeof(ExcelContext).Assembly);
        var book = ExcelInterfaceProxy.Create<Excel.Workbook>((method, _) => method.Name switch
        {
            "get_Queries" => queries,
            "get_Application" => app,
            "get_Worksheets" => sheets,
            "get_Connections" => connections,
            _ => throw new NotSupportedException(method.Name)
        }, typeof(ExcelContext).Assembly);
        var context = new ExcelContext("fixture.xlsx", app, book);
        using var batch = ExcelInterfaceProxy.Create<IExcelBatch>((method, args) => method.Name switch
        {
            nameof(IDisposable.Dispose) => null,
            "get_WorkbookPath" => context.WorkbookPath,
            "get_OperationTimeout" => TimeSpan.FromSeconds(10),
            nameof(IExcelBatch.Execute) => Assert.IsType<Func<ExcelContext, CancellationToken, PowerQueryEvaluateResult>>(
                args![0])(context, Assert.IsType<CancellationToken>(args[1])),
            _ => throw new NotSupportedException(method.Name)
        });

        var commands = new PowerQueryCommands(new DataModelCommands());
        var code = "let Source = #table({\"Value\"}, {{1}}) in Source // " + new string('x', extraSourceLength);
        if (limit is null)
        {
            var result = commands.Evaluate(batch, code);
            Assert.True(result.Success);
            Assert.Equal(code, result.MCode);
            Assert.Empty(result.Rows);
            Assert.Equal(0, result.RowCount);
        }
        else
        {
            var error = Assert.Throws<ArgumentOutOfRangeException>(() => commands.Evaluate(batch, code));
            Assert.Contains(limit, error.Message, StringComparison.Ordinal);
            Assert.Contains("Table.FirstN", error.Message, StringComparison.Ordinal);
        }

        if (refreshExpected)
            Assert.Equal(["refresh", "delete-table", "delete-sheet", "delete-query"], events);
        else
            Assert.Empty(events);
    }
}
