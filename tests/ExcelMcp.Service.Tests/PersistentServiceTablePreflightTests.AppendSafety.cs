using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceTablePreflightTests
{
    [Fact]
    public void Append_WithTotals_PreservesTotalsAndAddsEveryRow()
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_tableCommands.SetColumnTotal(batch, "SalesTable", "Amount", "sum"));
        RequireSuccess(_rangeCommands.SetValues(batch, "Sales", "A6", [["Fiscal totals"]],
            overwritePolicy: Sbroenne.ExcelMcp.Core.Commands.Range.OverwritePolicy.Allow));
        SetValues(batch, "E7", [["Outside table"]]);
        RequireSuccess(_rangeCommands.SetFormulas(batch, "Sales", "B9", [["=100+23"]]));
        var totalsFormula = RequireSuccess(_rangeCommands.GetFormulas(batch, "Sales", "C6")).Formulas[0][0];
        Assert.StartsWith("=", totalsFormula, StringComparison.Ordinal);

        RequireSuccess(_tableCommands.Append(batch, "SalesTable",
            [["North", "One", 500, "2025-04-01"], ["East", "Two", 600, "2025-04-02"]]));

        var table = RequireSuccess(_tableCommands.Read(batch, "SalesTable")).Table;
        Assert.NotNull(table);
        Assert.True(table.ShowTotals);
        Assert.Equal(6, table.RowCount);
        var values = RequireSuccess(_rangeCommands.GetValues(batch, "Sales", "A6:E8")).Values;
        Assert.Equal("One", values[0][1]);
        Assert.Equal("Two", values[1][1]);
        Assert.Equal("Outside table", values[1][4]);
        Assert.Equal("Fiscal totals", values[2][0]);
        Assert.Equal(1900, Convert.ToInt32(values[2][2], CultureInfo.InvariantCulture));
        Assert.Equal(totalsFormula, RequireSuccess(_rangeCommands.GetFormulas(batch, "Sales", "C8")).Formulas[0][0]);
        Assert.Equal("=100+23", RequireSuccess(_rangeCommands.GetFormulas(batch, "Sales", "B9")).Formulas[0][0]);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Append_WhenExpansionIsOccupied_RejectsWithoutChangingCells(bool formulaBlank)
    {
        var batch = _fixture.BatchToken;
        if (formulaBlank)
            RequireSuccess(_rangeCommands.SetFormulas(batch, "Sales", "B7", [["=\"\""]]));
        else
            SetValues(batch, "B7", [["Keep neighboring data"]]);
        var before = GetSourceState("Sales", "A1:E8");
        var calculation = _fixture.ExecuteRawVerification((context, _) => context.App.Calculation);

        var response = await _fixture.SendForFailureAsync("table.append", new
        {
            tableName = "SalesTable",
            rows = new object[][] { ["North", "One", 500, "2025-04-01"], ["East", "Two", 600, "2025-04-02"] }
        });

        Assert.Contains("occupied", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        AssertPreflightPreserved(before, "A1:E8");
        Assert.Equal(calculation, _fixture.ExecuteRawVerification((context, _) => context.App.Calculation));
    }

    [Fact]
    public async Task Append_BeyondWorksheetEdge_LeavesTheLastRowEmpty()
    {
        var batch = _fixture.BatchToken;
        SetValues(batch, "A1048574:B1048575", [["Item", "Value"], ["Original", 1]]);
        RequireSuccess(_tableCommands.Create(batch, "Sales", "BottomTable", "A1048574:B1048575", true));
        var before = GetSourceState("Sales", "A1048574:B1048576");

        var response = await _fixture.SendForFailureAsync("table.append", new
        {
            tableName = "BottomTable",
            rows = new object[][] { ["First", 2], ["Second", 3] }
        });

        var after = GetSourceState("Sales", "A1048574:B1048576");
        Assert.Equal(before.Formulas.Cast<object>(), after.Formulas.Cast<object>());
        Assert.Equal(before.Formats, after.Formats);
        Assert.Contains("worksheet", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(1, RequireSuccess(_tableCommands.Read(batch, "BottomTable")).Table!.RowCount);
    }

    [Fact]
    public void Append_EmptyTable_AddsRowsWithoutMovingOutsideCells()
    {
        var batch = _fixture.BatchToken;
        SetValues(batch, "F1:G2", [["Item", "Value"], ["Remove", 1]]);
        RequireSuccess(_tableCommands.Create(batch, "Sales", "EmptyAppend", "F1:G2", true));
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.ListObjects? tables = null;
            Excel.ListObject? table = null;
            Excel.ListRows? rows = null;
            Excel.ListRow? row = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, "Sales");
                Assert.NotNull(sheet);
                tables = sheet.ListObjects;
                table = tables["EmptyAppend"];
                rows = table.ListRows;
                row = rows[1];
                row.Delete();
                Assert.Equal(0, rows.Count);
            }
            finally
            {
                ComUtilities.Release(ref row);
                ComUtilities.Release(ref rows);
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref tables);
                ComUtilities.Release(ref sheet);
            }
        });
        RequireSuccess(_rangeCommands.SetFormulas(batch, "Sales", "F4", [["=100+23"]]));
        Assert.Equal(2, RequireSuccess(_tableCommands.Read(batch, "EmptyAppend")).Table!.Columns.Count);

        RequireSuccess(_tableCommands.Append(batch, "EmptyAppend", [["First", 10], ["Second", 20]]));

        var result = RequireSuccess(_tableCommands.GetData(batch, "EmptyAppend"));
        Assert.Equal(["Item", "Value"], result.Headers);
        Assert.Equal(2, result.RowCount);
        Assert.Equal("First", result.Data[0][0]);
        Assert.Equal("Second", result.Data[1][0]);
        Assert.Equal(10, Convert.ToInt32(result.Data[0][1], CultureInfo.InvariantCulture));
        Assert.Equal(20, Convert.ToInt32(result.Data[1][1], CultureInfo.InvariantCulture));
        Assert.Equal("=100+23", RequireSuccess(_rangeCommands.GetFormulas(batch, "Sales", "F4")).Formulas[0][0]);
    }

    [Fact]
    public void Append_PreservesExistingCalculatedColumnAndWritesRequestedFormula()
    {
        var batch = _fixture.BatchToken;
        SetValues(batch, "F1:G3", [["Units", "DoubleUnits"], [3, null], [5, null]]);
        RequireSuccess(_tableCommands.Create(batch, "Sales", "CalculatedAppend", "F1:G3", true));
        const string formula = "=[@Units]*2";
        RequireSuccess(_rangeCommands.SetFormulas(batch, "Sales", "G2:G3", [[formula], [formula]]));

        RequireSuccess(_tableCommands.Append(batch, "CalculatedAppend", [[7, formula], [11, null]]));

        var actual = RequireSuccess(_rangeCommands.GetFormulas(batch, "Sales", "G2:G5"));
        Assert.Equal([formula, formula, formula, string.Empty], actual.Formulas.Select(row => row[0]));
        Assert.Equal([6, 10, 14, null], actual.Values.Select(row =>
            row[0] is null ? (int?)null : Convert.ToInt32(row[0], CultureInfo.InvariantCulture)));
    }
}
