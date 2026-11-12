using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceTablePreflightTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GetData_BlankSingletonKeepsTheExistingRow(bool visibleOnly)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_rangeCommands.SetValues(batch, sheet, "A1:A2", [["Item"], ["Clear me"]]));
        RequireSuccess(_tableCommands.Create(batch, sheet, "BlankSingleton", "A1:A2", true));
        RequireSuccess(_rangeCommands.ClearContents(batch, sheet, "A2"));
        Assert.Equal(1, RequireSuccess(_tableCommands.Read(batch, "BlankSingleton")).Table!.RowCount);

        var result = RequireSuccess(_tableCommands.GetData(batch, "BlankSingleton", visibleOnly));

        Assert.Equal(["Item"], result.Headers);
        Assert.Equal(1, result.ColumnCount);
        Assert.Equal(1, result.RowCount);
        Assert.Null(Assert.Single(Assert.Single(result.Data)));
    }

    [Theory]
    [InlineData(null, 4)]
    [InlineData("A3", 3)]
    [InlineData("A2:A5", 0)]
    public void GetData_VisibleRowsStillIncludeHiddenColumns(string? hiddenRows, int expectedCount)
    {
        var batch = _fixture.BatchToken;
        var before = RequireSuccess(_tableCommands.GetData(batch, "SalesTable", visibleOnly: false));
        RequireSuccess(_rangeCommands.SetVisibility(batch, "Sales", "B1", VisibilityAxis.Columns, hidden: true));
        if (hiddenRows is not null)
            RequireSuccess(_rangeCommands.SetVisibility(batch, "Sales", hiddenRows, VisibilityAxis.Rows, hidden: true));

        var visible = RequireSuccess(_tableCommands.GetData(batch, "SalesTable", visibleOnly: true));
        var all = RequireSuccess(_tableCommands.GetData(batch, "SalesTable", visibleOnly: false));

        Assert.Equal(expectedCount, visible.RowCount);
        Assert.Equal(before.Headers, visible.Headers);
        Assert.Equal(4, visible.ColumnCount);
        var expected = hiddenRows switch
        {
            null => before.Data,
            "A3" => before.Data.Where((_, index) => index != 1).ToList(),
            _ => []
        };
        Assert.Equal(expected.SelectMany(row => row), visible.Data.SelectMany(row => row));
        Assert.Equal(before.Data.SelectMany(row => row), all.Data.SelectMany(row => row));
    }
}
