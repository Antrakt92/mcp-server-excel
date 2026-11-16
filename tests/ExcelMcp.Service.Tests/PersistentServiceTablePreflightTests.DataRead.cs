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

    [Fact]
    public void GetData_MultiTileTablePreservesVisibleRowsAndColumnOrder()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        List<List<object?>> rows = [["Item", "Even"]];
        rows.AddRange(Enumerable.Range(1, 1001).Select(index => new List<object?> { index, index % 2 == 0 }));
        RequireSuccess(_rangeCommands.SetValues(batch, sheet, "A1:B1002", rows));
        RequireSuccess(_tableCommands.Create(batch, sheet, "TiledRead", "A1:B1002", true));
        RequireSuccess(_rangeCommands.SetVisibility(batch, sheet, "A502:A1001", VisibilityAxis.Rows, hidden: true));

        var all = RequireSuccess(_tableCommands.GetData(batch, "TiledRead"));
        var visible = RequireSuccess(_tableCommands.GetData(batch, "TiledRead", visibleOnly: true));

        Assert.Equal(["Item", "Even"], visible.Headers);
        Assert.Equal(1001, all.RowCount);
        Assert.Equal(501, visible.RowCount);
        Assert.Equal(all.Data.Take(500).SelectMany(row => row).Concat(all.Data[^1]), visible.Data.SelectMany(row => row));
        RangeReadAssertions.Number(500, visible.Data[499][0]);
        Assert.True(Assert.IsType<bool>(visible.Data[499][1]));
        RangeReadAssertions.Number(1001, visible.Data[500][0]);
        Assert.False(Assert.IsType<bool>(visible.Data[500][1]));
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
