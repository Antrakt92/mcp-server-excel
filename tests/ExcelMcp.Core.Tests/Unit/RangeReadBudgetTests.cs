using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "false")]
public sealed class RangeReadBudgetTests
{
    [Theory]
    [InlineData("plain")]
    [InlineData("\"\\\t\r\n\u0001")]
    [InlineData("é😀")]
    public void ScalarEstimate_CoversJsonEscapingAndUnicode(string text)
    {
        long estimate = RangeReadBudget.EstimateScalar(text);
        Assert.True(estimate >= JsonSerializer.SerializeToUtf8Bytes(text).LongLength);
        Assert.Equal(2L + 6L * text.Length, estimate);
    }

    [Fact]
    public void LongStrings_ExceedBudgetWithoutExposingCellContent()
    {
        var budget = new RangeReadBudget();
        budget.ReserveGrid(1, 32, false);
        string text = new('x', 32_000);
        var failure = Assert.Throws<ArgumentOutOfRangeException>(() =>
        {
            for (int index = 0; index < 32; index++) budget.AddScalar(text);
        });
        Assert.Contains("byte budget", failure.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(text, failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void GridOverhead_IsAdmittedBeforeAllocatingLargeRowLists()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new RangeReadBudget().ReserveGrid(1_000_000, 1, false));
        var budget = new RangeReadBudget();
        budget.ReserveGrid(10, 10, true);
        long before = budget.EstimatedBytes;
        budget.AddScalar("value");
        budget.AddScalar("=A1");
        budget.AddError("=A1");
        Assert.True(budget.EstimatedBytes > before + RangeReadBudget.EstimateScalar("value") + RangeReadBudget.EstimateScalar("=A1"));
    }

    [Theory]
    [InlineData(31, 37)]
    [InlineData(2, 1501)]
    [InlineData(2001, 1)]
    [InlineData(1, 1)]
    public void Tiles_CoverEveryCellExactlyOnceWithinTheNativeReadBound(int rows, int columns)
    {
        var visited = new HashSet<(int Row, int Column)>();
        var tiles = RangeReadWindow.Tiles(rows, columns).ToArray();
        Assert.NotEmpty(tiles);
        foreach (var tile in tiles)
        {
            Assert.InRange(tile.Rows * tile.Columns, 1, RangeReadWindow.MaximumTileCells);
            Assert.InRange(tile.RowOffset + tile.Rows, 1, rows);
            Assert.InRange(tile.ColumnOffset + tile.Columns, 1, columns);
            for (int row = 0; row < tile.Rows; row++)
                for (int column = 0; column < tile.Columns; column++)
                    Assert.True(visited.Add((tile.RowOffset + row, tile.ColumnOffset + column)));
        }
        Assert.Equal(rows * columns, visited.Count);
    }
}
