using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "false")]
public sealed class RangeReadWindowTests
{
    [Fact]
    public void Create_ReturnsClippedPageAndColumnFirstContinuation()
    {
        var first = RangeReadWindow.Create(5, 7, 0, 0, 3, 4);
        Assert.Equal((3, 4, 0, 4), (first.Rows, first.Columns, first.NextRowOffset, first.NextColumnOffset));
        var right = RangeReadWindow.Create(5, 7, 0, 4, 3, 4);
        Assert.Equal((3, 3, 3, 0), (right.Rows, right.Columns, right.NextRowOffset, right.NextColumnOffset));
        var last = RangeReadWindow.Create(5, 7, 3, 4, 3, 4);
        Assert.Equal((2, 3), (last.Rows, last.Columns));
        Assert.Null(last.NextRowOffset);
        Assert.Null(last.NextColumnOffset);
    }

    [Theory]
    [InlineData(-1, 0, 1, 1)]
    [InlineData(0, -1, 1, 1)]
    [InlineData(0, 0, 0, 1)]
    [InlineData(0, 0, 1, 0)]
    [InlineData(0, 0, 101, 100)]
    [InlineData(0, 0, int.MaxValue, int.MaxValue)]
    public void Validate_RejectsInvalidOrOversizedPage(int row, int column, int rows, int columns)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => RangeReadWindow.Validate(row, column, rows, columns));
    }

    [Fact]
    public void Create_RejectsOffsetOutsideScope()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => RangeReadWindow.Create(5, 7, 5, 0, 1, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => RangeReadWindow.Create(5, 7, 0, 7, 1, 1));
    }

    [Fact]
    public void EnsureCellLimit_RejectsHugeDimensionsWithoutIntegerOverflow()
    {
        RangeReadWindow.EnsureCellLimit(1000, 1000, 1_000_000);
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            RangeReadWindow.EnsureCellLimit(1_048_576, 16_384, 1_000_000));
    }
}
