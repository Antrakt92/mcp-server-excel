using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "false")]
public sealed class BoundedRangeValueReaderTests
{
    [Theory]
    [InlineData(1_048_576, 16_384, "cell read limit")]
    [InlineData(100_000, 1, "4 MiB response byte budget")]
    public void OversizedRead_IsRejectedBeforeRequestingAnyNativeValues(int rows, int columns, string expected)
    {
        var range = ExcelInterfaceProxy.Create<Excel.Range>((method, _) => throw new NotSupportedException(method.Name));
        var rejected = Assert.Throws<ArgumentOutOfRangeException>(() =>
            BoundedRangeValueReader.Read(range, (rows, columns), new RangeReadBudget(), CancellationToken.None));
        Assert.Contains(expected, rejected.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void VisibleRead_SkipsEntireHiddenTileAndPreservesRowsColumnsAndNativeTypes()
    {
        var tiles = new List<(int Row, int Column, int Rows, int Columns)>();
        var cells = ExcelInterfaceProxy.Create<Excel.Range>((method, args) =>
        {
            Assert.Equal("get__Default", method.Name);
            int row = Assert.IsType<int>(args![0]);
            int column = Assert.IsType<int>(args[1]);
            return ExcelInterfaceProxy.Create<Excel.Range>((resize, dimensions) =>
            {
                Assert.Equal("get_Resize", resize.Name);
                int height = Assert.IsType<int>(dimensions![0]);
                int width = Assert.IsType<int>(dimensions[1]);
                return ExcelInterfaceProxy.Create<Excel.Range>((value, _) =>
                {
                    Assert.Equal("get_Value2", value.Name);
                    Assert.InRange(height * width, 1, 1_000);
                    tiles.Add((row, column, height, width));
                    var result = new object?[height, width];
                    for (int index = 0; index < height; index++)
                    {
                        result[index, 0] = (double)(row + index);
                        result[index, 1] = row == 1;
                    }
                    return result;
                });
            });
        });
        var range = ExcelInterfaceProxy.Create<Excel.Range>((method, _) => method.Name == "get_Cells"
            ? cells : throw new NotSupportedException(method.Name));

        var result = BoundedRangeValueReader.Read(range, (1001, 2), new RangeReadBudget(), CancellationToken.None,
            includeRow: row => row is 499 or 1000);

        Assert.Equal([(1, 1, 500, 2), (1001, 1, 1, 2)], tiles);
        Assert.Equal<object?>([500d, true], result[0]);
        Assert.Equal<object?>([1001d, false], result[1]);
    }
}
