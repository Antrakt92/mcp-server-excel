using System.Reflection;
using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
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
            BoundedRangeValueReader.Read(CreateContext(), range, (rows, columns), new RangeReadBudget(),
                CancellationToken.None));
        Assert.Contains(expected, rejected.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void OversizedVisibleRead_StopsCheckingRowVisibilityWhenItsBudgetIsExhausted()
    {
        var range = ExcelInterfaceProxy.Create<Excel.Range>((method, _) => throw new NotSupportedException(method.Name));
        var budget = new RangeReadBudget();
        long available = RangeReadBudget.MaximumResponseBytes - budget.EstimatedBytes;
        var sample = new RangeReadBudget();
        long initial = sample.EstimatedBytes;
        sample.ReserveGrid(1, 1, includeFormulas: false);
        long bytesPerRow = sample.EstimatedBytes - initial;
        int checkedRows = 0;

        var rejected = Assert.Throws<ArgumentOutOfRangeException>(() =>
            BoundedRangeValueReader.Read(CreateContext(), range, (100_000, 1), budget, CancellationToken.None,
                includeRow: _ => ++checkedRows % 2 == 0));

        Assert.Contains("4 MiB response byte budget", rejected.Message, StringComparison.Ordinal);
        Assert.Equal(2 * ((int)(available / bytesPerRow) + 1), checkedRows);
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

        var result = BoundedRangeValueReader.Read(CreateContext(), range, (1001, 2), new RangeReadBudget(),
            CancellationToken.None, includeRow: row => row is 499 or 1000);

        Assert.Equal([(1, 1, 500, 2), (1001, 1, 1, 2)], tiles);
        Assert.Equal<object?>([500d, true], result[0]);
        Assert.Equal<object?>([1001d, false], result[1]);
    }

    [Fact]
    public void Read_ReusesNativeLookupsAndProbesOnlyMatchingErrorCodes()
    {
        const int errorCode = -2146826281;
        var candidateColumns = new List<int>();
        int cellsRequests = 0;
        int functionRequests = 0;
        var nativeError = ExcelInterfaceProxy.Create<Excel.Range>((method, _) => throw new NotSupportedException(method.Name));
        var numericCollision = ExcelInterfaceProxy.Create<Excel.Range>((method, _) => throw new NotSupportedException(method.Name));
        var functions = ExcelInterfaceProxy.Create<Excel.WorksheetFunction>((method, args) =>
        {
            Assert.Equal("IsError", method.Name);
            return ReferenceEquals(args![0], nativeError);
        });
        var tileCells = ExcelInterfaceProxy.Create<Excel.Range>((method, args) =>
        {
            Assert.Equal("get__Default", method.Name);
            Assert.Equal(1, Assert.IsType<int>(args![0]));
            int column = Assert.IsType<int>(args[1]);
            candidateColumns.Add(column);
            return column == 3 ? numericCollision : nativeError;
        });
        var tile = ExcelInterfaceProxy.Create<Excel.Range>((method, _) => method.Name switch
        {
            "get_Cells" => GetTileCells(),
            "get_Value2" => new object[,] { { 7d, errorCode, (double)errorCode, errorCode } },
            _ => throw new NotSupportedException(method.Name)
        });
        var anchor = ExcelInterfaceProxy.Create<Excel.Range>((method, _) => method.Name == "get_Resize"
            ? tile : throw new NotSupportedException(method.Name));
        var cells = ExcelInterfaceProxy.Create<Excel.Range>((method, _) => method.Name == "get__Default"
            ? anchor : throw new NotSupportedException(method.Name));
        var range = ExcelInterfaceProxy.Create<Excel.Range>((method, _) => method.Name == "get_Cells"
            ? cells : throw new NotSupportedException(method.Name));
        var context = CreateContext(() =>
        {
            functionRequests++;
            return functions;
        });

        var result = BoundedRangeValueReader.Read(context, range, (1, 4), new RangeReadBudget(),
            CancellationToken.None);

        Assert.Equal<object?>([7d, "#DIV/0!", (double)errorCode, "#DIV/0!"], Assert.Single(result));
        Assert.Equal([2, 3, 4], candidateColumns);
        Assert.Equal(1, cellsRequests);
        Assert.Equal(1, functionRequests);

        Excel.Range GetTileCells()
        {
            cellsRequests++;
            return tileCells;
        }
    }

    private static ExcelContext CreateContext(Func<Excel.WorksheetFunction>? getFunctions = null)
    {
        var app = ExcelInterfaceProxy.Create<Excel.Application>((method, _) => method.Name switch
        {
            "get_WorksheetFunction" when getFunctions is not null => getFunctions(),
            _ => throw new NotSupportedException(method.Name)
        });
        // The reader uses Core's embedded PIA, while ExcelContext normally probes locale through ComInterop's PIA.
        var context = (ExcelContext)RuntimeHelpers.GetUninitializedObject(typeof(ExcelContext));
        typeof(ExcelContext).GetField("<App>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic)!
            .SetValue(context, app);
        return context;
    }
}
