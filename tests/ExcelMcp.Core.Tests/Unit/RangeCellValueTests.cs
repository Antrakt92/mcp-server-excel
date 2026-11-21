using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "false")]
public sealed class RangeCellValueTests
{
    [Theory]
    [InlineData("1e400")]
    [InlineData("-1e400")]
    public void ConvertToCellValue_RejectsJsonOverflow(string json)
    {
        using var document = JsonDocument.Parse(json);

        Assert.Throws<ArgumentException>(() => RangeHelpers.ConvertToCellValue(document.RootElement));
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    [InlineData(float.NaN)]
    [InlineData(float.PositiveInfinity)]
    [InlineData(float.NegativeInfinity)]
    public void ConvertToCellValue_RejectsNonFiniteNumbers(object value) =>
        Assert.Throws<ArgumentException>(() => RangeHelpers.ConvertToCellValue(value));

    [Fact]
    public void ConvertToCellValue_PreservesFiniteValuesAndLiteralNumberText()
    {
        using var whole = JsonDocument.Parse("42");
        using var fraction = JsonDocument.Parse("125e-2");

        Assert.Equal(42d, Assert.IsType<double>(RangeHelpers.ConvertToCellValue(whole.RootElement)));
        Assert.Equal(1.25, Assert.IsType<double>(RangeHelpers.ConvertToCellValue(fraction.RootElement)));
        Assert.Equal(1.25f, Assert.IsType<float>(RangeHelpers.ConvertToCellValue(1.25f)));
        Assert.Equal("'1e400", RangeHelpers.ConvertToCellValue("1e400", literalText: true));
    }
}
