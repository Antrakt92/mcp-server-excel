using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "Range")]
public sealed class RangeValidationPreflightTests
{
    [Theory]
    [InlineData("list", null, "ignored")]
    [InlineData("custom", "", "ignored")]
    [InlineData("whole", null, "")]
    [InlineData("decimal", null, null)]
    [InlineData("date", null, null)]
    [InlineData("time", null, null)]
    [InlineData("textLength", null, null)]
    public void ValidateRange_MissingRequiredFormula_RejectsBeforeExecutingExcel(
        string type, string? formula1, string? formula2)
    {
        using var batch = ExcelInterfaceProxy.Create<IExcelBatch>((method, _) =>
            method.Name == nameof(IDisposable.Dispose)
                ? null
                : throw new InvalidOperationException($"Excel batch was accessed through {method.Name}."));

        var failure = Assert.Throws<ArgumentException>(() => ValidateRange(batch, type, formula1, formula2));

        Assert.Equal("formula1", failure.ParamName);
    }

    [Theory]
    [InlineData("any", null, null)]
    [InlineData("whole", null, "10")]
    [InlineData("list", "=CategoryList", null)]
    public void ValidateRange_AcceptedArgumentShape_ReachesExcel(
        string type, string? formula1, string? formula2)
    {
        var expected = new InvalidOperationException("Reached Excel execution.");
        using var batch = ExcelInterfaceProxy.Create<IExcelBatch>((method, _) =>
            method.Name == nameof(IDisposable.Dispose) ? null : throw expected);

        var failure = Assert.Throws<InvalidOperationException>(() => ValidateRange(batch, type, formula1, formula2));

        Assert.Same(expected, failure);
    }

    private static void ValidateRange(IExcelBatch batch, string type, string? formula1, string? formula2) =>
        new RangeCommands().ValidateRange(batch, "Inputs", "A1", type, null, formula1, formula2,
            null, null, null, null, null, null, null, null, null);
}
