using Sbroenne.ExcelMcp.Core.Commands.Drawing;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Drawing")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class DrawingObjectUpdateValidationTests
{
    [Theory]
    [InlineData("width", 1e300)]
    [InlineData("width", 1e-100)]
    [InlineData("height", 1e-100)]
    [InlineData("left", double.NaN)]
    [InlineData("top", double.NegativeInfinity)]
    [InlineData("rotation", 1e100)]
    [InlineData("lineWeight", 1e100)]
    [InlineData("lineWeight", -1d)]
    [InlineData("fontSize", double.PositiveInfinity)]
    public void UpdateObject_RejectsInvalidNumbersBeforeWorkbookExecution(string parameter, double value)
    {
        var exception = Assert.ThrowsAny<ArgumentException>(() =>
            new DrawingCommands().UpdateObject(null!, "Sheet1", "Group",
                newName: "MustNotBeApplied",
                left: parameter == "left" ? value : 30,
                top: parameter == "top" ? value : null,
                width: parameter == "width" ? value : null,
                height: parameter == "height" ? value : null,
                rotation: parameter == "rotation" ? value : null,
                lineWeight: parameter == "lineWeight" ? value : null,
                fontSize: parameter == "fontSize" ? value : null));

        Assert.Equal(parameter, exception.ParamName);
    }
}
