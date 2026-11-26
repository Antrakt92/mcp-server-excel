using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "WorkbookLifecycle")]
[Trait("RequiresExcel", "false")]
public sealed class WorkbookOutputValidationTests
{
    [Fact]
    public void SaveAs_RequiresRequestedLocalPathRatherThanJustTheFileName()
    {
        var root = Path.Join(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        var target = Path.Join(root, "requested", "My report.xlsx");
        var other = Path.Join(root, "other", "My report.xlsx");
        var original = Path.Join(root, "original.xlsx");
        var equivalent = Path.Join(root, "requested", ".", "My report.xlsx");
        const string cloudTarget = "https://example.sharepoint.com/documents/My%20report.xlsx";

        Assert.False(WorkbookCommands.SaveAsReachedTarget(target, original, original, other));
        Assert.False(WorkbookCommands.SaveAsReachedTarget(target, target, original, other));
        Assert.True(WorkbookCommands.SaveAsReachedTarget(target, original, original, target));
        Assert.False(WorkbookCommands.SaveAsReachedTarget(target, equivalent, equivalent, equivalent));
        Assert.True(WorkbookCommands.SaveAsReachedTarget(target, original, original, equivalent));
        Assert.True(WorkbookCommands.SaveAsReachedTarget(target, original, original, cloudTarget));
        Assert.False(WorkbookCommands.SaveAsReachedTarget(target, original, cloudTarget, cloudTarget));
    }

    [Theory]
    [InlineData("save-as", ".xlsx")]
    [InlineData("save-copy-as", ".xlsx")]
    [InlineData("export-fixed-format", ".pdf")]
    [InlineData("export-image", ".png")]
    public void WorkbookAndChartOutputs_RejectRelativePathsBeforeStartingExcelWork(string operation, string extension)
    {
        var workbook = new WorkbookCommands();
        var chart = new ChartCommands();
        var fileName = Guid.NewGuid().ToString("N") + extension;
        Action execute = operation switch
        {
            "save-as" => () => workbook.SaveAs(null!, fileName),
            "save-copy-as" => () => workbook.SaveCopyAs(null!, fileName),
            "export-fixed-format" => () => workbook.ExportFixedFormat(null!, fileName),
            "export-image" => () => chart.ExportImage(null!, "Chart", fileName),
            _ => throw new ArgumentOutOfRangeException(nameof(operation))
        };

        var error = Assert.Throws<ArgumentException>(execute);
        Assert.Contains("absolute Windows path", error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void CrossWorkbookSheetTransfers_RejectRelativePathsBeforeOpeningFiles(bool move, bool relativeSource)
    {
        var sheet = new SheetCommands();
        var sourceName = Guid.NewGuid().ToString("N") + ".xlsx";
        var targetName = Guid.NewGuid().ToString("N") + ".xlsx";
        var source = relativeSource ? sourceName : Path.Join(Path.GetTempPath(), sourceName);
        var target = relativeSource ? Path.Join(Path.GetTempPath(), targetName) : targetName;

        Action execute = move
            ? () => sheet.MoveToFile(source, "Sheet1", target)
            : () => sheet.CopyToFile(source, "Sheet1", target);
        var error = Assert.Throws<ArgumentException>(execute);
        Assert.Contains("absolute Windows path", error.Message, StringComparison.Ordinal);
    }
}
