using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Screenshot")]
[Trait("RequiresExcel", "false")]
public sealed class ScreenshotRangeAdmissionTests
{
    [Theory]
    [InlineData("A1:B2,Z25:AA26")]
    [InlineData("CombinedSelection")]
    public void CaptureRange_RejectsResolvedMultiAreaSelectionBeforeRendering(string address)
    {
        var areas = ExcelInterfaceProxy.Create<Excel.Areas>((method, _) => method.Name == "get_Count"
            ? 2 : throw new NotSupportedException(method.Name));
        var range = ExcelInterfaceProxy.Create<Excel.Range>((method, _) => method.Name == "get_Areas"
            ? areas : throw new NotSupportedException(method.Name));
        var sheet = ExcelInterfaceProxy.Create<Excel.Worksheet>((method, args) =>
        {
            if (method.Name == "get_Name")
                return "Sheet1";
            Assert.Equal("get_Range", method.Name);
            Assert.Equal(address, args![0]);
            return range;
        });
        var workbook = ExcelInterfaceProxy.Create<Excel.Workbook>((method, _) => method.Name == "get_ActiveSheet"
            ? sheet : throw new NotSupportedException(method.Name));
        var app = ExcelInterfaceProxy.Create<Excel.Application>((method, _) => method.Name == "get_International"
            ? "General" : throw new NotSupportedException(method.Name), typeof(ExcelContext).Assembly);
        var context = new ExcelContext("fixture.xlsx", app, workbook);
        using var batch = ExcelInterfaceProxy.Create<IExcelBatch>((method, args) => method.Name switch
        {
            nameof(IDisposable.Dispose) => null,
            nameof(IExcelBatch.Execute) => Assert.IsType<Func<ExcelContext, CancellationToken, ScreenshotResult>>(
                args![0])(context, Assert.IsType<CancellationToken>(args[1])),
            _ => throw new NotSupportedException(method.Name)
        });

        var error = Assert.Throws<ArgumentException>(() => new ScreenshotCommands().CaptureRange(batch, rangeAddress: address));

        Assert.Equal("rangeAddress", error.ParamName);
        Assert.Contains("one rectangular range", error.Message, StringComparison.Ordinal);
    }
}
