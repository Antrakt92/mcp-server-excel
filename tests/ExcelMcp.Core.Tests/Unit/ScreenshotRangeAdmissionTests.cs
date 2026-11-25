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
        using var batch = CreateBatch(range, address, CancellationToken.None);

        var error = Assert.Throws<ArgumentException>(() => new ScreenshotCommands().CaptureRange(batch, rangeAddress: address));

        Assert.Equal("rangeAddress", error.ParamName);
        Assert.Contains("one rectangular range", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void CaptureRange_CancellationDuringRangeResolution_DoesNotChangeExcelView()
    {
        const string address = "A1:Z30";
        using var cancellation = new CancellationTokenSource();
        var areas = ExcelInterfaceProxy.Create<Excel.Areas>((method, _) => method.Name == "get_Count"
            ? 1 : throw new NotSupportedException(method.Name));
        var range = ExcelInterfaceProxy.Create<Excel.Range>((method, _) =>
        {
            if (method.Name == "get_Areas")
                return areas;
            Assert.Equal("get_Address", method.Name);
            cancellation.Cancel();
            return "$A$1:$Z$30";
        });
        using var batch = CreateBatch(range, address, cancellation.Token);

        var error = Assert.Throws<OperationCanceledException>(
            () => new ScreenshotCommands().CaptureRange(batch, rangeAddress: address));
        Assert.Equal(cancellation.Token, error.CancellationToken);
    }

    [Fact]
    public void BuildPixelOffsets_CancellationDuringFrozenPaneMeasurement_StopsBeforeNextCell()
    {
        using var cancellation = new CancellationTokenSource();
        int rowsRead = 0;
        var item = ExcelInterfaceProxy.Create<Excel.Range>((method, _) =>
        {
            Assert.Equal("get_Height", method.Name);
            cancellation.Cancel();
            return 0d;
        });
        var items = ExcelInterfaceProxy.Create<Excel.Range>((method, _) =>
        {
            Assert.True(method.Name is "get_Item" or "get__Default", method.Name);
            rowsRead++;
            return item;
        });
        var range = ExcelInterfaceProxy.Create<Excel.Range>((method, _) =>
            method.Name == "get_Rows" ? items : throw new NotSupportedException(method.Name));

        var error = Assert.Throws<OperationCanceledException>(() =>
            ScreenshotCommands.BuildPixelOffsets(range, true, [new CapturePlanner.Segment(1, 3, 0, 0)],
                1, int.MaxValue, cancellation.Token));
        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.Equal(1, rowsRead);
    }

    private static IExcelBatch CreateBatch(Excel.Range range, string address, CancellationToken operationToken)
    {
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
        return ExcelInterfaceProxy.Create<IExcelBatch>((method, args) => method.Name switch
        {
            nameof(IDisposable.Dispose) => null,
            nameof(IExcelBatch.Execute) => Assert.IsType<Func<ExcelContext, CancellationToken, ScreenshotResult>>(
                args![0])(context, operationToken),
            _ => throw new NotSupportedException(method.Name)
        });
    }
}
