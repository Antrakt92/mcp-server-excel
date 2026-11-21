using System.Runtime.InteropServices;
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
public sealed class VisibleSpecialCellsFallbackTests
{
    [Theory]
    [InlineData(true, true, null)]
    [InlineData(false, true, null)]
    [InlineData(true, false, null)]
    [InlineData(false, null, null)]
    [InlineData(true, null, "get_EntireColumn")]
    [InlineData(false, null, "get_Hidden")]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2201",
        Justification = "Exercises the ambiguous native SpecialCells error without launching Excel.")]
    public void GetSpecialCells_UsesOnlyConclusiveHiddenDimension(bool oneColumn, bool? hidden, string? failedMember)
    {
        string address = oneColumn ? "A1:A2" : "A1:B1";
        int rowCount = oneColumn ? 2 : 1;
        int columnCount = oneColumn ? 1 : 2;
        var nativeFailure = new COMException("Synthetic SpecialCells no-match code.", unchecked((int)0x800A03EC));
        var probeFailure = new COMException("Synthetic error reading complete dimension.", unchecked((int)0x800A03EC));
        int cellReads = 0;
        var visibleDimension = Properties<Excel.Range>(("get_Hidden", false));
        var singleDimension = ExcelInterfaceProxy.Create<Excel.Range>((method, _) =>
        {
            Assert.Equal("get_Hidden", method.Name);
            return failedMember == method.Name ? throw probeFailure : hidden;
        });
        var cellRow = Properties<Excel.Range>(("get_Hidden", failedMember is not null));
        var cell = Properties<Excel.Range>(("get_EntireRow", cellRow), ("get_EntireColumn", visibleDimension));
        var cells = ExcelInterfaceProxy.Create<Excel.Range>((method, _) =>
        {
            Assert.Equal("get__Default", method.Name);
            cellReads++;
            return cell;
        });
        Excel.Range? scope = null;
        var sheet = ExcelInterfaceProxy.Create<Excel.Worksheet>((method, _) => method.Name switch
        {
            "get_Name" => "Sheet1",
            "get_Range" => scope,
            _ => throw new NotSupportedException(method.Name)
        });
        var areas = ExcelInterfaceProxy.Create<Excel.Areas>((method, _) => method.Name switch
        {
            "get_Count" => 1,
            "get_Item" or "get__Default" => scope,
            _ => throw new NotSupportedException(method.Name)
        });
        var rows = Properties<Excel.Range>(("get_Count", rowCount));
        var columns = Properties<Excel.Range>(("get_Count", columnCount));
        scope = ExcelInterfaceProxy.Create<Excel.Range>((method, _) => method.Name switch
        {
            "get_Parent" => sheet,
            "get_Address" => address,
            "get_Areas" => areas,
            "get_CountLarge" => 2d,
            "get_Rows" => rows,
            "get_Columns" => columns,
            "get_EntireColumn" when oneColumn => failedMember == method.Name ? throw probeFailure : singleDimension,
            "get_EntireRow" when !oneColumn => singleDimension,
            "get_Cells" => cells,
            "SpecialCells" => throw nativeFailure,
            _ => throw new NotSupportedException(method.Name)
        });
        var sheets = Properties<Excel.Sheets>(("get_Count", 1), ("get__Default", sheet));
        var workbook = Properties<Excel.Workbook>(("get_Worksheets", sheets));
        var app = ExcelInterfaceProxy.Create<Excel.Application>((method, _) => method.Name == "get_International"
            ? "General" : throw new NotSupportedException(method.Name), typeof(ExcelContext).Assembly);
        var context = new ExcelContext("fixture.xlsx", app, workbook);
        using var batch = ExcelInterfaceProxy.Create<IExcelBatch>((method, args) => method.Name switch
        {
            nameof(IDisposable.Dispose) => null,
            "get_WorkbookPath" => context.WorkbookPath,
            nameof(IExcelBatch.Execute) => Assert.IsType<Func<ExcelContext, CancellationToken, SpecialCellsResult>>(
                args![0])(context, Assert.IsType<CancellationToken>(args[1])),
            _ => throw new NotSupportedException(method.Name)
        });

        var commands = new RangeCommands();
        if (hidden is true || failedMember is not null)
        {
            var result = commands.GetSpecialCells(batch, "Sheet1", address, SpecialCellKind.Visible);
            Assert.True(result.Success);
            Assert.Empty(result.Areas);
            Assert.Equal(0, result.CellCount);
            Assert.Equal(failedMember is null ? 0 : 2, cellReads);
        }
        else
        {
            var error = Assert.Throws<COMException>(() => commands.GetSpecialCells(batch, "Sheet1", address, SpecialCellKind.Visible));
            Assert.Same(nativeFailure, error);
            Assert.Equal(1, cellReads);
        }
    }

    private static T Properties<T>(params (string Name, object? Value)[] properties) where T : class =>
        ExcelInterfaceProxy.Create<T>((method, _) => properties.Single(property => property.Name == method.Name).Value);
}
