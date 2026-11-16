using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Tables")]
[Trait("RequiresExcel", "false")]
public sealed class TableAppendDisconnectTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2201",
        Justification = "Synthetic RPC disconnection drives the native append boundary without starting or terminating Excel.")]
    public void Append_Disconnect_MarksOnlyThePhasesExcelReached(bool duringWrite)
    {
        var disconnect = new COMException("Synthetic native disconnect.", ResiliencePipelines.RPC_E_DISCONNECTED);
        var resizedRange = Properties<Excel.Range>();
        var destination = ExcelInterfaceProxy.Create<Excel.Range>((method, _) =>
            method.Name == "set_Value2" ? throw disconnect : throw new NotSupportedException(method.Name));
        var firstNewCell = Properties<Excel.Range>(("get_Row", 3), ("get_Column", 1), ("get_Resize", destination));
        var source = Properties<Excel.Range>(
            ("get_Rows", CountRange(2)), ("get_Columns", CountRange(1)),
            ("get_Row", 1), ("get_Column", 1), ("get_Resize", resizedRange));
        var body = Properties<Excel.Range>(
            ("get_Rows", CountRange(2)), ("get_Columns", CountRange(1)),
            ("get_Cells", Properties<Excel.Range>(("get__Default", firstNewCell))));
        var expansion = Properties<Excel.Range>(("get_MergeCells", false));
        var anchor = Properties<Excel.Range>(("get_Resize", expansion));
        var sheetCells = Properties<Excel.Range>(("get__Default", anchor));
        bool resizeReturned = false;
        var listRows = ExcelInterfaceProxy.Create<Excel.ListRows>((method, _) =>
            method.Name == "get_Count" ? duringWrite && resizeReturned ? 2 : 1 : throw new NotSupportedException(method.Name));
        Excel.ListObjects? tables = null;
        var sheet = ExcelInterfaceProxy.Create<Excel.Worksheet>((method, _) => method.Name switch
        {
            "get_ListObjects" => tables,
            "get_Cells" => sheetCells,
            "get_Rows" => CountRange(100),
            _ => throw new NotSupportedException(method.Name)
        });
        var table = ExcelInterfaceProxy.Create<Excel.ListObject>((method, _) => method.Name switch
        {
            "get_Name" => "Stock",
            "get_Parent" => sheet,
            "get_Range" => source,
            "get_DataBodyRange" => body,
            "get_ShowTotals" => false,
            "get_ListRows" => listRows,
            "Resize" when !duringWrite => throw disconnect,
            "Resize" => RecordResize(),
            _ => throw new NotSupportedException(method.Name)
        });
        tables = Properties<Excel.ListObjects>(("get_Count", 1), ("get__Default", table));
        var sheets = ExcelInterfaceProxy.Create<Excel.Sheets>((method, _) => method.Name switch
        {
            "get_Count" => 1,
            "get__Default" => sheet,
            _ => throw new NotSupportedException(method.Name)
        });
        var workbook = ExcelInterfaceProxy.Create<Excel.Workbook>((method, _) => method.Name == "get_Worksheets"
            ? sheets : throw new NotSupportedException(method.Name));
        var worksheetFunctions = Properties<Excel.WorksheetFunction>(("CountA", 0d));
        var application = ExcelInterfaceProxy.Create<Excel.Application>((method, _) => method.Name switch
        {
            "get_International" => "General",
            "get_WorksheetFunction" => worksheetFunctions,
            _ => throw new NotSupportedException(method.Name)
        });
        // ExcelContext normally probes locale through a different embedded PIA.
        // These callbacks only need the Excel objects and do not inspect locale.
        var context = (ExcelContext)RuntimeHelpers.GetUninitializedObject(typeof(ExcelContext));
        const BindingFlags fields = BindingFlags.Instance | BindingFlags.NonPublic;
        typeof(ExcelContext).GetField("<App>k__BackingField", fields)!.SetValue(context, application);
        typeof(ExcelContext).GetField("<Book>k__BackingField", fields)!.SetValue(context, workbook);
        typeof(ExcelContext).GetField("<WorkbookPath>k__BackingField", fields)!.SetValue(context, "fixture.xlsx");
        using var batch = ExcelInterfaceProxy.Create<IExcelBatch>((method, args) => method.Name switch
        {
            nameof(IDisposable.Dispose) => null,
            "get_WorkbookPath" => context.WorkbookPath,
            nameof(IExcelBatch.Execute) => Assert.IsType<Func<ExcelContext, CancellationToken, OperationResult>>(args![0])(
                context, Assert.IsType<CancellationToken>(args[1])),
            _ => throw new NotSupportedException(method.Name)
        });

        var error = Assert.Throws<TableAppendException>(() => new TableCommands().Append(batch, "Stock", [[123]]));

        Assert.Same(disconnect, error.InnerException);
        Assert.Equal(new TableAppendPartialState("Stock", 1, 2, duringWrite, duringWrite), error.PartialState);

        object? RecordResize()
        {
            resizeReturned = true;
            return null;
        }
    }

    private static Excel.Range CountRange(int count) => Properties<Excel.Range>(("get_Count", count));

    private static T Properties<T>(params (string Member, object? Value)[] values) where T : class =>
        ExcelInterfaceProxy.Create<T>((method, _) =>
        {
            foreach (var (member, value) in values)
                if (method.Name == member)
                    return value;
            throw new NotSupportedException(method.Name);
        });
}
