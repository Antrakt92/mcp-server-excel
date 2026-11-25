using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    internal static ExcelErrorMapper.ExcelErrorInfo? ReadNativeError(
        ExcelContext context, Excel.Range range, int row, int column, object? value) =>
        ExcelErrorMapper.TryGet(value, out _, out var error) && IsNativeError(context, range, row, column)
            ? error
            : null;

    internal static ExcelErrorMapper.ExcelErrorInfo? ReadNativeError(
        ExcelContext context, Excel.Range range, int row, int column, object? value,
        ref Excel.Range? cells, ref Excel.WorksheetFunction? functions) =>
        ExcelErrorMapper.TryGet(value, out _, out var error)
            && IsNativeError(context, range, row, column, ref cells, ref functions)
            ? error
            : null;

    // Numeric values can equal error HRESULTs; confirm the cell's native type before mapping.
    private static bool IsNativeError(ExcelContext context, Excel.Range range, int row, int column)
    {
        Excel.Range? cells = null;
        Excel.WorksheetFunction? functions = null;
        try
        {
            return IsNativeError(context, range, row, column, ref cells, ref functions);
        }
        finally
        {
            ComUtilities.Release(ref functions);
            ComUtilities.Release(ref cells);
        }
    }

    private static bool IsNativeError(ExcelContext context, Excel.Range range, int row, int column,
        ref Excel.Range? cells, ref Excel.WorksheetFunction? functions)
    {
        Excel.Range? cell = null;
        try
        {
            cells ??= range.Cells;
            cell = cells[row, column];
            functions ??= context.App.WorksheetFunction;
            return functions.IsError(cell);
        }
        finally
        {
            ComUtilities.Release(ref cell);
        }
    }
}
