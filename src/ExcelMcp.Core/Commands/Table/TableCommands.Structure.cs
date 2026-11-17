using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table structure operations (Resize, ToggleTotals, SetColumnTotal, SetStyle)
/// </summary>
public partial class TableCommands
{
    /// <inheritdoc />
    public OperationResult Resize(IExcelBatch batch, string tableName, string newRange)
    {
        ValidateRequiredTableName(tableName);

        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? newRangeObj = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                sheet = (Excel.Worksheet)table.Parent;
                newRangeObj = sheet.Range[newRange];

                // Resize the table
                table.Resize(newRangeObj);

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref newRangeObj);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult ToggleTotals(IExcelBatch batch, string tableName, bool showTotals)
    {
        ValidateRequiredTableName(tableName);

        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                table.ShowTotals = showTotals;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetColumnTotal(IExcelBatch batch, string tableName, string columnName, string totalFunction)
    {
        ValidateRequiredTableName(tableName);
        ArgumentException.ThrowIfNullOrWhiteSpace(totalFunction);
        var calculation = totalFunction.ToLowerInvariant() switch
        {
            "sum" => Excel.XlTotalsCalculation.xlTotalsCalculationSum,
            "average" or "avg" => Excel.XlTotalsCalculation.xlTotalsCalculationAverage,
            "count" => Excel.XlTotalsCalculation.xlTotalsCalculationCount,
            "countnums" => Excel.XlTotalsCalculation.xlTotalsCalculationCountNums,
            "max" => Excel.XlTotalsCalculation.xlTotalsCalculationMax,
            "min" => Excel.XlTotalsCalculation.xlTotalsCalculationMin,
            "stddev" => Excel.XlTotalsCalculation.xlTotalsCalculationStdDev,
            "var" => Excel.XlTotalsCalculation.xlTotalsCalculationVar,
            "none" => Excel.XlTotalsCalculation.xlTotalsCalculationNone,
            _ => throw new ArgumentException(
                $"Unknown total function '{totalFunction}'. Valid: sum, average, count, countnums, max, min, stddev, var, none",
                nameof(totalFunction))
        };

        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            Excel.ListColumns? listColumns = null;
            Excel.ListColumn? column = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                listColumns = table.ListColumns;
                int columnCount = listColumns.Count;
                for (int i = 1; i <= columnCount; i++)
                {
                    Excel.ListColumn? col = null;
                    try
                    {
                        col = listColumns[i];
                        if (col.Name == columnName)
                        {
                            column = col;
                            col = null;
                            break;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref col);
                    }
                }

                if (column == null)
                {
                    throw new InvalidOperationException($"Column '{columnName}' not found in table '{tableName}'");
                }

                if (!table.ShowTotals)
                {
                    table.ShowTotals = true;
                }
                column.TotalsCalculation = calculation;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref listColumns);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetStyle(IExcelBatch batch, string tableName, string tableStyle)
    {
        ValidateRequiredTableName(tableName);

        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                table.TableStyle = tableStyle;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref table);
            }
        });
    }
}
