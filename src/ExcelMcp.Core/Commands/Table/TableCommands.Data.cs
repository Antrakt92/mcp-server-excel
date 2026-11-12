using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table data operations (AppendRows)
/// </summary>
public partial class TableCommands
{
    /// <inheritdoc />
    public OperationResult Append(IExcelBatch batch, string tableName, List<List<object?>>? rows = null, string? rowsFile = null)
    {
        ValidateRequiredTableName(tableName);
        var resolvedRows = ParameterTransforms.ResolveValuesOrFile(rows, rowsFile, "rows");

        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? tableRange = null;
            Excel.Range? expandedRange = null;
            Excel.ListRows? listRows = null;
            Excel.Range? body = null;
            Excel.Range? bodyCells = null;
            Excel.Range? firstNewCell = null;
            Excel.Range? destination = null;
            bool expansionAttempted = false;
            bool writeAttempted = false;

            try
            {
                table = FindTable(ctx.Book, tableName);
                sheet = (Excel.Worksheet)table.Parent;
                tableRange = table.Range;
                var (tableHeight, columnCount) = RangeHelpers.GetDimensions(tableRange);
                for (int index = 0; index < resolvedRows.Count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    var row = resolvedRows[index];
                    if (row is null || row.Count != columnCount)
                    {
                        throw new ArgumentException(
                            $"Append row {index + 1} must have {columnCount} columns to match table '{tableName}'.",
                            nameof(rows));
                    }
                }

                listRows = table.ListRows;
                int originalRowCount = listRows.Count;
                bool showTotals = table.ShowTotals;
                // An empty table's Range can include its blank insert row, which ListRows excludes.
                int dataAndHeaderHeight = originalRowCount == 0 && !showTotals
                    ? (table.ShowHeaders ? 1 : 0)
                    : tableHeight;
                int lastTableRow = tableRange.Row + dataAndHeaderHeight - 1;
                int firstColumn = tableRange.Column;
                int rowsToAdd = resolvedRows.Count;
                EnsureAppendSpace(ctx, sheet, lastTableRow, firstColumn, rowsToAdd, columnCount);
                var values = (object[,])Array.CreateInstance(typeof(object), [rowsToAdd, columnCount], [1, 1]);
                for (int row = 0; row < rowsToAdd; row++)
                {
                    ct.ThrowIfCancellationRequested();
                    for (int column = 0; column < columnCount; column++)
                        values[row + 1, column + 1] = RangeHelpers.ConvertToCellValue(resolvedRows[row][column]);
                }

                expandedRange = tableRange.Resize[dataAndHeaderHeight + rowsToAdd, columnCount];
                ct.ThrowIfCancellationRequested();
                expansionAttempted = true;
                table.Resize(expandedRange);

                body = table.DataBodyRange;
                int expectedRows = originalRowCount + rowsToAdd;
                int actualRows = listRows.Count;
                var bodySize = body is null ? (Rows: 0, Columns: 0) : RangeHelpers.GetDimensions(body);
                bool actualTotals = table.ShowTotals;
                if (actualRows != expectedRows || bodySize != (expectedRows, columnCount) || actualTotals != showTotals)
                    throw new InvalidOperationException(
                        $"Excel returned {actualRows} table rows and a {bodySize.Rows} by {bodySize.Columns} data range " +
                        $"(totals visible: {actualTotals}); expected {expectedRows} rows and {columnCount} columns " +
                        $"(totals visible: {showTotals}).");

                bodyCells = body!.Cells;
                firstNewCell = bodyCells[originalRowCount + 1, 1];
                if (firstNewCell.Row != lastTableRow + (showTotals ? 0 : 1) || firstNewCell.Column != firstColumn)
                    throw new InvalidOperationException("Excel moved the new data rows outside the expected append location.");

                destination = firstNewCell.Resize[rowsToAdd, columnCount];
                ct.ThrowIfCancellationRequested();
                writeAttempted = true;
                destination.Value2 = values;
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            catch (Exception ex) when (expansionAttempted && CreatedObjectFailure.CanReport(ex))
            {
                string state = writeAttempted
                    ? "Writing the new rows began. Their values may have changed."
                    : "Writing the supplied values did not begin. Excel may have filled formulas in the new rows.";
                throw new InvalidOperationException(
                    $"Append to table '{tableName}' failed after table expansion began. " +
                    $"{state} Inspect the table before retrying. {ex.Message}", ex);
            }
            finally
            {
                ComUtilities.Release(ref destination);
                ComUtilities.Release(ref firstNewCell);
                ComUtilities.Release(ref bodyCells);
                ComUtilities.Release(ref body);
                ComUtilities.Release(ref expandedRange);
                ComUtilities.Release(ref listRows);
                ComUtilities.Release(ref tableRange);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref table);
            }
        });
    }

    private static void EnsureAppendSpace(ExcelContext context, Excel.Worksheet sheet, int lastTableRow,
        int firstColumn, int rowsToAdd, int columnCount)
    {
        Excel.Range? sheetRows = null;
        Excel.Range? cells = null;
        Excel.Range? anchor = null;
        Excel.Range? expansion = null;
        Excel.WorksheetFunction? functions = null;
        try
        {
            sheetRows = sheet.Rows;
            if ((long)lastTableRow + rowsToAdd > sheetRows.Count)
                throw new ArgumentException("The requested table append exceeds the worksheet row boundary.");

            cells = sheet.Cells;
            anchor = cells[lastTableRow + 1, firstColumn];
            expansion = anchor.Resize[rowsToAdd, columnCount];
            functions = context.App.WorksheetFunction;
            if (RangeMergeDiscovery.GetMergeCellsState((object)expansion.MergeCells) != false ||
                functions.CountA(expansion) != 0)
            {
                throw new OperationFailureException(
                    OperationFailureCategory.Conflict,
                    $"Table expansion would reach occupied or merged cells on worksheet '{sheet.Name}' at {expansion.Address}.");
            }
        }
        finally
        {
            ComUtilities.Release(ref functions);
            ComUtilities.Release(ref expansion);
            ComUtilities.Release(ref anchor);
            ComUtilities.Release(ref cells);
            ComUtilities.Release(ref sheetRows);
        }
    }

    /// <inheritdoc />
    public TableDataResult GetData(IExcelBatch batch, string tableName, bool visibleOnly)
    {
        ValidateRequiredTableName(tableName);

        var result = new TableDataResult
        {
            FilePath = batch.WorkbookPath,
            TableName = tableName
        };

        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            Excel.ListColumns? listColumns = null;
            Excel.ListRows? listRows = null;
            Excel.Range? dataBodyRange = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                listColumns = table.ListColumns;
                int columnCount = listColumns.Count;
                result.ColumnCount = columnCount;

                for (int i = 1; i <= columnCount; i++)
                {
                    Excel.ListColumn? column = null;
                    try
                    {
                        column = listColumns[i];
                        string? columnName = column.Name;
                        if (!string.IsNullOrEmpty(columnName))
                        {
                            result.Headers.Add(columnName);
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref column);
                    }
                }

                dataBodyRange = table.DataBodyRange;
                if (dataBodyRange == null)
                {
                    result.Success = true;
                    result.RowCount = 0;
                    return result;
                }

                object? rawValues = dataBodyRange.Value2;
                if (rawValues is null && RangeHelpers.GetDimensions(dataBodyRange) != (1, 1))
                    throw new InvalidOperationException("Excel did not return values for the table's data range.");

                if (visibleOnly)
                    listRows = table.ListRows;
                int listRowCount = listRows?.Count ?? 0;

                if (rawValues is object[,] array2D)
                {
                    int rows = array2D.GetLength(0);
                    int cols = array2D.GetLength(1);

                    for (int r = 1; r <= rows; r++)
                    {
                        ct.ThrowIfCancellationRequested();
                        if (visibleOnly && !IsListRowVisible(listRows, listRowCount, r))
                            continue;

                        var row = new List<object?>(cols);
                        for (int c = 1; c <= cols; c++)
                        {
                            row.Add(array2D[r, c]);
                        }
                        result.Data.Add(row);
                    }
                }
                else
                {
                    if (!visibleOnly || IsListRowVisible(listRows, listRowCount, 1))
                    {
                        result.Data.Add([rawValues]);
                    }
                }

                result.RowCount = result.Data.Count;
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref listRows);
                ComUtilities.Release(ref dataBodyRange);
                ComUtilities.Release(ref listColumns);
                ComUtilities.Release(ref table);
            }
        });
    }

    private static bool IsListRowVisible(Excel.ListRows? listRows, int listRowCount, int index)
    {
        if (listRows is null || index > listRowCount)
        {
            return true;
        }

        Excel.ListRow? listRow = null;
        Excel.Range? rowRange = null;
        Excel.Range? entireRow = null;
        try
        {
            listRow = listRows[index];
            rowRange = listRow.Range;
            entireRow = rowRange.EntireRow;

            object? hiddenValue = entireRow.Hidden;
            bool hidden = hiddenValue switch
            {
                bool b => b,
                null => false,
                string s when bool.TryParse(s, out var parsed) => parsed,
                IConvertible convertible => convertible.ToBoolean(CultureInfo.InvariantCulture),
                _ => false
            };

            return !hidden;
        }
        finally
        {
            ComUtilities.Release(ref entireRow);
            ComUtilities.Release(ref rowRange);
            ComUtilities.Release(ref listRow);
        }
    }
}
