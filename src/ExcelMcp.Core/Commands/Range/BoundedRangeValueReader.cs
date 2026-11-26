using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Reads native values from one rectangle without materializing its entire COM array.</summary>
internal static class BoundedRangeValueReader
{
    internal static void EnsureCellLimit((int Rows, int Columns) size) =>
        RangeReadWindow.EnsureCellLimit(size.Rows, size.Columns, RangeReadWindow.MaximumReadCells,
            "Use range get-values-page with the named range or a table worksheet range. No values were read.");

    internal static List<List<object?>> Read(
        ExcelContext context, Excel.Range range, (int Rows, int Columns) size, RangeReadBudget budget,
        CancellationToken ct, Func<int, bool>? includeRow = null, Func<object?, object?>? convertValue = null)
    {
        EnsureCellLimit(size);
        int[]? outputRows = null;
        int rowCount = size.Rows;
        if (includeRow is null)
        {
            budget.ReserveGrid(rowCount, size.Columns, includeFormulas: false);
        }
        else
        {
            outputRows = new int[size.Rows];
            rowCount = 0;
            for (int row = 0; row < size.Rows; row++)
            {
                ct.ThrowIfCancellationRequested();
                if (!includeRow(row))
                {
                    outputRows[row] = -1;
                    continue;
                }

                // Stop requesting visibility after the response becomes too large even without cell values.
                budget.ReserveGrid(1, size.Columns, includeFormulas: false);
                outputRows[row] = rowCount++;
            }
        }

        var rows = new List<List<object?>>(rowCount);
        if (rowCount == 0)
            return rows;
        for (int row = 0; row < rowCount; row++)
            rows.Add(new List<object?>(size.Columns));

        Excel.Range? cells = null;
        Excel.WorksheetFunction? functions = null;
        try
        {
            cells = range.Cells;
            foreach (var window in RangeReadWindow.Tiles(size.Rows, size.Columns))
            {
                ct.ThrowIfCancellationRequested();
                if (outputRows is not null && !IncludesRow(outputRows, window))
                    continue;

                Excel.Range? anchor = null;
                Excel.Range? tile = null;
                Excel.Range? tileCells = null;
                try
                {
                    anchor = cells[window.RowOffset + 1, window.ColumnOffset + 1];
                    tile = anchor.Resize[window.Rows, window.Columns];
                    var grid = ExcelValueNormalizer.Normalize((object?)tile.Value2);
                    if (grid.RowCount != window.Rows || grid.ColumnCount != window.Columns)
                        throw new InvalidOperationException("Excel returned dimensions different from the requested read tile.");

                    for (int row = 0; row < window.Rows; row++)
                    {
                        ct.ThrowIfCancellationRequested();
                        int sourceRow = window.RowOffset + row;
                        int outputRow = outputRows?[sourceRow] ?? sourceRow;
                        if (outputRow < 0)
                            continue;

                        for (int column = 0; column < window.Columns; column++)
                        {
                            ct.ThrowIfCancellationRequested();
                            var raw = grid.Values[row][column];
                            var error = RangeCommands.ReadNativeError(context, tile, row + 1, column + 1, raw,
                                ref tileCells, ref functions);
                            var value = error?.Name ?? (convertValue is null ? raw : convertValue(raw));
                            budget.AddScalar(value);
                            rows[outputRow].Add(value);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref tileCells);
                    ComUtilities.Release(ref tile);
                    ComUtilities.Release(ref anchor);
                }
            }

            return rows;
        }
        finally
        {
            ComUtilities.Release(ref functions);
            ComUtilities.Release(ref cells);
        }
    }

    private static bool IncludesRow(int[] outputRows, RangeReadTile tile)
    {
        int end = tile.RowOffset + tile.Rows;
        for (int row = tile.RowOffset; row < end; row++)
        {
            if (outputRows[row] >= 0)
                return true;
        }
        return false;
    }
}
