using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Typed NumberFormat reads return invariant codes. Explicit NumberFormatLocal writes preserve
/// currency literals while translating the invariant codes accepted by both entry points.
/// </summary>
public partial class RangeCommands
{
    // === NUMBER FORMAT OPERATIONS ===

    /// <inheritdoc />
    public RangeNumberFormatResult GetNumberFormats(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        var result = new RangeNumberFormatResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = rangeAddress
        };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                ct.ThrowIfCancellationRequested();
                var (rowCount, columnCount) = RequireReadRectangle(range);
                RangeReadWindow.EnsureCellLimit(rowCount, columnCount, RangeReadWindow.MaximumPageCells,
                    "Use get-format-summary for aggregate number formats or smaller explicit rectangles; no formatting was inspected.");
                result.RangeAddress = range.Address;
                result.RowCount = rowCount;
                result.ColumnCount = columnCount;
                var budget = new RangeReadBudget();
                budget.ReserveGrid(rowCount, columnCount, false);
                budget.AddScalar(result.FilePath);
                budget.AddScalar(result.SheetName);
                budget.AddScalar(result.RangeAddress);
                object? numberFormats = range.NumberFormat;
                Excel.Range? cells = null;
                try
                {
                    if (numberFormats is null or DBNull) cells = range.Cells;
                    for (int row = 0; row < rowCount; row++)
                    {
                        ct.ThrowIfCancellationRequested();
                        var rowList = new List<string>(columnCount);
                        for (int column = 0; column < columnCount; column++)
                        {
                            ct.ThrowIfCancellationRequested();
                            string format;
                            if (cells is not null)
                            {
                                Excel.Range? cell = null;
                                try
                                {
                                    cell = cells[row + 1, column + 1];
                                    format = cell.NumberFormat as string
                                        ?? throw new InvalidOperationException("Excel returned an indeterminate number format for a single cell.");
                                }
                                finally { ComUtilities.Release(ref cell); }
                            }
                            else
                            {
                                format = (numberFormats is string uniform ? uniform : GridEntry(numberFormats, row, column)) as string
                                    ?? throw new InvalidOperationException("Excel returned an unsupported number-format value.");
                            }
                            budget.AddScalar(format);
                            rowList.Add(format);
                        }
                        result.Formats.Add(rowList);
                    }
                }
                finally { ComUtilities.Release(ref cells); }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetNumberFormat(IExcelBatch batch, string sheetName, string rangeAddress, string formatCode)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "set-number-format"
        };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                range.NumberFormatLocal = ctx.FormatTranslator.TranslateToLocale(formatCode);

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetNumberFormats(IExcelBatch batch, string sheetName, string rangeAddress, List<List<string>>? formats = null, string? formatsFile = null)
    {
        // Resolve formats from inline parameter or file
        var resolvedFormats = ParameterTransforms.ResolveFormulasOrFile(formats, formatsFile, "formats");

        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "set-number-formats"
        };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? range = null;
            Excel.Range? cells = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                (int rowCount, int columnCount) = RangeHelpers.GetDimensions(range);

                // Validate dimensions match
                if (resolvedFormats.Count != rowCount)
                {
                    throw new ArgumentException($"Format array row count ({resolvedFormats.Count}) doesn't match range row count ({rowCount})", nameof(formats));
                }

                for (int i = 0; i < resolvedFormats.Count; i++)
                {
                    if (resolvedFormats[i].Count != columnCount)
                    {
                        throw new ArgumentException($"Format array row {i + 1} column count ({resolvedFormats[i].Count}) doesn't match range column count ({columnCount})", nameof(formats));
                    }
                }

                // If single row or column, can't use 2D array - must set cell by cell
                if (rowCount == 1 || columnCount == 1)
                {
                    cells = range.Cells;
                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int col = 1; col <= columnCount; col++)
                        {
                            Excel.Range? cell = null;
                            try
                            {
                                cell = (Excel.Range)cells[row, col];
                                cell.NumberFormatLocal = ctx.FormatTranslator.TranslateToLocale(resolvedFormats[row - 1][col - 1]);
                            }
                            finally
                            {
                                ComUtilities.Release(ref cell);
                            }
                        }
                    }
                }
                else
                {
                    // For multi-row, multi-column ranges, Excel COM expects 1-based 2D array
                    object[,] formatArray = new object[rowCount, columnCount];
                    for (int row = 0; row < rowCount; row++)
                    {
                        for (int col = 0; col < columnCount; col++)
                        {
                            formatArray[row, col] = ctx.FormatTranslator.TranslateToLocale(resolvedFormats[row][col]);
                        }
                    }

                    // Set number formats via 2D array
                    range.NumberFormatLocal = formatArray;
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref range);
            }
        });
    }
}
