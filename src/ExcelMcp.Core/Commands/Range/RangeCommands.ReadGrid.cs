using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    private static (int Rows, int Columns) RequireReadRectangle(Excel.Range range)
    {
        Excel.Areas? areas = null;
        try
        {
            areas = range.Areas;
            if (areas.Count != 1)
                throw new ArgumentException("Grid reads require one rectangular range; read each disjoint area separately.");
            return RangeHelpers.GetDimensions(range);
        }
        finally { ComUtilities.Release(ref areas); }
    }

    private static void ReadValueGrid(ExcelContext ctx, Excel.Range range, RangeValueResult result, CancellationToken ct,
        List<List<string>>? formulaRows = null, FormulaReferenceStyle referenceStyle = FormulaReferenceStyle.A1)
    {
        var dimensions = RequireReadRectangle(range);
        RangeReadWindow.EnsureCellLimit(dimensions.Rows, dimensions.Columns, RangeReadWindow.MaximumReadCells);
        result.RangeAddress = range.Address;
        result.RowCount = dimensions.Rows;
        result.ColumnCount = dimensions.Columns;
        var budget = new RangeReadBudget();
        budget.ReserveGrid(dimensions.Rows, dimensions.Columns, formulaRows is not null);
        budget.AddScalar(result.FilePath);
        budget.AddScalar(result.SheetName);
        budget.AddScalar(result.RangeAddress);
        if (result is RangeReadPageResult pageResult) budget.AddScalar(pageResult.ScopeAddress);
        for (int row = 0; row < dimensions.Rows; row++)
        {
            ct.ThrowIfCancellationRequested();
            result.Values.Add(new List<object?>(dimensions.Columns));
            formulaRows?.Add(new List<string>(dimensions.Columns));
        }
        int startRow = range.Row;
        int startColumn = range.Column;
        Excel.Range? cells = null;
        try
        {
            cells = range.Cells;
            foreach (var window in RangeReadWindow.Tiles(dimensions.Rows, dimensions.Columns))
            {
                ct.ThrowIfCancellationRequested();
                Excel.Range? anchor = null;
                Excel.Range? tile = null;
                try
                {
                    anchor = cells[window.RowOffset + 1, window.ColumnOffset + 1];
                    tile = anchor.Resize[window.Rows, window.Columns];
                    var grid = ExcelValueNormalizer.Normalize((object?)tile.Value2);
                    if (grid.RowCount != window.Rows || grid.ColumnCount != window.Columns)
                        throw new InvalidOperationException("Excel returned dimensions different from the requested read tile.");
                    // Reject oversized values before requesting an additional native formula array.
                    foreach (var values in grid.Values)
                        foreach (var value in values)
                        {
                            ct.ThrowIfCancellationRequested();
                            budget.AddScalar(value);
                        }
                    object? formulas = formulaRows is null ? null : ReadFormulas(ctx, tile, referenceStyle);
                    bool formulasRead = formulaRows is not null;
                    for (int row = 0; row < window.Rows; row++)
                    {
                        for (int column = 0; column < window.Columns; column++)
                        {
                            ct.ThrowIfCancellationRequested();
                            object? value = grid.Values[row][column];
                            string formula = formulaRows is null ? string.Empty : GetReturnedFormula(GridEntry(formulas, row, column));
                            if (formulaRows is not null && !IsReturnedFormula(tile, row + 1, column + 1, formula, value))
                                formula = string.Empty;
                            if (ExcelErrorMapper.TryGet(value, out int code, out var error) && IsNativeError(ctx, tile, row + 1, column + 1))
                            {
                                if (!formulasRead)
                                {
                                    formulas = ReadFormulas(ctx, tile, referenceStyle);
                                    formulasRead = true;
                                }
                                string errorFormula = GetReturnedFormula(GridEntry(formulas, row, column));
                                budget.AddError(errorFormula);
                                value = ConvertMappedErrorForRead(value, errorFormula,
                                    startRow + window.RowOffset + row, startColumn + window.ColumnOffset + column,
                                    result.CellErrors, code, error);
                                budget.AddScalar(value);
                            }
                            if (formulaRows is not null) budget.AddScalar(formula);
                            result.Values[window.RowOffset + row].Add(value);
                            formulaRows?[window.RowOffset + row].Add(formula);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref tile);
                    ComUtilities.Release(ref anchor);
                }
            }
        }
        finally { ComUtilities.Release(ref cells); }
    }

    private static object? GridEntry(object? grid, int row, int column) => grid is Array array
        ? array.GetValue(row + array.GetLowerBound(0), column + array.GetLowerBound(1)) : grid;

    private static bool IsNativeError(ExcelContext ctx, Excel.Range range, int row, int column)
    {
        Excel.Range? cells = null;
        Excel.Range? cell = null;
        Excel.WorksheetFunction? functions = null;
        try
        {
            cells = range.Cells;
            cell = cells[row, column];
            functions = ctx.App.WorksheetFunction;
            return functions.IsError(cell);
        }
        finally
        {
            ComUtilities.Release(ref functions);
            ComUtilities.Release(ref cell);
            ComUtilities.Release(ref cells);
        }
    }

    /// <inheritdoc />
    public RangeBoundsResult GetUsedRangeInfo(IExcelBatch batch, string sheetName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                ct.ThrowIfCancellationRequested();
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new OperationFailureException(OperationFailureCategory.NotFound, $"Sheet '{sheetName}' not found.");
                range = sheet.UsedRange;
                var dimensions = RangeHelpers.GetDimensions(range);
                return new RangeBoundsResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    SheetName = sheet.Name,
                    RangeAddress = range.Address,
                    RowCount = dimensions.Rows,
                    ColumnCount = dimensions.Columns,
                    CellCount = (long)dimensions.Rows * dimensions.Columns
                };
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public RangeReadPageResult GetValuesPage(IExcelBatch batch, string sheetName, string rangeAddress,
        int rowOffset = 0, int columnOffset = 0, int pageRows = 100, int pageColumns = 100)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(rangeAddress);
        RangeReadWindow.Validate(rowOffset, columnOffset, pageRows, pageColumns);
        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? scope = null;
            Excel.Range? cells = null;
            Excel.Range? offset = null;
            Excel.Range? page = null;
            try
            {
                scope = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress)
                    ?? throw new ArgumentException(RangeHelpers.GetResolveError(sheetName, rangeAddress));
                var dimensions = RequireReadRectangle(scope);
                var window = RangeReadWindow.Create(dimensions.Rows, dimensions.Columns, rowOffset, columnOffset, pageRows, pageColumns);
                cells = scope.Cells;
                offset = cells[rowOffset + 1, columnOffset + 1];
                page = offset.Resize[window.Rows, window.Columns];
                var result = new RangeReadPageResult
                {
                    FilePath = batch.WorkbookPath,
                    SheetName = sheetName,
                    ScopeAddress = scope.Address,
                    TotalRowCount = dimensions.Rows,
                    TotalColumnCount = dimensions.Columns,
                    RowOffset = rowOffset,
                    ColumnOffset = columnOffset,
                    HasMore = window.NextRowOffset.HasValue,
                    NextRowOffset = window.NextRowOffset,
                    NextColumnOffset = window.NextColumnOffset
                };
                ReadValueGrid(ctx, page, result, ct);
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref page);
                ComUtilities.Release(ref offset);
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref scope);
            }
        });
    }
}
