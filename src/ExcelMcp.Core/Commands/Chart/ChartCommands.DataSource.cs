using Excel = Microsoft.Office.Interop.Excel;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Result of finding a chart by name.
/// </summary>
internal sealed class ChartFindResult
{
    public Excel.Chart? Chart;
    public Excel.Shape? Shape;
    public string SheetName = string.Empty;
}

/// <summary>
/// Chart data source operations - set range, add/remove series.
/// </summary>
public partial class ChartCommands
{
    /// <inheritdoc />
    public OperationResult SetSourceRange(IExcelBatch batch, string chartName, string sourceRange)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Find chart by name
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            try
            {
                // Determine strategy and delegate
                IChartStrategy strategy = _pivotStrategy.CanHandle(findResult.Chart) ? _pivotStrategy : _regularStrategy;
                strategy.SetSourceRange(findResult.Chart, sourceRange, ct);

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Void operation completed
            }
            finally
            {
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public SeriesInfo AddSeries(
        IExcelBatch batch,
        string chartName,
        string seriesName,
        string valuesRange,
        string? categoryRange = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Find chart by name
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            try
            {
                // Determine strategy and delegate
                IChartStrategy strategy = _pivotStrategy.CanHandle(findResult.Chart) ? _pivotStrategy : _regularStrategy;
                var result = strategy.AddSeries(findResult.Chart, seriesName, valuesRange, categoryRange);

                return result;
            }
            finally
            {
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult RemoveSeries(IExcelBatch batch, string chartName, int seriesIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Find chart by name
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            try
            {
                // Determine strategy and delegate
                IChartStrategy strategy = _pivotStrategy.CanHandle(findResult.Chart) ? _pivotStrategy : _regularStrategy;
                strategy.RemoveSeries(findResult.Chart, seriesIndex);

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Void operation completed
            }
            finally
            {
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <summary>Finds a unique embedded chart. The caller releases the returned chart and shape.</summary>
    private static ChartFindResult FindChart(Excel.Workbook workbook, string chartName, string? chartSheetName = null)
    {
        Excel.Shape? shape = FindChartShape(workbook, chartName, out var sheetName, chartSheetName);
        try
        {
            var found = new ChartFindResult { Shape = shape, Chart = shape?.Chart, SheetName = sheetName };
            shape = null;
            return found;
        }
        finally
        {
            ComUtilities.Release(ref shape);
        }
    }

    private static Excel.Shape? FindChartShape(Excel.Workbook workbook, string chartName, out string sheetName,
        string? chartSheetName = null)
    {
        sheetName = string.Empty;
        Excel.Sheets? worksheets = null;
        Excel.Shape? match = null;
        try
        {
            if (chartSheetName is null)
                worksheets = workbook.Worksheets;
            else
                ArgumentException.ThrowIfNullOrWhiteSpace(chartSheetName);

            int worksheetCount = worksheets?.Count ?? 1;
            for (int index = 1; index <= worksheetCount; index++)
            {
                Excel.Worksheet? worksheet = null;
                Excel.Shapes? shapes = null;
                try
                {
                    worksheet = worksheets is not null
                        ? (Excel.Worksheet)worksheets.Item[index]
                        : Utilities.WorksheetLookup.Get(workbook, chartSheetName);
                    shapes = worksheet.Shapes;
                    int shapeCount = shapes.Count;
                    for (int item = 1; item <= shapeCount; item++)
                    {
                        Excel.Shape? shape = null;
                        try
                        {
                            shape = shapes.Item(item);
                            if ((int)shape.Type != 3 ||
                                !string.Equals(shape.Name, chartName, StringComparison.OrdinalIgnoreCase))
                                continue;

                            if (match is not null)
                            {
                                throw new OperationFailureException(
                                    OperationFailureCategory.Conflict,
                                    $"Chart '{chartName}' is ambiguous. It matches a chart on worksheet " +
                                    $"'{sheetName}' and a chart on worksheet '{worksheet.Name}'. " +
                                    "Use chartSheetName on actions that accept it, or give the charts distinct names in Excel before retrying.");
                            }

                            sheetName = worksheet.Name;
                            match = shape;
                            shape = null;
                        }
                        finally
                        {
                            ComUtilities.Release(ref shape);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref shapes);
                    ComUtilities.Release(ref worksheet);
                }
            }

            if (match is null && chartSheetName is not null)
                throw new OperationFailureException(OperationFailureCategory.NotFound,
                    $"Chart '{chartName}' was not found on worksheet '{chartSheetName}'.");

            var result = match;
            match = null;
            return result;
        }
        finally
        {
            ComUtilities.Release(ref match);
            ComUtilities.Release(ref worksheets);
        }
    }
}
