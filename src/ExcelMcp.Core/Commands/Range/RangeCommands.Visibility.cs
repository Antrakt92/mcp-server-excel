using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    private const int MaxVisibilityPageSize = 10_000;

    /// <inheritdoc />
    public RangeVisibilityResult GetVisibility(
        IExcelBatch batch, string sheetName, string rangeAddress, VisibilityAxis axis,
        int? offset = null, int? limit = null)
    {
        ValidateVisibilityInputs(rangeAddress, axis);
        int start = offset.GetValueOrDefault();
        ArgumentOutOfRangeException.ThrowIfNegative(start, nameof(offset));
        if (limit is < 1 or > MaxVisibilityPageSize)
            throw new ArgumentOutOfRangeException(nameof(limit), $"Limit must be between 1 and {MaxVisibilityPageSize}.");
        if (start != 0 && limit is null)
            throw new ArgumentException("A positive offset requires a page limit.", nameof(offset));
        return batch.Execute((context, ct) =>
        {
            Excel.Range? scope = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? dimensions = null;
            Excel.AutoFilter? filter = null;
            Excel.Range? filterRange = null;
            Excel.Range? filterRows = null;
            try
            {
                ct.ThrowIfCancellationRequested();
                scope = RangeHelpers.ResolveRange(context.Book, sheetName, rangeAddress, out string? error);
                if (scope is null)
                    throw new InvalidOperationException(error ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                var intervals = GetVisibilityIntervals(scope, axis, ct);
                int total = intervals.Sum(interval => interval.Last - interval.First + 1);
                if (limit is null && total > MaxVisibilityPageSize)
                    throw new ArgumentException(
                        $"The scope contains {total} dimensions. Use a limit of at most {MaxVisibilityPageSize} and continue with nextOffset.",
                        nameof(limit));
                int pageSize = Math.Min(limit ?? MaxVisibilityPageSize, Math.Max(0, total - start));
                sheet = scope.Worksheet;
                var result = new RangeVisibilityResult
                {
                    FilePath = batch.WorkbookPath,
                    Action = "get-visibility",
                    SheetName = sheet.Name,
                    RangeAddress = scope.Address,
                    Axis = axis,
                    SheetFilterMode = sheet.FilterMode,
                    TotalDimensions = total,
                    NextOffset = start < total && pageSize < total - start ? start + pageSize : null
                };
                int firstFilterDataRow = 0;
                int lastFilterDataRow = -1;
                if (axis == VisibilityAxis.Rows && sheet.AutoFilterMode)
                {
                    filter = sheet.AutoFilter;
                    if (filter is not null)
                    {
                        filterRange = filter.Range;
                        filterRows = filterRange.Rows;
                        firstFilterDataRow = filterRange.Row + 1;
                        lastFilterDataRow = filterRange.Row + filterRows.Count - 1;
                    }
                }
                dimensions = axis == VisibilityAxis.Rows ? sheet.Rows : sheet.Columns;
                foreach (int index in EnumerateVisibilityPage(intervals, start, pageSize))
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Range? dimension = null;
                    try
                    {
                        dimension = dimensions[index];
                        object hidden = dimension.Hidden;
                        if (hidden is not bool isHidden)
                            throw new InvalidOperationException($"Excel returned an indeterminate hidden state for {axis} {index}.");
                        object? nativeSize = axis == VisibilityAxis.Rows ? dimension.RowHeight : dimension.ColumnWidth;
                        result.Items.Add(new DimensionVisibility(index, isHidden,
                            isHidden ? "undetermined" : "not-hidden",
                            nativeSize is null ? null : Convert.ToDouble(nativeSize, CultureInfo.InvariantCulture),
                            axis == VisibilityAxis.Rows ? "points" : "character-width",
                            Convert.ToInt32((object)dimension.OutlineLevel, CultureInfo.InvariantCulture),
                            axis == VisibilityAxis.Rows && index >= firstFilterDataRow && index <= lastFilterDataRow));
                    }
                    finally
                    {
                        ComUtilities.Release(ref dimension);
                    }
                }
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref dimensions);
                ComUtilities.Release(ref filterRows);
                ComUtilities.Release(ref filterRange);
                ComUtilities.Release(ref filter);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref scope);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetVisibility(IExcelBatch batch, string sheetName, string rangeAddress,
        VisibilityAxis axis, bool hidden)
    {
        ValidateVisibilityInputs(rangeAddress, axis);
        return batch.Execute((context, ct) =>
        {
            Excel.Range? scope = null;
            Excel.Range? dimensions = null;
            try
            {
                ct.ThrowIfCancellationRequested();
                scope = RangeHelpers.ResolveRange(context.Book, sheetName, rangeAddress, out string? error);
                if (scope is null)
                    throw new InvalidOperationException(error ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                dimensions = axis == VisibilityAxis.Rows ? scope.EntireRow : scope.EntireColumn;
                ct.ThrowIfCancellationRequested();
                dimensions.Hidden = hidden;
                return new OperationResult
                {
                    FilePath = batch.WorkbookPath,
                    Action = "set-visibility",
                    Success = true,
                    Message = "Whole dimensions updated; filter criteria and outline groups are unchanged."
                };
            }
            finally
            {
                ComUtilities.Release(ref dimensions);
                ComUtilities.Release(ref scope);
            }
        });
    }

    private static void ValidateVisibilityInputs(string rangeAddress, VisibilityAxis axis)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(rangeAddress);
        if (!Enum.IsDefined(axis))
            throw new ArgumentOutOfRangeException(nameof(axis));
    }

    private static List<(int First, int Last)> GetVisibilityIntervals(
        Excel.Range scope, VisibilityAxis axis, CancellationToken ct)
    {
        List<(int First, int Last)> intervals = [];
        Excel.Areas? areas = null;
        try
        {
            areas = scope.Areas;
            int areaCount = areas.Count;
            for (int index = 1; index <= areaCount; index++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.Range? area = null;
                Excel.Range? dimensions = null;
                try
                {
                    area = areas[index];
                    dimensions = axis == VisibilityAxis.Rows ? area.Rows : area.Columns;
                    int first = axis == VisibilityAxis.Rows ? area.Row : area.Column;
                    intervals.Add((first, first + dimensions.Count - 1));
                }
                finally
                {
                    ComUtilities.Release(ref dimensions);
                    ComUtilities.Release(ref area);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref areas);
        }

        intervals.Sort(static (left, right) => left.First.CompareTo(right.First));
        List<(int First, int Last)> merged = [];
        foreach (var interval in intervals)
        {
            ct.ThrowIfCancellationRequested();
            if (merged.Count > 0 && interval.First <= merged[^1].Last + 1)
            {
                var previous = merged[^1];
                merged[^1] = (previous.First, Math.Max(previous.Last, interval.Last));
            }
            else
            {
                merged.Add(interval);
            }
        }
        return merged;
    }

    private static IEnumerable<int> EnumerateVisibilityPage(
        List<(int First, int Last)> intervals, int offset, int count)
    {
        foreach (var (first, last) in intervals)
        {
            if (count == 0)
                yield break;
            int length = last - first + 1;
            if (offset >= length)
            {
                offset -= length;
                continue;
            }
            int pageLength = Math.Min(count, length - offset);
            int pageStart = first + offset;
            for (int index = pageStart; index < pageStart + pageLength; index++)
                yield return index;
            count -= pageLength;
            offset = 0;
        }
    }
}
