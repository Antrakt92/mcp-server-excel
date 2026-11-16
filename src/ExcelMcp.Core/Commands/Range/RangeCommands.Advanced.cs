using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Merge, conditional formatting, and protection operations for Excel ranges (partial class)
/// </summary>
public partial class RangeCommands
{
    /// <inheritdoc />
    public OperationResult MergeCells(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty)
    {
        ValidateOverwritePolicy(overwritePolicy);
        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? range = null;

            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out var error)
                    ?? throw new InvalidOperationException(error ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                var dimensions = GetContentDimensions(range);
                EnsureMergeDoesNotSplitExistingAreas(range, dimensions, ct);
                EnsureDestinationWritable(ctx, range, overwritePolicy, ct, static (row, column) => row != 0 || column != 0);
                MergeWithAuthorizedDiscard(ctx.App, range, overwritePolicy);
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }

    private static void EnsureMergeDoesNotSplitExistingAreas(
        Excel.Range range, (int Rows, int Columns) dimensions, CancellationToken ct)
    {
        var merged = RangeMergeDiscovery.GetMergeCellsState((object)range.MergeCells);
        if (merged == false)
            return;

        int firstRow = range.Row;
        int firstColumn = range.Column;
        int lastRow = firstRow + dimensions.Rows - 1;
        int lastColumn = firstColumn + dimensions.Columns - 1;
        foreach (var address in RangeMergeDiscovery.CollectMergedRanges(range, merged, ct))
        {
            string[] corners = address.Split(':', 2);
            if (corners.Length == 2
                && RangeHelpers.TryParseCellReference(corners[0], out var startColumn, out var startRow)
                && RangeHelpers.TryParseCellReference(corners[1], out var endColumn, out var endRow)
                && startRow >= firstRow && startColumn >= firstColumn
                && endRow <= lastRow && endColumn <= lastColumn)
                continue;

            throw new OperationFailureException(OperationFailureCategory.Conflict,
                $"Cannot merge '{range.Address}' because it intersects existing merged area '{address}' beyond the requested range.");
        }
    }

    internal static void MergeWithAuthorizedDiscard(Excel.Application app, Excel.Range range, OverwritePolicy overwritePolicy)
    {
        if (overwritePolicy != OverwritePolicy.Allow || !app.DisplayAlerts)
        {
            range.Merge();
            return;
        }

        try
        {
            app.DisplayAlerts = false;
            range.Merge();
        }
        catch (Exception ex)
        {
            OperationCleanup.Run(() => app.DisplayAlerts = true, ex);
            throw;
        }

        try
        {
            app.DisplayAlerts = true;
        }
        catch (Exception ex)
        {
            throw new OperationCompletionUnconfirmedException(
                "Excel returned from merging the cells, but restoring its alert setting failed. Inspect the merged range before retrying.", ex);
        }
    }

    /// <inheritdoc />
    public OperationResult UnmergeCells(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;

            try
            {
                // Get sheet
                sheet = Utilities.WorksheetLookup.Get(ctx.Book, sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Unmerge cells
                range.UnMerge();

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <inheritdoc />
    public RangeMergeInfoResult GetMergeInfo(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;

            try
            {
                // Get sheet
                sheet = Utilities.WorksheetLookup.Get(ctx.Book, sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                object? mergeCells = range.MergeCells;
                bool? isMergedState = RangeMergeDiscovery.GetMergeCellsState(mergeCells);
                List<string> mergedRanges = isMergedState == false
                    ? []
                    : RangeMergeDiscovery.CollectMergedRanges(range, isMergedState, ct);

                return new RangeMergeInfoResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    SheetName = sheetName,
                    RangeAddress = rangeAddress,
                    IsMerged = isMergedState ?? mergedRanges.Count > 0,
                    MergedRanges = mergedRanges
                };
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

}
