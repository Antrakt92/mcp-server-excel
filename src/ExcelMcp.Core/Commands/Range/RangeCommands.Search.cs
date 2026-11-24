using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;


namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range search and sort operations (find, replace, sort)
/// </summary>
public partial class RangeCommands
{
    // === FIND/REPLACE OPERATIONS ===

    /// <inheritdoc />
    public RangeFindResult Find(IExcelBatch batch, string sheetName, string rangeAddress, string searchValue, FindOptions? findOptions = null, int maxMatches = 10)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxMatches);
        findOptions ??= new FindOptions();

        var result = new RangeFindResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = rangeAddress,
            SearchValue = searchValue
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

                List<Excel.XlFindLookIn> lookIns = [];
                if (findOptions.SearchComments)
                {
                    lookIns.Add(Excel.XlFindLookIn.xlComments);
                    lookIns.Add(Excel.XlFindLookIn.xlCommentsThreaded);
                }

                // Keep the values fallback when no search source was selected.
                if (findOptions.SearchValues || (!findOptions.SearchFormulas && !findOptions.SearchComments))
                    lookIns.Add(Excel.XlFindLookIn.xlValues);
                if (findOptions.SearchFormulas)
                    lookIns.Add(Excel.XlFindLookIn.xlFormulas);

                HashSet<string>? matchedAddresses = lookIns.Count > 1 ? new(StringComparer.Ordinal) : null;
                for (int index = 0; index < lookIns.Count; index++)
                {
                    // The final pass only checks earlier addresses; its new matches need no storage.
                    AddFindMatches(range, result, findOptions, maxMatches, lookIns[index], ct, matchedAddresses,
                        recordMatches: index < lookIns.Count - 1);
                }

                ct.ThrowIfCancellationRequested();
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }

    private static void AddFindMatches(Excel.Range range, RangeFindResult result, FindOptions options, int maxMatches,
        Excel.XlFindLookIn lookIn, CancellationToken ct, HashSet<string>? matchedAddresses, bool recordMatches)
    {
        Excel.Range? foundCell = null;
        Excel.Range? nextCell = null;
        try
        {
            ct.ThrowIfCancellationRequested();
            foundCell = range.Find(
                What: result.SearchValue,
                LookIn: lookIn,
                LookAt: options.MatchEntireCell ? Excel.XlLookAt.xlWhole : Excel.XlLookAt.xlPart,
                SearchOrder: Excel.XlSearchOrder.xlByRows,
                SearchDirection: Excel.XlSearchDirection.xlNext,
                MatchCase: options.MatchCase,
                MatchByte: false,
                SearchFormat: false);

            string? firstAddress = null;
            while (foundCell is not null)
            {
                ct.ThrowIfCancellationRequested();
                string address = foundCell.Address;
                if (address == firstAddress)
                    break;
                firstAddress ??= address;

                bool newMatch = matchedAddresses is null ||
                    (recordMatches ? matchedAddresses.Add(address) : !matchedAddresses.Contains(address));
                if (newMatch)
                {
                    result.TotalCount++;
                    if (result.MatchingCells.Count < maxMatches)
                    {
                        result.MatchingCells.Add(new RangeCell
                        {
                            Address = address,
                            Row = foundCell.Row,
                            Column = foundCell.Column,
                            Value = foundCell.Value2
                        });
                    }
                }

                nextCell = range.FindNext(foundCell);
                ComUtilities.Release(ref foundCell);
                foundCell = nextCell;
                nextCell = null;
            }
        }
        finally
        {
            ComUtilities.Release(ref nextCell);
            ComUtilities.Release(ref foundCell);
        }
    }

    /// <inheritdoc />
    public OperationResult Replace(IExcelBatch batch, string sheetName, string rangeAddress, string findValue, string replaceValue, ReplaceOptions replaceOptions)
    {
        string? firstLiteral = null;
        if (!replaceOptions.ReplaceAll)
        {
            ArgumentException.ThrowIfNullOrEmpty(findValue);
            if (!replaceOptions.MatchEntireCell)
                firstLiteral = DecodeFirstReplacementLiteral(findValue);
        }

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

                if (replaceOptions.ReplaceAll)
                {
                    range.Replace(
                        What: findValue,
                        Replacement: replaceValue,
                        LookAt: replaceOptions.MatchEntireCell ? Excel.XlLookAt.xlWhole : Excel.XlLookAt.xlPart,
                        SearchOrder: Excel.XlSearchOrder.xlByRows,
                        MatchCase: replaceOptions.MatchCase,
                        MatchByte: false);
                }
                else
                {
                    ReplaceFirst(ctx, range, findValue, replaceValue, firstLiteral, replaceOptions, ct);
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }

    // === SORT OPERATIONS ===

    /// <inheritdoc />
    public OperationResult Sort(IExcelBatch batch, string sheetName, string rangeAddress, List<SortColumn> sortColumns, bool hasHeaders = true)
    {
        if (sortColumns.Count is < 1 or > 3)
            throw new ArgumentException("Specify one to three sort columns.", nameof(sortColumns));

        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? range = null;
            Excel.Areas? areas = null;
            Excel.Range? columns = null;
            Excel.Range? key1 = null;
            Excel.Range? key2 = null;
            Excel.Range? key3 = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? specificError);
                if (range == null)
                {
                    throw new InvalidOperationException(specificError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));
                }

                areas = range.Areas;
                if (areas.Count != 1)
                    throw new ArgumentException("Sort requires one rectangular range.", nameof(rangeAddress));
                columns = range.Columns;
                int width = columns.Count;
                // Excel's column indexer can return a column outside its parent range.
                if (sortColumns.Any(column => column is null || column.ColumnIndex < 1 || column.ColumnIndex > width))
                    throw new ArgumentException($"Sort column indices must be between 1 and {width} for this range.", nameof(sortColumns));

                key1 = (Excel.Range)columns[sortColumns[0].ColumnIndex];
                if (sortColumns.Count >= 2) key2 = (Excel.Range)columns[sortColumns[1].ColumnIndex];
                if (sortColumns.Count >= 3) key3 = (Excel.Range)columns[sortColumns[2].ColumnIndex];

                // Excel COM constants
                int order1 = sortColumns[0].Ascending ? 1 : 2; // xlAscending : xlDescending
                int order2 = sortColumns.Count >= 2 ? (sortColumns[1].Ascending ? 1 : 2) : 1;
                int order3 = sortColumns.Count >= 3 ? (sortColumns[2].Ascending ? 1 : 2) : 1;
                int header = hasHeaders ? 1 : 2; // xlYes : xlNo

                // Call Range.Sort method
                range.Sort(
                    Key1: key1,
                    Order1: (Excel.XlSortOrder)order1,
                    Key2: (object?)key2 ?? Type.Missing,
                    Order2: (Excel.XlSortOrder)order2,
                    Key3: (object?)key3 ?? Type.Missing,
                    Order3: (Excel.XlSortOrder)order3,
                    Header: (Excel.XlYesNoGuess)header,
                    OrderCustom: 1,
                    MatchCase: false,
                    Orientation: (Excel.XlSortOrientation)1,
                    SortMethod: Excel.XlSortMethod.xlPinYin
                );

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref key3);
                ComUtilities.Release(ref key2);
                ComUtilities.Release(ref key1);
                ComUtilities.Release(ref columns);
                ComUtilities.Release(ref areas);
                ComUtilities.Release(ref range);
            }
        });
    }

    // === NATIVE EXCEL COM OPERATIONS ===

}
