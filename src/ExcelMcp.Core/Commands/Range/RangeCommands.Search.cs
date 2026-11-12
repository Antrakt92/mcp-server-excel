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

                HashSet<string>? commentAddresses = null;
                if (findOptions.SearchComments)
                {
                    commentAddresses = new(StringComparer.Ordinal);
                    AddFindMatches(range, result, findOptions, maxMatches,
                        Excel.XlFindLookIn.xlComments, ct, commentAddresses, readingComments: true);
                    AddFindMatches(range, result, findOptions, maxMatches,
                        Excel.XlFindLookIn.xlCommentsThreaded, ct, commentAddresses, readingComments: true);
                }

                if (findOptions.SearchFormulas || findOptions.SearchValues || !findOptions.SearchComments)
                {
                    var lookIn = findOptions.SearchFormulas && !findOptions.SearchValues
                        ? Excel.XlFindLookIn.xlFormulas
                        : Excel.XlFindLookIn.xlValues;
                    AddFindMatches(range, result, findOptions, maxMatches, lookIn, ct, commentAddresses,
                        readingComments: false);
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
        Excel.XlFindLookIn lookIn, CancellationToken ct, HashSet<string>? commentAddresses, bool readingComments)
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

                bool newMatch = commentAddresses is null ||
                    (readingComments ? commentAddresses.Add(address) : !commentAddresses.Contains(address));
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
                    ReplaceFirst(ctx, range, findValue, replaceValue, replaceOptions, ct);
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
        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? range = null;
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

                if (sortColumns.Count == 0)
                {
                    throw new ArgumentException("At least one sort column must be specified", nameof(sortColumns));
                }

                // Excel COM supports up to 3 sort keys
                if (sortColumns.Count > 3)
                {
                    throw new ArgumentException("Excel COM API supports maximum 3 sort columns", nameof(sortColumns));
                }

                // Get sort key ranges
                columns = range.Columns;
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
                ComUtilities.Release(ref range);
            }
        });
    }

    // === NATIVE EXCEL COM OPERATIONS ===

}
