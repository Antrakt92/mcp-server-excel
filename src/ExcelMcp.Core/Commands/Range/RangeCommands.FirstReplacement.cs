using System.Text;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    private static void ReplaceFirst(ExcelContext context, Excel.Range range, string findValue,
        string replaceValue, ReplaceOptions options, CancellationToken ct)
    {
        ArgumentException.ThrowIfNullOrEmpty(findValue);
        Excel.Range? cell = null;
        try
        {
            cell = FindFirstReplacementCell(range, findValue, options, ct);
            if (cell is null)
                return;

            bool hasFormula = cell.HasFormula is true;
            string updated = replaceValue;
            if (!options.MatchEntireCell)
            {
                string literal = DecodeFirstReplacementLiteral(findValue);
                object? original = hasFormula ? ReadFormulas(context, cell) : cell.Value2;
                if (original is not string text)
                {
                    throw new ArgumentException(
                        "Partial first replacement requires stored text or a formula. Use matchEntireCell for a number or other non-text value.");
                }

                var comparison = options.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                int index = text.IndexOf(literal, comparison);
                if (index < 0)
                {
                    throw new InvalidOperationException(
                        "Excel found a match that could not be located unambiguously in the stored text or formula.");
                }

                updated = string.Concat(text.AsSpan(0, index), replaceValue, text.AsSpan(index + literal.Length));
            }

            ct.ThrowIfCancellationRequested();
            if (updated.Length == 0)
                cell.ClearContents();
            else if (!hasFormula)
                // Range.Replace on a single cell can modify other worksheet cells.
                cell.Value2 = options.MatchEntireCell ? updated : "'" + updated;
            else if (context.Capabilities.SupportsFormula2)
                cell.Formula2 = updated;
            else
                cell.Formula = updated;
        }
        finally
        {
            ComUtilities.Release(ref cell);
        }
    }

    private static Excel.Range? FindFirstReplacementCell(Excel.Range range, string findValue,
        ReplaceOptions options, CancellationToken ct)
    {
        Excel.Areas? areas = null;
        Excel.Range? first = null;
        try
        {
            areas = range.Areas;
            int count = areas.Count;
            int firstRow = int.MaxValue;
            int firstColumn = int.MaxValue;
            for (int index = 1; index <= count; index++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.Range? area = null;
                Excel.Range? cells = null;
                Excel.Range? last = null;
                Excel.Range? candidate = null;
                try
                {
                    area = areas[index];
                    var (rows, columns) = RangeHelpers.GetDimensions(area);
                    cells = area.Cells;
                    last = cells[rows, columns];
                    // Find starts after the supplied cell and wraps, so start after the bottom-right.
                    candidate = area.Find(
                        What: findValue,
                        After: last,
                        LookIn: Excel.XlFindLookIn.xlFormulas,
                        LookAt: options.MatchEntireCell ? Excel.XlLookAt.xlWhole : Excel.XlLookAt.xlPart,
                        SearchOrder: Excel.XlSearchOrder.xlByRows,
                        SearchDirection: Excel.XlSearchDirection.xlNext,
                        MatchCase: options.MatchCase,
                        MatchByte: false,
                        SearchFormat: false);
                    if (candidate is null)
                        continue;

                    int row = candidate.Row;
                    int column = candidate.Column;
                    int areaRow = area.Row;
                    int areaColumn = area.Column;
                    if (row < areaRow || row >= areaRow + rows || column < areaColumn || column >= areaColumn + columns)
                        throw new InvalidOperationException("Excel returned a replacement match outside the requested range.");

                    if (row < firstRow || (row == firstRow && column < firstColumn))
                    {
                        ComUtilities.Release(ref first);
                        first = candidate;
                        candidate = null;
                        firstRow = row;
                        firstColumn = column;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref candidate);
                    ComUtilities.Release(ref last);
                    ComUtilities.Release(ref cells);
                    ComUtilities.Release(ref area);
                }
            }

            var result = first;
            first = null;
            return result;
        }
        finally
        {
            ComUtilities.Release(ref first);
            ComUtilities.Release(ref areas);
        }
    }

    private static string DecodeFirstReplacementLiteral(string pattern)
    {
        var literal = new StringBuilder(pattern.Length);
        for (int index = 0; index < pattern.Length; index++)
        {
            char character = pattern[index];
            if (character is '*' or '?')
            {
                throw new ArgumentException(
                    "Partial first replacement requires literal text. Escape wildcard characters with '~' or use matchEntireCell.");
            }
            if (character == '~' && index + 1 < pattern.Length && pattern[index + 1] is '*' or '?' or '~')
                character = pattern[++index];
            literal.Append(character);
        }
        return literal.ToString();
    }
}
