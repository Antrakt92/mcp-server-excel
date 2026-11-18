using System.Text;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// TableCommands partial class - Structured Reference operations
/// </summary>
public partial class TableCommands
{
    /// <summary>
    /// Gets structured reference information for a table region or column
    /// </summary>
    public TableStructuredReferenceResult GetStructuredReference(
        IExcelBatch batch,
        string tableName,
        TableRegion region,
        string? columnName = null)
    {
        ValidateRequiredTableName(tableName);

        var result = new TableStructuredReferenceResult { FilePath = batch.WorkbookPath };
        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                result.TableName = tableName;
                result.Region = region;
                result.ColumnName = columnName;

                // Get sheet name
                Excel.Range? range = null;
                Excel.Worksheet? sheet = null;
                try
                {
                    range = table.Range;
                    sheet = range.Worksheet;
                    result.SheetName = sheet.Name;
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref range);
                }

                // Build structured reference
                result.StructuredReference = BuildStructuredReference(tableName, region, columnName);

                // Get region range
                Excel.Range? regionRange = null;
                Excel.Range? rows = null;
                Excel.Range? columns = null;
                try
                {
                    regionRange = GetRegionRange(table, region, columnName);
                    result.RangeAddress = regionRange.Address;
                    rows = regionRange.Rows;
                    columns = regionRange.Columns;
                    result.RowCount = rows.Count;
                    result.ColumnCount = columns.Count;
                }
                finally
                {
                    ComUtilities.Release(ref columns);
                    ComUtilities.Release(ref rows);
                    ComUtilities.Release(ref regionRange);
                }

                result.Success = true;

                return result;
            }
            finally
            {
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <summary>
    /// Builds a structured reference formula string for a table region
    /// </summary>
    private static string BuildStructuredReference(string tableName, TableRegion region, string? columnName)
    {
        if (!string.IsNullOrEmpty(columnName))
        {
            string column = EscapeStructuredReferenceColumn(columnName);
            return region switch
            {
                TableRegion.All => $"{tableName}[[#All],[{column}]]",
                TableRegion.Data => $"{tableName}[[{column}]]", // Default is data
                TableRegion.Headers => $"{tableName}[[#Headers],[{column}]]",
                TableRegion.Totals => $"{tableName}[[#Totals],[{column}]]",
                TableRegion.ThisRow => $"{tableName}[@[{column}]]",
                _ => $"{tableName}[[{column}]]"
            };
        }
        else
        {
            // Entire region reference
            return region switch
            {
                TableRegion.All => $"{tableName}[#All]",
                TableRegion.Data => $"{tableName}[#Data]",
                TableRegion.Headers => $"{tableName}[#Headers]",
                TableRegion.Totals => $"{tableName}[#Totals]",
                TableRegion.ThisRow => $"{tableName}[@]",
                _ => $"{tableName}[#All]"
            };
        }
    }

    private static string EscapeStructuredReferenceColumn(string columnName)
    {
        const string escapeCharacters = "[]#'@";
        int firstSpecial = columnName.AsSpan().IndexOfAny(escapeCharacters);
        if (firstSpecial < 0)
            return columnName;

        var escaped = new StringBuilder(columnName.Length);
        escaped.Append(columnName.AsSpan(0, firstSpecial));
        foreach (char character in columnName.AsSpan(firstSpecial))
        {
            if (escapeCharacters.Contains(character))
                escaped.Append('\'');
            escaped.Append(character);
        }
        return escaped.ToString();
    }

    /// <summary>
    /// Gets the Excel Range object for a specific table region
    /// </summary>
    private static Excel.Range GetRegionRange(Excel.ListObject table, TableRegion region, string? columnName)
    {
        Excel.Range? regionRange = region switch
        {
            TableRegion.All => table.Range,
            TableRegion.Data => table.DataBodyRange,
            TableRegion.Headers => table.HeaderRowRange,
            TableRegion.Totals => table.TotalsRowRange,
            TableRegion.ThisRow => table.Range, // Full range for ThisRow context
            _ => throw new ArgumentOutOfRangeException(nameof(region), region, "Unknown table region.")
        };
        if (regionRange is null)
        {
            string reason = region switch
            {
                TableRegion.Data => "has no data rows",
                TableRegion.Headers => "has no header row",
                TableRegion.Totals => "does not show a totals row",
                _ => $"has no {region} range"
            };
            throw new InvalidOperationException($"Table '{table.Name}' {reason}.");
        }

        if (string.IsNullOrEmpty(columnName))
            return regionRange;

        Excel.ListColumns? columns = null;
        Excel.ListColumn? column = null;
        Excel.Range? columnRange = null;
        Excel.Application? app = null;
        try
        {
            columns = table.ListColumns;
            column = columns[columnName];
            columnRange = column.Range;
            app = table.Application;
            return app.Intersect(regionRange, columnRange)
                ?? throw new InvalidOperationException(
                    $"Table column '{columnName}' has no cells in the {region} region.");
        }
        finally
        {
            ComUtilities.Release(ref app);
            ComUtilities.Release(ref columnRange);
            ComUtilities.Release(ref column);
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref regionRange);
        }
    }
}
