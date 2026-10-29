namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Pure geometry and admission checks performed before materializing Excel grids.</summary>
internal sealed record RangeReadWindow(int Rows, int Columns, int? NextRowOffset, int? NextColumnOffset)
{
    internal const int MaximumPageCells = 10_000;
    internal const int MaximumReadCells = 1_000_000;
    internal const int MaximumTileCells = 1000;

    internal static IEnumerable<RangeReadTile> Tiles(int rows, int columns)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(rows, 1);
        ArgumentOutOfRangeException.ThrowIfLessThan(columns, 1);
        int tileColumns = Math.Min(columns, MaximumTileCells);
        int tileRows = MaximumTileCells / tileColumns;
        for (int row = 0; row < rows; row += tileRows)
            for (int column = 0; column < columns; column += tileColumns)
                yield return new(row, column, Math.Min(tileRows, rows - row), Math.Min(tileColumns, columns - column));
    }

    internal static void Validate(int rowOffset, int columnOffset, int pageRows, int pageColumns)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(rowOffset);
        ArgumentOutOfRangeException.ThrowIfNegative(columnOffset);
        ArgumentOutOfRangeException.ThrowIfLessThan(pageRows, 1);
        ArgumentOutOfRangeException.ThrowIfLessThan(pageColumns, 1);
        EnsureCellLimit(pageRows, pageColumns, MaximumPageCells);
    }

    internal static void EnsureCellLimit(long rows, long columns, int limit, string? guidance = null)
    {
        if (rows * columns > limit)
            throw new ArgumentOutOfRangeException(nameof(rows),
                $"Requested scope exceeds the {limit:N0}-cell read limit. " +
                (guidance ?? "Use get-used-range-info for bounds and get-values-page or smaller explicit ranges; no values were read."));
    }

    internal static RangeReadWindow Create(int totalRows, int totalColumns, int rowOffset, int columnOffset,
        int pageRows, int pageColumns)
    {
        Validate(rowOffset, columnOffset, pageRows, pageColumns);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(rowOffset, totalRows);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(columnOffset, totalColumns);
        int rows = Math.Min(pageRows, totalRows - rowOffset);
        int columns = Math.Min(pageColumns, totalColumns - columnOffset);
        if (columnOffset + columns < totalColumns)
            return new(rows, columns, rowOffset, columnOffset + columns);
        if (rowOffset + rows < totalRows)
            return new(rows, columns, rowOffset + rows, 0);
        return new(rows, columns, null, null);
    }
}

internal sealed record RangeReadTile(int RowOffset, int ColumnOffset, int Rows, int Columns);
