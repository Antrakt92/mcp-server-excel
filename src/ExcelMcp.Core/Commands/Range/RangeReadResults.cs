using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Native UsedRange dimensions without reading workbook values.</summary>
public sealed class RangeBoundsResult : ResultBase
{
    /// <summary>Resolved worksheet.</summary>
    public string SheetName { get; set; } = string.Empty;
    /// <summary>Absolute native UsedRange address, including formatting-only cells.</summary>
    public string RangeAddress { get; set; } = string.Empty;
    /// <summary>Number of native rows.</summary>
    public int RowCount { get; set; }
    /// <summary>Number of native columns.</summary>
    public int ColumnCount { get; set; }
    /// <summary>Number of native cells, not the count of populated values.</summary>
    public long CellCount { get; set; }
}

/// <summary>One bounded rectangular page; separate calls are not a workbook snapshot.</summary>
public sealed class RangeReadPageResult : RangeValueResult
{
    /// <summary>Complete resolved scope used to compute this page.</summary>
    public string ScopeAddress { get; set; } = string.Empty;
    /// <summary>Complete scope row count.</summary>
    public int TotalRowCount { get; set; }
    /// <summary>Complete scope column count.</summary>
    public int TotalColumnCount { get; set; }
    /// <summary>Zero-based requested row offset.</summary>
    public int RowOffset { get; set; }
    /// <summary>Zero-based requested column offset.</summary>
    public int ColumnOffset { get; set; }
    /// <summary>Whether another page follows in column-first order.</summary>
    public bool HasMore { get; set; }
    /// <summary>Next row offset; null when complete.</summary>
    public int? NextRowOffset { get; set; }
    /// <summary>Next column offset; null when complete.</summary>
    public int? NextColumnOffset { get; set; }
}

/// <summary>Native aggregate formatting; mixed properties are not a per-cell style histogram.</summary>
public sealed class RangeFormatSummaryResult : ResultBase
{
    /// <summary>Resolved worksheet.</summary>
    public string SheetName { get; set; } = string.Empty;
    /// <summary>Resolved rectangular scope.</summary>
    public string RangeAddress { get; set; } = string.Empty;
    /// <summary>Native cell count.</summary>
    public long CellCount { get; set; }
    /// <summary>Requested stored/displayed view.</summary>
    public FormatView View { get; set; }
    /// <summary>Aggregate stored properties; null when not requested.</summary>
    public CellFormatSnapshot? Stored { get; set; }
    /// <summary>Aggregate displayed properties; null when not requested.</summary>
    public CellFormatSnapshot? Displayed { get; set; }
}
