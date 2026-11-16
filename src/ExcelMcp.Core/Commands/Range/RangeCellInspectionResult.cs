using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Native value, formula, and displayed-text inspection of a small rectangle.</summary>
public sealed class RangeCellInspectionResult : OperationResult
{
    /// <summary>Resolved worksheet, including when the request used a named range.</summary>
    public string SheetName { get; init; } = string.Empty;

    /// <summary>Resolved A1 rectangle.</summary>
    public string RangeAddress { get; init; } = string.Empty;

    /// <summary>Cells in row-major order.</summary>
    public List<InspectedRangeCell> Cells { get; init; } = [];
}

/// <summary>Native snapshots and explicit per-field truncation.</summary>
public sealed record InspectedRangeCell(
    string Address,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.Never)] object? Value,
    string? Formula,
    string DisplayedText,
    int? ExcelErrorCode,
    bool ValueTruncated,
    bool FormulaTruncated,
    bool DisplayedTextTruncated);
