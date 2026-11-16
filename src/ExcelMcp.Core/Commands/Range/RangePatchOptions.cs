using System.Text.Json;
using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>A bounded, ordered set of edits on one worksheet.</summary>
/// <param name="Edits">Single-cell edits, in application order.</param>
public sealed record RangePatchOptions(IReadOnlyList<RangeCellEdit> Edits);

/// <summary>One edit guarded by exact existing content.</summary>
/// <param name="Address">A sheet-local A1 cell address.</param>
/// <param name="Expected">Content required before writing.</param>
/// <param name="Replacement">Content to write.</param>
public sealed record RangeCellEdit(string Address, RangeCellState Expected, RangeCellState Replacement);

/// <summary>Undefined distinguishes an omitted value from a requested blank (JSON null).</summary>
/// <param name="Value">One literal scalar or blank.</param>
/// <param name="Formula">An explicit A1 formula, mutually exclusive with Value.</param>
public sealed record RangeCellState(
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)] JsonElement Value,
    string? Formula);

/// <summary>The observed result of a conditional patch.</summary>
public sealed class RangePatchResult : ResultBase
{
    /// <summary>The targeted worksheet.</summary>
    public required string SheetName { get; init; }
    /// <summary>Addresses for which Excel returned from the native write.</summary>
    public List<string> WriteReturnedAddresses { get; } = [];
    /// <summary>Addresses last observed holding the requested content.</summary>
    public List<string> ConfirmedAddresses { get; } = [];
    /// <summary>Addresses that did not match a precondition or final readback.</summary>
    public List<string> ConflictAddresses { get; } = [];
    /// <summary>Address whose setter or immediate readback could not be confirmed.</summary>
    public string? UnconfirmedAddress { get; set; }
    /// <summary>True if Excel returned different content in the immediate readback.</summary>
    public bool ReadbackMismatch { get; set; }
    /// <summary>Machine-readable classification of a content conflict.</summary>
    public string? ErrorCategory { get; set; }
}

/// <summary>An Excel failure with observations from a patch already in progress.</summary>
/// <param name="state">The observations at the point of failure.</param>
/// <param name="cause">The original Excel or readback failure.</param>
public sealed class RangePatchException(RangePatchResult state, Exception cause)
    : InvalidOperationException($"Excel did not confirm the patch at {state.UnconfirmedAddress ?? "final verification"}.", cause)
{
    /// <summary>The observations at the point of failure.</summary>
    public RangePatchResult State { get; } = state;
}
