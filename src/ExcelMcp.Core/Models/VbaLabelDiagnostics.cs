using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>Bounded advisory inspection of VBA labels in an accessible project's source.</summary>
public sealed class VbaLabelDiagnostics
{
    /// <summary>Potential accidental labels, in project module and source order.</summary>
    public List<VbaLabelWarning> Warnings { get; set; } = [];

    /// <summary>Whether the output limit omitted additional matches.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public bool HasMore { get; set; }
}

/// <summary>A line label that matches a visible zero-argument VBA Sub.</summary>
public sealed class VbaLabelWarning
{
    /// <summary>Module containing the label.</summary>
    public required string ModuleName { get; init; }

    /// <summary>Procedure containing the label.</summary>
    public required string ProcedureName { get; init; }

    /// <summary>One-based physical line number in the native code module.</summary>
    public required int Line { get; init; }

    /// <summary>Identifier used as the label.</summary>
    public required string Label { get; init; }

    /// <summary>Why the matching label may be unintended.</summary>
    public string Message { get; init; } = "This label matches a zero-argument Sub and has no jump reference in this procedure. If you intended to invoke the Sub, use an explicit Call and qualify the procedure name when needed.";
}
