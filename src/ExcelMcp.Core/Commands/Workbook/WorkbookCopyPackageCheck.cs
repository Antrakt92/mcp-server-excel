using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>Checks a staged Excel copy against an independent saved package before publishing it.</summary>
/// <param name="BaselinePath">Absolute path to the saved workbook used as the baseline.</param>
/// <param name="AllowedParts">Exact, case-sensitive package entry names allowed to change. An empty list permits no changes.</param>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed record WorkbookCopyPackageCheck(
    [property: JsonRequired] string BaselinePath,
    [property: JsonRequired] IReadOnlyList<string> AllowedParts);

/// <summary>The copy outcome and, when requested, its package comparison.</summary>
public sealed class WorkbookSaveCopyResult : OperationResult
{
    /// <summary>Whether the checked candidate was published at the requested path.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? FilePublished { get; init; }

    /// <summary>Complete hash-only package difference, present for a completed package check.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public WorkbookPackageComparisonResult? PackageComparison { get; init; }

    /// <summary>Machine-readable category when the package check refuses publication.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ErrorCategory { get; init; }
}
