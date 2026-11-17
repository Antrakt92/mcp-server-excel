using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>Complete comparison of two saved workbook packages.</summary>
public sealed class WorkbookPackageComparisonResult : ResultBase
{
    /// <summary>Number of package entries in the first workbook.</summary>
    public required int BeforePartCount { get; init; }

    /// <summary>Number of package entries in the second workbook.</summary>
    public required int AfterPartCount { get; init; }

    /// <summary>Whether every added, removed, or changed entry was explicitly allowed.</summary>
    public required bool MatchesAllowlist { get; init; }

    /// <summary>Added, removed, and changed entries, ordered by their exact names.</summary>
    public required IReadOnlyList<WorkbookPackagePartChange> Changes { get; init; }

    /// <summary>Optional saved-package structural diagnostics. These do not change allowlist decisions.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public WorkbookPackageStructureReport? Structure { get; init; }
}

/// <summary>A saved workbook package entry whose presence or decompressed bytes changed.</summary>
public sealed class WorkbookPackagePartChange
{
    /// <summary>Exact case-sensitive ZIP entry name without a leading slash.</summary>
    public required string PartName { get; init; }

    /// <summary>One of added, removed, or changed.</summary>
    public required string ChangeKind { get; init; }

    /// <summary>Lowercase SHA-256 of the first workbook's entry, when present.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.Never)]
    public required string? BeforeSha256 { get; init; }

    /// <summary>Lowercase SHA-256 of the second workbook's entry, when present.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.Never)]
    public required string? AfterSha256 { get; init; }

    /// <summary>Whether the exact entry name was in the supplied allowlist.</summary>
    public required bool Allowed { get; init; }
}
