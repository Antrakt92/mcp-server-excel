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

    /// <summary>Optional saved VBA project comparison, independent of the ZIP allowlist.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public WorkbookVbaProjectComparison? VbaProject { get; init; }
}

/// <summary>Comparison of serialized VBA project data without exposing stream contents.</summary>
public sealed class WorkbookVbaProjectComparison
{
    /// <summary>One of compared, identical-bytes, absent, or unavailable.</summary>
    public required string Status { get; init; }

    /// <summary>Exact saved project part in the first workbook, when identified.</summary>
    public string? BeforePart { get; init; }

    /// <summary>Exact saved project part in the second workbook, when identified.</summary>
    public string? AfterPart { get; init; }

    /// <summary>Why the optional comparison was unavailable, without file contents.</summary>
    public string? Reason { get; init; }

    /// <summary>Whether the serialized reference records changed, or unknown.</summary>
    public string References { get; init; } = "unknown";

    /// <summary>Number of serialized reference records in the first VBA project, when inspected.</summary>
    public int? BeforeReferenceCount { get; init; }

    /// <summary>Number of serialized reference records in the second VBA project, when inspected.</summary>
    public int? AfterReferenceCount { get; init; }

    /// <summary>Whether the serialized protection metadata changed, or unknown.</summary>
    public string ProtectionMetadata { get; init; } = "unknown";

    /// <summary>Number of changed streams outside the recognized project categories.</summary>
    public int OtherChangedStreams { get; init; }

    /// <summary>Changed serialized items. Changes to a designer do not establish a visual or behavioral difference.</summary>
    public IReadOnlyList<WorkbookVbaProjectItemChange> Changes { get; init; } = [];
}

/// <summary>A classified item whose serialized data or presence changed.</summary>
public sealed record WorkbookVbaProjectItemChange(string Category, string Name, string ChangeKind);

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
