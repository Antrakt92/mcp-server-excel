namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>Structural metadata of two saved packages and conservative named-style correspondence.</summary>
public sealed class WorkbookPackageStructureReport
{
    /// <summary>Structure of the first saved package.</summary>
    public required WorkbookPackageStructureSnapshot Before { get; init; }

    /// <summary>Structure of the second saved package.</summary>
    public required WorkbookPackageStructureSnapshot After { get; init; }

    /// <summary>Named parent styles matched only by name and definition fingerprint, never by position.</summary>
    public required IReadOnlyList<WorkbookPackageParentStyleMapping> ParentStyleMappings { get; init; }
}

/// <summary>Read-only metadata. Formulas, cell values and external relationship targets are excluded.</summary>
public sealed class WorkbookPackageStructureSnapshot
{
    /// <summary>The root officeDocument relationship's workbook part.</summary>
    public required string WorkbookPart { get; init; }

    /// <summary>Sheets in saved tab order, including chartsheets.</summary>
    public required IReadOnlyList<WorkbookPackageSheet> Sheets { get; init; }

    /// <summary>Package relationships with resolved internal targets.</summary>
    public required IReadOnlyList<WorkbookPackageRelationship> Relationships { get; init; }

    /// <summary>Defined-name identities and their saved sheet scope, without their formula text.</summary>
    public required IReadOnlyList<WorkbookPackageDefinedName> DefinedNames { get; init; }

    /// <summary>Named cellStyle records and the fingerprints of their referenced cellStyleXfs.</summary>
    public required IReadOnlyList<WorkbookPackageParentStyle> ParentStyles { get; init; }

    /// <summary>Non-relationship parts unreachable from package-root internal relationships. This is graph reachability, not permission to remove parts.</summary>
    public required IReadOnlyList<string> UnreachableParts { get; init; }
}

/// <summary>A sheet's saved identity and package location.</summary>
public sealed class WorkbookPackageSheet
{
    /// <summary>Zero-based saved tab position.</summary>
    public required int Position { get; init; }
    /// <summary>Saved visible sheet name.</summary>
    public required string Name { get; init; }
    /// <summary>Workbook sheetId, which is distinct from tab position and part filename.</summary>
    public required uint SheetId { get; init; }
    /// <summary>Relationship ID in the workbook relationships part.</summary>
    public required string RelationshipId { get; init; }
    /// <summary>Resolved sheet part name.</summary>
    public required string PartName { get; init; }
    /// <summary>Relationship kind, such as worksheet, chartsheet or xlMacrosheet.</summary>
    public required string SheetType { get; init; }
    /// <summary>Saved sheetPr codeName when present and captured for this part.</summary>
    public string? CodeName { get; init; }
}

/// <summary>A relationship identity, source owner and resolved internal target.</summary>
public sealed class WorkbookPackageRelationship
{
    /// <summary>Source owner part, or an empty string for package-root relationships.</summary>
    public required string OwnerPart { get; init; }
    /// <summary>Whether the source owner exists. The package root always exists.</summary>
    public required bool OwnerExists { get; init; }
    /// <summary>Relationship ID, unique within its source relationships part.</summary>
    public required string Id { get; init; }
    /// <summary>Relationship type URI.</summary>
    public required string Type { get; init; }
    /// <summary>Whether the relationship is external. External targets are never returned.</summary>
    public required bool External { get; init; }
    /// <summary>Resolved package part for internal relationships, otherwise null.</summary>
    public string? TargetPart { get; init; }
    /// <summary>Whether the internal target exists, otherwise null for external relationships.</summary>
    public bool? TargetExists { get; init; }
}

/// <summary>A defined-name identity with optional zero-based sheet scope.</summary>
public sealed class WorkbookPackageDefinedName
{
    /// <summary>Defined-name identifier, without formula text.</summary>
    public required string Name { get; init; }
    /// <summary>Zero-based saved tab scope, otherwise null for workbook scope.</summary>
    public int? LocalSheetId { get; init; }
    /// <summary>Scope's saved workbook sheetId, otherwise null for workbook scope.</summary>
    public uint? SheetId { get; init; }
    /// <summary>Scope's sheet part, otherwise null for workbook scope.</summary>
    public string? SheetPart { get; init; }
}

/// <summary>A named parent style and a conservative fingerprint of its saved definition.</summary>
public sealed class WorkbookPackageParentStyle
{
    /// <summary>The saved cellStyle name.</summary>
    public required string Name { get; init; }
    /// <summary>Zero-based parent record in cellStyleXfs.</summary>
    public required int XfId { get; init; }
    /// <summary>Canonical definition SHA-256 including referenced font, fill, border and number-format definitions. This does not prove rendered equivalence.</summary>
    public required string DefinitionSha256 { get; init; }
}

/// <summary>A conservative correspondence for a named parent style across the two packages.</summary>
public sealed class WorkbookPackageParentStyleMapping
{
    /// <summary>Saved cellStyle name.</summary>
    public required string Name { get; init; }
    /// <summary>Parent index before the edit, or null for an added name.</summary>
    public int? BeforeXfId { get; init; }
    /// <summary>Unique matching parent index after the edit, otherwise null.</summary>
    public int? AfterXfId { get; init; }
    /// <summary>One of matched, ambiguous, removed, definition-changed or added.</summary>
    public required string Status { get; init; }
    /// <summary>Distinct candidate parent indices sharing this name and definition after the edit.</summary>
    public required IReadOnlyList<int> CandidateAfterXfIds { get; init; }
}
