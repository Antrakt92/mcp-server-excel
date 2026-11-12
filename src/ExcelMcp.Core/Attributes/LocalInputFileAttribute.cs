namespace Sbroenne.ExcelMcp.Core.Attributes;

/// <summary>Marks a local input file whose path must be resolved at the entry point.</summary>
[AttributeUsage(AttributeTargets.Parameter)]
public sealed class LocalInputFileAttribute : Attribute;
