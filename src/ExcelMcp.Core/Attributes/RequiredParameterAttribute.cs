namespace Sbroenne.ExcelMcp.Core.Attributes;

/// <summary>
/// Indicates that this parameter is required for the action.
/// Required strings reject null, empty and whitespace-only values unless explicitly allowed.
/// </summary>
[AttributeUsage(AttributeTargets.Parameter, AllowMultiple = false, Inherited = false)]
public sealed class RequiredParameterAttribute : Attribute
{
    /// <summary>Allows whitespace-only string values without allowing null or empty strings. Values are not trimmed.</summary>
    public bool AllowWhitespace { get; set; }

    /// <summary>
    /// Creates a RequiredParameter attribute.
    /// </summary>
    public RequiredParameterAttribute() { }
}
