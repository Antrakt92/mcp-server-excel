namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>Observed module state after a source edit failed; source bodies are never copied here.</summary>
public sealed class VbaSourceEditState
{
    /// <summary>Selected component, or actual created component after import; empty if creation identity could not be confirmed.</summary>
    public string ModuleName { get; init; } = string.Empty;
    /// <summary>Last attempted stage: creating, naming, deleting, inserting, or verifying.</summary>
    public string Phase { get; init; } = string.Empty;
    /// <summary>Fingerprint read back after failure, or null when unavailable.</summary>
    public string? SourceHash { get; init; }
    /// <summary>Observed source line count, or null when unavailable.</summary>
    public int? LineCount { get; init; }
    /// <summary>Independent read-back failure, when the remaining source could not be inspected.</summary>
    public string? ReadBackError { get; init; }
}

/// <summary>A source edit failed after mutation started; retains its original cause and observed state.</summary>
public sealed class VbaSourceEditException(string message, Exception innerException, VbaSourceEditState state)
    : InvalidOperationException(message, innerException)
{
    /// <summary>Observed state without a recovery edit or automatic rollback.</summary>
    public VbaSourceEditState State { get; } = state;
}

/// <summary>Verified read-back of a whole-module replacement, without a compilation claim.</summary>
public sealed class VbaModuleEditResult : OperationResult
{
    /// <summary>Edited component name.</summary>
    public string ModuleName { get; set; } = string.Empty;
    /// <summary>Fingerprint of the source read back from Excel.</summary>
    public string SourceHash { get; set; } = string.Empty;
    /// <summary>Number of source lines read back.</summary>
    public int LineCount { get; set; }
}

/// <summary>One explicit VBA argument; missing is distinct from a VBA Null.</summary>
public sealed class VbaArgument
{
    /// <summary>string, number, boolean, null, or missing.</summary>
    public string Type { get; set; } = string.Empty;
    /// <summary>Scalar matching Type; null and missing must omit it or specify null.</summary>
    public object? Value { get; set; }
}

/// <summary>A macro return value limited to transportable scalar/null or rectangular values.</summary>
public sealed class VbaInvokeResult : OperationResult
{
    /// <summary>string, number, boolean, null, or array.</summary>
    public string ReturnKind { get; set; } = string.Empty;
    /// <summary>Scalar return value, including null.</summary>
    [System.Text.Json.Serialization.JsonIgnore(Condition = System.Text.Json.Serialization.JsonIgnoreCondition.Never)]
    public object? Value { get; set; }
    /// <summary>Rectangular two-dimensional return value when ReturnKind is array.</summary>
    public List<List<object?>>? Values { get; set; }
}
