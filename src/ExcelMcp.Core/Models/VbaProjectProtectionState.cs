namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>Nonsecret observed protection phase after failure; no rollback is implied.</summary>
public sealed class VbaProjectProtectionState
{
    /// <summary>Last attempted native or restoration phase.</summary>
    public string Phase { get; init; } = string.Empty;
    /// <summary>The owned dialog's OK control was invoked; completion can still be unknown.</summary>
    public bool AcceptanceRequested { get; init; }
    /// <summary>True after configuration read-back, false before acceptance, or null when unconfirmed.</summary>
    public bool? ConfigurationApplied { get; init; }
    /// <summary>Observed lock checkbox on a separately reopened Properties dialog.</summary>
    public bool? LockOptionReadback { get; init; }
    /// <summary>The exact introduced dialog was confirmed dismissed.</summary>
    public bool DialogCleanupConfirmed { get; init; }
    /// <summary>The previously selected VBA project was restored.</summary>
    public bool ActiveProjectRestored { get; init; }
    /// <summary>The editor's previous visibility was restored.</summary>
    public bool VisibilityRestored { get; init; }
    /// <summary>Explicit save is required if applied; null when completion is unknown.</summary>
    public bool? SaveRequired { get; init; }
    /// <summary>Explicit fresh reopen is required if applied; null when completion is unknown.</summary>
    public bool? ReopenRequired { get; init; }
    /// <summary>Always false: no workbook changes were rolled back.</summary>
    public bool RollbackAttempted { get; init; }
}

/// <summary>Native protection failure retaining its cause and final nonsecret observed state.</summary>
public sealed class VbaProjectProtectionException(string message, Exception innerException, VbaProjectProtectionState state)
    : InvalidOperationException(message, innerException)
{
    /// <summary>Observed phase and restoration state without submitted password or control values.</summary>
    public VbaProjectProtectionState State { get; } = state;
}
