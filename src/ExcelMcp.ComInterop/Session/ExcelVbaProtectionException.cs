namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal sealed record ExcelVbaProtectionState
{
    internal string Phase { get; init; } = "validating";
    internal bool AcceptanceRequested { get; init; }
    internal bool? ConfigurationApplied { get; init; } = false;
    internal bool? LockOptionReadback { get; init; }
    internal bool DialogCleanupConfirmed { get; init; }
    internal bool ActiveProjectRestored { get; init; }
    internal bool VisibilityRestored { get; init; }
    internal bool? SaveRequired { get; init; } = false;
    internal bool? ReopenRequired { get; init; } = false;
}

internal sealed class ExcelVbaProtectionProgress
{
    internal ExcelVbaProtectionState State { get; set; } = new();
    internal bool CanRestoreEditor { get; set; } = true;
}

internal sealed class ExcelVbaProtectionException(Exception cause, ExcelVbaProtectionState state)
    : InvalidOperationException("VBA project protection failed during " + state.Phase + ": " + cause.Message, cause)
{
    internal ExcelVbaProtectionState State { get; } = state;
}
