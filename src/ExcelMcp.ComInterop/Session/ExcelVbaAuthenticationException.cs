namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal sealed record ExcelVbaAuthenticationState
{
    internal string Phase { get; init; } = "validating";
    internal bool ChallengeObserved { get; init; }
    internal bool AuthenticationPerformed { get; init; }
    internal bool? PasswordVerified { get; init; }
    internal bool? AccessUnlocked { get; init; } = false;
    internal bool AlreadyAccessible { get; init; }
    internal string? CurrentProtection { get; init; }
    internal bool DialogCleanupConfirmed { get; init; }
    internal bool ActiveProjectRestored { get; init; }
    internal bool VisibilityRestored { get; init; }
}

internal sealed class ExcelVbaAuthenticationProgress
{
    internal ExcelVbaAuthenticationState State { get; set; } = new();
    internal bool CanRestoreEditor { get; set; } = true;
}

internal sealed class ExcelVbaAuthenticationException(Exception cause, ExcelVbaAuthenticationState state)
    : InvalidOperationException("VBA project authentication failed during " + state.Phase + ": " + cause.Message, cause)
{
    internal ExcelVbaAuthenticationState State { get; } = state;
}
