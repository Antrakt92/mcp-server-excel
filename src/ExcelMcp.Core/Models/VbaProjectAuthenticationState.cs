using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>Nonsecret observed authentication state after failure; cancellation is not rollback.</summary>
public sealed class VbaProjectAuthenticationState
{
    /// <summary>Last attempted native or restoration phase.</summary>
    public string Phase { get; init; } = string.Empty;
    /// <summary>Whether an authentication dialog was observed.</summary>
    public bool ChallengeObserved { get; init; }
    /// <summary>Whether the exact owned authentication OK button was requested.</summary>
    public bool AuthenticationPerformed { get; init; }
    /// <summary>Observed acceptance/rejection; null when not challenged or outcome is unknown.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.Never)]
    public bool? PasswordVerified { get; init; }
    /// <summary>Whether current-session access changed; null when outcome remains unknown.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.Never)]
    public bool? AccessUnlocked { get; init; }
    /// <summary>The project was already accessible and the password was not checked.</summary>
    public bool AlreadyAccessible { get; init; }
    /// <summary>Last protection state read through the selected project's COM object.</summary>
    public string? CurrentProtection { get; init; }
    /// <summary>Every introduced expected dialog closed and command execution returned.</summary>
    public bool DialogCleanupConfirmed { get; init; }
    /// <summary>The previously selected project was restored.</summary>
    public bool ActiveProjectRestored { get; init; }
    /// <summary>The editor's previous visibility was restored.</summary>
    public bool VisibilityRestored { get; init; }
    /// <summary>Always false: no protection settings are changed by authentication.</summary>
    public bool ProtectionSettingsChanged { get; init; }
    /// <summary>Always false: no save is performed.</summary>
    public bool Saved { get; init; }
    /// <summary>Always false: no close/reopen is performed.</summary>
    public bool Reopened { get; init; }
    /// <summary>Always false: no workbook changes or current access were rolled back.</summary>
    public bool RollbackAttempted { get; init; }
}

/// <summary>Authentication failure preserving native cause and observed nonsecret state.</summary>
public sealed class VbaProjectAuthenticationException(string message, Exception innerException,
    VbaProjectAuthenticationState state) : InvalidOperationException(message, innerException)
{
    /// <summary>Observed outcome, cleanup and restoration state; never contains password data.</summary>
    public VbaProjectAuthenticationState State { get; } = state;
}
