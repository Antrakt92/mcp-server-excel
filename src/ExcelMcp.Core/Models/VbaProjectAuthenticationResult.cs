using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>Observed known-password authentication; project protection settings are left untouched.</summary>
public sealed class VbaProjectAuthenticationResult : OperationResult
{
    /// <summary>Selected project whose access was inspected or authenticated.</summary>
    public string ProjectName { get; set; } = string.Empty;
    /// <summary>Current-session protection before the operation.</summary>
    public string ProtectionBefore { get; set; } = string.Empty;
    /// <summary>Current-session protection after the operation; this action does not change protection settings.</summary>
    public string ProtectionAfter { get; set; } = string.Empty;
    /// <summary>True only when a native authentication challenge was submitted.</summary>
    public bool AuthenticationPerformed { get; set; }
    /// <summary>The project was accessible before the call, so the supplied password was not checked.</summary>
    public bool AlreadyAccessible { get; set; }
    /// <summary>True for confirmed acceptance, false for a recognized rejection, null when no challenge occurred.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.Never)]
    public bool? PasswordVerified { get; set; }
    /// <summary>This call changed current-session access from locked to accessible.</summary>
    public bool AccessUnlocked { get; set; }
    /// <summary>Always false: authentication does not change the project's protection settings.</summary>
    public bool ProtectionSettingsChanged { get; set; }
    /// <summary>Always false: this action does not save the workbook.</summary>
    public bool Saved { get; set; }
    /// <summary>Always false: this action does not close or reopen the workbook.</summary>
    public bool Reopened { get; set; }
}
