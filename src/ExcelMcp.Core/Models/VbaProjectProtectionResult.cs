namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>Protection configured in Excel; explicit save and reopen remain separate operations.</summary>
public sealed class VbaProjectProtectionResult : OperationResult
{
    /// <summary>Project whose owned dialog accepted the configuration.</summary>
    public string ProjectName { get; set; } = string.Empty;
    /// <summary>Current COM protection state; the current editor can remain unlocked until reopen.</summary>
    public string CurrentProtection { get; set; } = string.Empty;
    /// <summary>The selected project dialog accepted the configuration; does not establish persisted protection.</summary>
    public bool ConfigurationApplied { get; set; }
    /// <summary>Explicitly save to persist the requested configuration.</summary>
    public bool SaveRequired { get; set; }
    /// <summary>Close and reopen explicitly before checking persisted lock state.</summary>
    public bool ReopenRequired { get; set; }
}
