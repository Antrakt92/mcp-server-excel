using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Generated;

namespace Sbroenne.ExcelMcp.Service;

/// <summary>Rejects audit-session mutations before entering Excel or changing security state.</summary>
internal static class AuditSessionPolicy
{
    internal static void EnsureAllowed(string command, string? args)
    {
        if (command is "session.notices" or "session.list")
            return;
        if (command == "session.close")
        {
            using var document = JsonDocument.Parse(args ?? "{}");
            if (!document.RootElement.TryGetProperty("save", out var save) ||
                save.ValueKind == JsonValueKind.False)
                return;
        }
        else if (ServiceRegistry.GetResultContract(command) is { ReadOnly: true })
            return;

        throw new OperationFailureException(OperationFailureCategory.Permissions,
            $"Audit session rejected '{command}'. Workbook changes, calculation, macro execution, " +
            "security changes and saving are disabled. Close without saving when inspection finishes; " +
            "open a normal session only when editing is authorized.");
    }
}
