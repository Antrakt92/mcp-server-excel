using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Generated;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Uses shared operation-outcome rules for telemetry, including authoritative
/// nested failures, without allowing malformed payloads to interrupt reporting.
/// </summary>
internal static class ServiceResultOutcome
{
    internal static bool TryReadNegative(
        string? resultJson,
        out string? errorMessage,
        out string? errorCategory,
        string command = "invocation")
    {
        errorMessage = null;
        errorCategory = null;
        var contract = ServiceRegistry.GetResultContract(command);
        if (string.IsNullOrWhiteSpace(resultJson) &&
            contract is not { ReadOnly: true } &&
            !command.Equals("session.test", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        try
        {
            var outcome = ResultJson.Read(resultJson, command, contract: contract);
            errorMessage = outcome.Error?["message"]?.GetValue<string>();
            errorCategory = outcome.Error?["category"]?.GetValue<string>();
            return !outcome.Success;
        }
        catch (JsonException exception)
        {
            errorMessage = exception.Message;
            errorCategory = "InvalidResponse";
            return true;
        }
        catch (ArgumentException exception)
        {
            // JsonObject can reject duplicate response properties on first access.
            errorMessage = exception.Message;
            errorCategory = "InvalidResponse";
            return true;
        }
    }
}
