using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Service;

/// <summary>Bounded diagnostics for close requests. History never authorizes release or confirms a save.</summary>
internal sealed class CloseHistory
{
    internal sealed record Target(string SessionId, ExcelProcessIdentity? Excel);
    private sealed record Failure(string? Category, string? ExceptionType, string? HResult);
    private sealed record Entry(Target Workbook, bool CloseConfirmed, bool SaveConfirmed,
        Failure? LastFailure, IReadOnlyList<Failure> RecordingFailures, JsonObject? PriorStatus);

    private const int Capacity = 16;
    private readonly Lock _gate = new();
    private readonly LinkedList<Entry> _entries = new();

    internal void Record(Target target, ServiceResponse response, JsonObject? priorStatus = null)
    {
        var recordingFailures = new List<Failure>();
        var close = response;
        bool confirmed = true;
        // The worker and controller may each wrap the original close in a failed completion publication.
        while (recordingFailures.Count < 2 && close.ReadCompletionFailure() is { } retained)
        {
            recordingFailures.Add(new(close.ErrorCategory, close.ExceptionType, close.HResult));
            confirmed &= retained.CompletionKnown;
            close = retained.Response;
        }

        confirmed &= close.Success;
        bool saved = confirmed && HasSaveConfirmation(close.Result);
        lock (_gate)
        {
            var previous = _entries.Last;
            while (previous is not null && previous.Value.Workbook != target)
                previous = previous.Previous;
            if (previous is not null)
                _entries.Remove(previous);

            // A later retry can add confirmation; it cannot erase the previous failure.
            var failure = close.Success ? previous?.Value.LastFailure
                : new Failure(close.ErrorCategory, close.ExceptionType, close.HResult);
            _entries.AddLast(new Entry(target, confirmed || previous?.Value.CloseConfirmed == true,
                saved || previous?.Value.SaveConfirmed == true, failure,
                recordingFailures.Count > 0 ? recordingFailures : previous?.Value.RecordingFailures ?? [],
                priorStatus ?? previous?.Value.PriorStatus));
            if (_entries.Count > Capacity)
                _entries.RemoveFirst();
        }
    }

    internal JsonArray Snapshot(IEnumerable<ExcelNativeProcessObservation> nativeExits, Func<string, bool> sessionReleased)
    {
        lock (_gate)
        {
            var result = new JsonArray();
            foreach (var entry in _entries)
            {
                var process = entry.Workbook.Excel;
                var exit = process is null ? null : nativeExits.FirstOrDefault(observation => observation.Identity == process);
                result.Add(new JsonObject
                {
                    ["sessionId"] = entry.Workbook.SessionId,
                    ["processId"] = process?.ProcessId,
                    ["startedAtUtcFileTime"] = process?.StartedAtUtcFileTime.ToString(CultureInfo.InvariantCulture),
                    ["closeConfirmed"] = entry.CloseConfirmed,
                    ["sessionReleased"] = sessionReleased(entry.Workbook.SessionId),
                    ["saveConfirmed"] = entry.SaveConfirmed,
                    ["lastFailure"] = entry.LastFailure is { } failure ? FailureDiagnostic(failure) : null,
                    ["recordingFailures"] = new JsonArray(entry.RecordingFailures.Select(value => (JsonNode)FailureDiagnostic(value)).ToArray()),
                    ["priorStatusDiagnostic"] = entry.PriorStatus?.DeepClone(),
                    ["nativeExit"] = exit is null ? null : NativeDiagnostic(exit)
                });
            }
            return result;
        }
    }

    private static bool HasSaveConfirmation(string? result)
    {
        if (result is null)
            return false;
        try
        {
            using var document = JsonDocument.Parse(result);
            return document.RootElement.ValueKind == JsonValueKind.Object
                && document.RootElement.TryGetProperty("saved", out var saved)
                && saved.ValueKind == JsonValueKind.True;
        }
        catch (JsonException)
        {
            return false;
        }
    }

    private static JsonObject FailureDiagnostic(Failure failure) => new()
    {
        ["category"] = failure.Category,
        ["exceptionType"] = failure.ExceptionType,
        ["hresult"] = failure.HResult
    };

    internal static JsonObject NativeDiagnostic(ExcelNativeProcessObservation observation) => new()
    {
        ["processId"] = observation.Identity.ProcessId,
        ["startedAtUtcFileTime"] = observation.Identity.StartedAtUtcFileTime.ToString(CultureInfo.InvariantCulture),
        ["state"] = observation.State.ToString(),
        ["exitCode"] = observation.ExitCode,
        ["nativeError"] = observation.NativeError
    };
}
