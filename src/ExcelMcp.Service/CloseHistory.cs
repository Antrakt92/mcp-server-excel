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
    private sealed record Entry(Target Workbook, bool Owned, bool CloseConfirmed, bool SaveConfirmed, bool InitialSaveConfirmed,
        Failure? LastFailure, IReadOnlyList<Failure> RecordingFailures, JsonObject? PriorStatus,
        ExcelOperationBlocker? PriorSaveBlocker);

    private const int Capacity = 16;
    private readonly Lock _gate = new();
    private readonly LinkedList<Entry> _entries = new();

    internal (bool CloseConfirmed, bool SaveConfirmed) Record(
        Target target, ServiceResponse response, JsonObject? priorStatus = null, bool onlyIfMissing = false,
        ExcelOperationBlocker? priorSaveBlocker = null, bool owned = false)
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
        bool saved = confirmed && HasTrueFlag(close.Result, "saved");
        bool initialSaveConfirmed = saved || HasTrueFlag(close.PartialState, "initialSaveConfirmed");
        lock (_gate)
        {
            var previous = _entries.Last;
            while (previous is not null && previous.Value.Workbook != target)
                previous = previous.Previous;
            if (previous is not null)
            {
                if (onlyIfMissing)
                    return (confirmed, saved);
                _entries.Remove(previous);
            }

            // A later retry can add confirmation; it cannot erase the previous failure.
            var failure = previous?.Value.LastFailure;
            if (!close.Success && (failure is null || !HasTrueFlag(close.PartialState, "reconciledFailedClose")))
                failure = new Failure(close.ErrorCategory, close.ExceptionType, close.HResult);
            _entries.AddLast(new Entry(target, owned || previous?.Value.Owned == true,
                confirmed || previous?.Value.CloseConfirmed == true,
                saved || previous?.Value.SaveConfirmed == true,
                initialSaveConfirmed || previous?.Value.InitialSaveConfirmed == true, failure,
                recordingFailures.Count > 0 ? recordingFailures : previous?.Value.RecordingFailures ?? [],
                priorStatus ?? previous?.Value.PriorStatus, previous?.Value.PriorSaveBlocker ?? priorSaveBlocker));
            if (_entries.Count > Capacity)
                _entries.RemoveFirst();
        }
        return (confirmed, saved);
    }

    internal bool HasUnconfirmedFailure(Target target)
    {
        lock (_gate)
            return _entries.Any(entry => entry.Workbook == target && !entry.CloseConfirmed && entry.LastFailure is not null);
    }

    internal ExcelProcessIdentity? FindUniqueOwnedProcess(string sessionId)
    {
        lock (_gate)
        {
            var matches = _entries.Where(entry => entry.Owned && entry.Workbook.SessionId == sessionId)
                .Select(entry => entry.Workbook.Excel).Distinct().Take(2).ToArray();
            return matches.Length == 1 ? matches[0] : null;
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
                    ["initialSaveConfirmed"] = entry.InitialSaveConfirmed,
                    ["lastFailure"] = entry.LastFailure is { } failure ? FailureDiagnostic(failure) : null,
                    ["recordingFailures"] = new JsonArray(entry.RecordingFailures.Select(value => (JsonNode)FailureDiagnostic(value)).ToArray()),
                    ["priorStatusDiagnostic"] = entry.PriorStatus?.DeepClone(),
                    ["priorSaveBlocker"] = entry.PriorSaveBlocker is { } blocker
                        ? new JsonObject { ["kind"] = blocker.Kind, ["observedAtUtc"] = blocker.ObservedAtUtc }
                        : null,
                    ["nativeExit"] = exit is null ? null : NativeDiagnostic(exit)
                });
            }
            return result;
        }
    }

    private static bool HasTrueFlag(string? result, string flag)
    {
        if (result is null)
            return false;
        try
        {
            using var document = JsonDocument.Parse(result);
            return document.RootElement.ValueKind == JsonValueKind.Object
                && document.RootElement.TryGetProperty(flag, out var value)
                && value.ValueKind == JsonValueKind.True;
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

    // An exact exit probe can arrive after a failed exit-code read. Keep that
    // failure until a matching observation supplies a numeric exit code.
    internal static bool IsStrongerNativeExit(ExcelNativeProcessObservation candidate,
        ExcelNativeProcessObservation? previous) =>
        candidate.State == ExcelNativeProcessState.Exited
        && (previous is null || previous.Identity == candidate.Identity
            && (previous.ExitCode is null
                && (candidate.ExitCode is not null || previous.NativeError is null && candidate.NativeError is not null)
                || previous.ExitCode == candidate.ExitCode && previous.NativeError == candidate.NativeError
                    && previous.ExitedAtUtc is null && candidate.ExitedAtUtc is not null));

    internal static JsonObject NativeDiagnostic(ExcelNativeProcessObservation observation) => new()
    {
        ["processId"] = observation.Identity.ProcessId,
        ["startedAtUtcFileTime"] = observation.Identity.StartedAtUtcFileTime.ToString(CultureInfo.InvariantCulture),
        ["state"] = observation.State.ToString(),
        ["exitCode"] = observation.ExitCode,
        ["nativeError"] = observation.NativeError
    };
}
