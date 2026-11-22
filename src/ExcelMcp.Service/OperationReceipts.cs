using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.Service.Workers;

namespace Sbroenne.ExcelMcp.Service;

/// <summary>Short-lived, process-local evidence for explicitly tracked workbook changes.</summary>
internal sealed class OperationReceipts(TimeProvider? timeProvider = null, int capacity = 128)
{
    internal const string LookupCommand = "session.operation-result";
    private static readonly TimeSpan Retention = TimeSpan.FromMinutes(15);
    private const int MaximumResultBytes = 8192;
    private readonly Lock _gate = new();
    private readonly TimeProvider _timeProvider = timeProvider ?? TimeProvider.System;
    private readonly Dictionary<Guid, Operation> _entries = [];

    internal static bool Supports(string command) => command is "table.append" or "session.save" or "vba.invoke";

    internal static bool IsValidSessionId(string? sessionId) =>
        sessionId is { Length: <= 128 } && !string.IsNullOrWhiteSpace(sessionId);

    internal static Guid ValidateTracking(string command, string id)
    {
        if (!Supports(command))
            throw new ArgumentException("operationId is supported for table.append, session.save, and vba.invoke.");
        return ParseId(id);
    }

    internal static Guid ParseId(string? id) =>
        (Guid.TryParseExact(id, "D", out var parsed) || Guid.TryParseExact(id, "N", out parsed)) && parsed != Guid.Empty
            ? parsed
            : throw new ArgumentException("operationId must be a nonzero UUID.");

    internal enum Admission { Accepted, Duplicate, Full }

    internal Admission TryBegin(Guid id, string command, string sessionId, out Operation operation)
    {
        lock (_gate)
        {
            var now = _timeProvider.GetUtcNow();
            RemoveExpired(now);
            if (_entries.ContainsKey(id))
            {
                operation = null!;
                return Admission.Duplicate;
            }
            if (_entries.Count >= capacity)
            {
                operation = null!;
                return Admission.Full;
            }
            operation = new(command, sessionId, now);
            _entries.Add(id, operation);
            return Admission.Accepted;
        }
    }

    internal void BindWorker(Operation operation, WorkerRegistration registration, ExcelNativeProcessObservation? observedExit)
    {
        lock (_gate)
        {
            operation.Worker = registration;
            if (observedExit is { State: ExcelNativeProcessState.Exited } && observedExit.Identity == registration.Excel)
                operation.NativeExit = observedExit;
        }
    }

    internal void ObserveWorkerEvidence(WorkerRegistration registration)
    {
        if (registration.OperationEvidence is not { } checkpoint || !checkpoint.IsValidFor(registration)
            || !Guid.TryParseExact(checkpoint.OperationId, "N", out var id)
            || !Guid.TryParseExact(checkpoint.DispatchId, "N", out var dispatch))
            return;

        lock (_gate)
        {
            if (!_entries.TryGetValue(id, out var operation)
                || operation.DispatchId != dispatch
                || operation.Worker is not { } source
                || source.PipeName != registration.PipeName || source.Worker != registration.Worker
                || source.Controller != registration.Controller || source.Generation != registration.Generation
                || source.Build != registration.Build
                || operation.SessionId != checkpoint.SessionId || operation.Command != checkpoint.Command
                || source.Excel != checkpoint.Excel)
                return;

            if (operation.Checkpoint?.Phase != WorkerOperationEvidence.ResponsePhase)
                operation.Checkpoint = checkpoint;
        }
    }

    internal void ObserveNativeExit(ExcelNativeProcessObservation observation)
    {
        if (observation.State != ExcelNativeProcessState.Exited)
            return;

        lock (_gate)
        {
            foreach (var operation in _entries.Values)
                if (operation.Worker?.Excel == observation.Identity && operation.NativeExit?.ExitCode is null)
                    operation.NativeExit = observation;
        }
    }

    internal void ObserveInitialFailure(Operation operation, ServiceResponse response)
    {
        for (var wrappers = 0; wrappers < 2 && response.ReadCompletionFailure() is { } retained; wrappers++)
            response = retained.Response;
        if (response.Success)
            return;

        lock (_gate)
            operation.InitialFailure ??= new(response.ErrorCategory, response.ExceptionType, response.HResult);
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "A malformed partial-state payload leaves the exact original response available to the caller. Only the optional diagnostic overlay is omitted.")]
    internal ServiceResponse IncludeEvidenceInFailure(Operation operation, ServiceResponse response)
    {
        if (response.Success)
            return response;

        JsonObject evidence;
        lock (_gate)
            evidence = FormatEvidence(operation);

        try
        {
            var state = response.PartialState is null ? new JsonObject() : JsonNode.Parse(response.PartialState) as JsonObject;
            if (state is null || state["diagnostics"] is { } existing && existing is not JsonObject)
                return response;
            if (state["diagnostics"] is not JsonObject diagnostics)
            {
                diagnostics = new JsonObject();
                state["diagnostics"] = diagnostics;
            }
            diagnostics["operationEvidence"] = evidence;
            return response.WithPartialState(state.ToJsonString(ServiceProtocol.JsonOptions));
        }
        catch (Exception ex) when (ex is JsonException or ArgumentException)
        {
            // Optional diagnostics must not replace the original failure.
            return response;
        }
    }

    private static JsonObject FormatEvidence(Operation operation)
    {
        var result = new JsonObject
        {
            ["phase"] = operation.Checkpoint?.Phase
                ?? (operation.Worker is null ? "controllerAccepted" : "controllerDispatchStarted")
        };
        if (operation.Worker?.Excel is { } excel)
        {
            result["processId"] = excel.ProcessId;
            result["startedAtUtcFileTime"] = excel.StartedAtUtcFileTime.ToString(CultureInfo.InvariantCulture);
        }
        if (operation.Checkpoint is { ErrorCategory: not null } checkpoint)
            result["workerFailure"] = FormatFailure(new(checkpoint.ErrorCategory, checkpoint.ExceptionType, checkpoint.HResult));
        if (operation.InitialFailure is { } initial)
            result["initialFailure"] = FormatFailure(initial);
        if (operation.NativeExit is { } nativeExit)
            result["nativeExit"] = CloseHistory.NativeDiagnostic(nativeExit);
        return result;
    }

    private static JsonObject FormatFailure(Failure failure) => new()
    {
        ["category"] = failure.Category,
        ["exceptionType"] = failure.ExceptionType,
        ["hresult"] = failure.HResult
    };

    internal sealed record Failure(string? Category, string? ExceptionType, string? HResult);

    internal void Complete(Operation operation, ServiceResponse response, bool forceUnknown = false)
    {
        string? recordingError = null;
        // Keep the primary response even if bookkeeping failed while completion was unknown.
        for (var wrappers = 0; wrappers < 2 && response.ReadCompletionFailure() is { } retained; wrappers++)
        {
            recordingError ??= response.ErrorMessage;
            forceUnknown |= !retained.CompletionKnown;
            response = retained.Response;
        }
        ObserveInitialFailure(operation, response);
        var result = ReadOutcome(operation, response);
        var resultText = result.ToToolJson();
        var state = forceUnknown || result.CompletionUnconfirmed ? "unknown"
            : result.Success ? ProvesCompletion(operation.Command, response) ? "completed" : "unknown"
            : (result.PartialState as JsonObject)?["outcome"]?.GetValue<string>() == "partially-applied"
                ? "partially-applied" : "failed";
        Publish(operation, state, Encoding.UTF8.GetByteCount(resultText) <= MaximumResultBytes ? resultText : null,
            recordingError is not null && Encoding.UTF8.GetByteCount(recordingError) <= MaximumResultBytes ? recordingError : null);
    }

    private static bool ProvesCompletion(string command, ServiceResponse response)
    {
        if (response.Result is null)
            return false;
        using var result = JsonDocument.Parse(response.Result);
        var property = command == "session.save" ? "saved" : "success";
        return result.RootElement.ValueKind == JsonValueKind.Object &&
            result.RootElement.TryGetProperty(property, out var value) && value.ValueKind == JsonValueKind.True;
    }

    private static ResultJson ReadOutcome(Operation operation, ServiceResponse response) =>
        ResultJson.Read(response.Result, operation.Command, response.Success, response.ErrorMessage,
            response.ErrorCategory, response.SessionId, response.ExceptionType, response.HResult, response.InnerError,
            response.PartialState, ServiceRegistry.GetResultContract(operation.Command));

    internal void Publish(Operation operation, string state, string? result = null, string? recordingError = null)
    {
        lock (_gate)
        {
            if (operation.FinishedAt is not null)
                return;
            operation.State = state;
            operation.Result = result;
            operation.RecordingError = recordingError;
            if (state != "pending")
                operation.FinishedAt = _timeProvider.GetUtcNow();
        }
    }

    internal ServiceResponse Lookup(string? args, string? sessionId)
    {
        using var document = JsonDocument.Parse(args ?? "{}");
        if (document.RootElement.ValueKind != JsonValueKind.Object)
            throw new ArgumentException("operation-result arguments must be an object.");
        var operationId = document.RootElement.TryGetProperty("operationId", out var value)
            && value.ValueKind == JsonValueKind.String ? value.GetString() : null;
        var id = ParseId(operationId);
        if (!IsValidSessionId(sessionId))
            throw new ArgumentException("A valid sessionId is required for operation-result.");

        lock (_gate)
        {
            RemoveExpired(_timeProvider.GetUtcNow());
            if (!_entries.TryGetValue(id, out var operation) || operation.SessionId != sessionId)
                return Format(new
                {
                    operationId = id,
                    state = "unknown",
                    message = "No matching receipt is retained by this host. Inspect the workbook before repeating the command."
                });
            return Format(new
            {
                operationId = id,
                operation.Command,
                operation.State,
                operation.StartedAt,
                operation.FinishedAt,
                operation.RecordingError,
                operationEvidence = FormatEvidence(operation),
                result = operation.Result is null ? null : JsonNode.Parse(operation.Result)
            });
        }
    }

    private void RemoveExpired(DateTimeOffset now)
    {
        foreach (var entry in _entries.Where(pair =>
            pair.Value.FinishedAt is { } finished && now - finished >= Retention).ToArray())
            _entries.Remove(entry.Key);
    }

    private static ServiceResponse Format<T>(T result) => new()
    {
        Success = true,
        Result = JsonSerializer.Serialize(result, ServiceProtocol.JsonOptions)
    };

    internal sealed class Operation(string command, string sessionId, DateTimeOffset startedAt)
    {
        // A public UUID may be reused after retention expires. A worker snapshot still needs this dispatch's identity.
        internal Guid DispatchId { get; } = Guid.NewGuid();
        internal string Command { get; } = command;
        internal string SessionId { get; } = sessionId;
        internal DateTimeOffset StartedAt { get; } = startedAt;
        internal DateTimeOffset? FinishedAt { get; set; }
        internal string State { get; set; } = "pending";
        internal string? Result { get; set; }
        internal string? RecordingError { get; set; }
        internal WorkerRegistration? Worker { get; set; }
        internal WorkerOperationEvidence? Checkpoint { get; set; }
        internal ExcelNativeProcessObservation? NativeExit { get; set; }
        internal Failure? InitialFailure { get; set; }
    }
}
