using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Generated;

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

    internal static bool Supports(string command) => command is "table.append" or "session.save";

    internal static bool IsValidSessionId(string? sessionId) =>
        sessionId is { Length: <= 128 } && !string.IsNullOrWhiteSpace(sessionId);

    internal static Guid ValidateTracking(string command, string id)
    {
        if (!Supports(command))
            throw new ArgumentException("operationId is supported for table.append and session.save.");
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

    internal void Complete(Operation operation, ServiceResponse response)
    {
        string? recordingError = null;
        var result = ReadOutcome(operation, response);
        // Both the controller and worker can wrap a response when ownership bookkeeping fails.
        // Use the original command's outcome, retaining the bookkeeping error alongside it.
        for (var wrappers = 0; wrappers < 2 && result.Error?["category"]?.GetValue<string>() == "ServiceUnavailable"
            && result.PartialState is JsonObject retained
            && retained["outcome"]?.GetValue<string>() == "completed"
            && retained["response"]?.Deserialize<ServiceResponse>(ServiceProtocol.JsonOptions) is { } original; wrappers++)
        {
            recordingError ??= response.ErrorMessage;
            response = original;
            result = ReadOutcome(operation, original);
        }
        var resultText = result.ToToolJson();
        var state = result.MayStillBeRunning || result.Error?["category"]?.GetValue<string>() == "ExcelProcessDied"
            || LostCallbackResult(result) ? "unknown"
            : result.Success ? ProvesCompletion(operation.Command, response) ? "completed" : "unknown"
            : (result.PartialState as JsonObject)?["outcome"]?.GetValue<string>() == "partially-applied"
                ? "partially-applied" : "failed";
        Publish(operation, state, Encoding.UTF8.GetByteCount(resultText) <= MaximumResultBytes ? resultText : null,
            recordingError is not null && Encoding.UTF8.GetByteCount(recordingError) <= MaximumResultBytes ? recordingError : null);
    }

    private static bool LostCallbackResult(ResultJson result) =>
        !result.Success &&
        result.PartialState is JsonObject state &&
        state["diagnostics"] is JsonObject diagnostics &&
        diagnostics["execution"] is JsonObject execution &&
        execution["complete"] is JsonValue complete && complete.TryGetValue<bool>(out var exact) && exact &&
        execution["callbackReturned"] is JsonValue returned && returned.TryGetValue<bool>(out var completed) && completed;

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
        internal string Command { get; } = command;
        internal string SessionId { get; } = sessionId;
        internal DateTimeOffset StartedAt { get; } = startedAt;
        internal DateTimeOffset? FinishedAt { get; set; }
        internal string State { get; set; } = "pending";
        internal string? Result { get; set; }
        internal string? RecordingError { get; set; }
    }
}
