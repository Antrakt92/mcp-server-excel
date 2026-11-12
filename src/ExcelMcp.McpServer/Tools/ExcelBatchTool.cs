using System.Buffers;
using System.ComponentModel;
using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.McpServer.Telemetry;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

[JsonConverter(typeof(JsonStringEnumConverter<BatchAction>))]
public enum BatchAction
{
    [JsonStringEnumMemberName("run")]
    Run
}

/// <summary>One command in the same format as an excelcli batch entry.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class BatchCommandEntry
{
    /// <summary>CLI group.action name, for example range.get-values or session.open.</summary>
    [JsonPropertyName("command")]
    [Description("CLI group.action name, for example range.get-values or session.open.")]
    public string Command { get; init; } = string.Empty;

    /// <summary>Optional session override for this entry, outside args.</summary>
    [JsonPropertyName("sessionId")]
    [Description("Optional session override for this entry, outside args.")]
    public string? SessionId { get; init; }

    /// <summary>Command arguments using Service camelCase keys, not MCP snake_case inputs or CLI flags.</summary>
    [JsonPropertyName("args")]
    [Description("Command arguments using Service camelCase keys, for example sheetName and rangeAddress.")]
    public JsonElement? Args { get; init; }
}

[McpServerToolType]
public static partial class ExcelBatchTool
{
    private const int MaximumCommands = 128;
    private const int ResultBudgetBytes = 8 * 1024 * 1024;

    /// <summary>
    /// Run a known sequence of commands one at a time in one MCP call.
    /// Uses the excelcli batch entry format: command is group.action (range.get-values, sheet.list, session.open),
    /// args uses Service camelCase keys (sheetName, rangeAddress), and each entry may override sessionId.
    /// Local code, data and formula input files require absolute Windows paths.
    /// Only commands exposed by MCP tools are accepted; service lifecycle commands and nested batches are rejected.
    /// Tool action and parameter descriptions give the exact batch command and argument names.
    /// Supply JSON arrays and objects directly, not strings containing JSON.
    /// workbook_session_id supplies the default session. Without one, successful session.open/create/attach
    /// captures the returned workbook session for later entries; closing that default session clears it.
    /// Other returned values are not substituted. Use individual tools when later steps need inspection.
    /// Continues after errors by default; set stop_on_error:true for dependent steps.
    /// Timeouts, cancellation, and reported unknown completion always stop further commands, because Excel work may still be running.
    /// After 8 MiB of result JSON, the batch keeps the last result intact and stops before another command.
    /// If commands remain, the reply reports stopReason:result-budget. Inspect the returned results before continuing.
    /// Results contain each executed or rejected entry's zero-based index, command, success and normal MCP JSON fields.
    /// totalCount, executedCount and stoppedEarly report how much of the list was processed.
    /// Any failed entry sets the batch's success:false and MCP error flag. Cancellation stops further commands.
    /// Does not roll back after failure or cancellation. Individual commands can save or discard edits.
    /// Does not automatically undo, save, close, or change calculation mode.
    /// Excel AutoSave can still persist edits. Inspect results and session state before retrying.
    /// Screenshots in a batch return JSON data; use the screenshot tool for native image content.
    /// </summary>
    /// <param name="action">Run the supplied sequence.</param>
    /// <param name="commands">1 to 128 entries in CLI batch format: [{"command":"range.get-values","args":{"sheetName":"Sales","rangeAddress":"A1:B2"}}].</param>
    /// <param name="workbook_session_id">Optional default workbook session; each entry's sessionId overrides it.</param>
    /// <param name="stop_on_error">Stop after the first failed entry. Default false, matching excelcli batch.</param>
    [McpServerTool(Name = "batch", Title = "Batch Operations", Destructive = true, Idempotent = false)]
    [McpMeta("category", "batch")]
    [McpMeta("requiresSession", false)]
    public static partial Task<CallToolResult> ExcelBatch(
        BatchAction action,
        ServiceBridge.ServiceBridge bridge,
        [McpActionParameter("run", Required = true)] BatchCommandEntry[] commands,
        [McpActionParameter("run"), DefaultValue(null)] string? workbook_session_id = null,
        [McpActionParameter("run"), DefaultValue(false)] bool stop_on_error = false,
        CancellationToken cancellationToken = default) =>
        ExecuteBatchAsync(action, bridge, commands, workbook_session_id, stop_on_error, cancellationToken,
            ServiceRegistry.ValidateCommandArguments, ExcelMcpTelemetry.TrackToolInvocation);

    internal static Task<CallToolResult> ExecuteBatchAsync(
        BatchAction action,
        ServiceBridge.ServiceBridge bridge,
        BatchCommandEntry[] commands,
        string? workbook_session_id,
        bool stop_on_error,
        CancellationToken cancellationToken,
        Action<string, string?> validateArguments,
        Action<string, string, long, ToolInvocationResult> trackInvocation) =>
        ExcelToolsBase.ExecuteToolActionAsync("batch", "run", async () =>
        {
            if (action != BatchAction.Run)
                throw new ArgumentOutOfRangeException(nameof(action));
            if (commands is null || commands.Length == 0)
                throw new ArgumentException("commands must contain at least one command.");
            if (commands.Length > MaximumCommands)
                throw new ArgumentException($"commands must contain at most {MaximumCommands} entries.");
            for (var index = 0; index < commands.Length; index++)
            {
                if (commands[index] is null || string.IsNullOrWhiteSpace(commands[index].Command))
                    throw new ArgumentException($"Command at index {index} must be an object with a non-empty command field.");
            }

            var buffer = new ArrayBufferWriter<byte>();
            using var writer = new Utf8JsonWriter(buffer);
            writer.WriteStartObject();
            writer.WriteStartArray("results");
            var resultsStart = writer.BytesCommitted + writer.BytesPending;
            var executedCount = 0;
            var resultBudgetReached = false;
            string? activeSession = workbook_session_id;
            string? firstFailureCategory = null;
            for (var index = 0; index < commands.Length; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var entry = commands[index];
                var operation = McpActionContract.GetOperation(entry.Command);
                var sessionId = operation is { RequiresSession: true } ? entry.SessionId ?? activeSession : null;
                var outcome = await ExecuteCommandAsync(
                    bridge, entry, operation, sessionId, cancellationToken, validateArguments, trackInvocation);

                if (outcome.Success && activeSession is null &&
                    (entry.Command.Equals("session.open", StringComparison.OrdinalIgnoreCase) ||
                     entry.Command.Equals("session.create", StringComparison.OrdinalIgnoreCase) ||
                     entry.Command.Equals("session.attach", StringComparison.OrdinalIgnoreCase)))
                {
                    activeSession = outcome.Data?["sessionId"]?.GetValue<string>();
                }
                if (outcome.Success &&
                    entry.Command.Equals("session.close", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(sessionId, activeSession, StringComparison.OrdinalIgnoreCase))
                    activeSession = null;

                var filePath = entry.Args is { ValueKind: JsonValueKind.Object } args &&
                    args.TryGetProperty("filePath", out var path) && path.ValueKind == JsonValueKind.String
                    ? path.GetString() : null;
                var item = operation is { Tool: "file" or "file_read" } fileOperation
                    ? ExcelFileTool.ConvertSessionResult(outcome, fileOperation.Action, filePath)
                    : JsonNode.Parse(outcome.ToToolJson())!.AsObject();
                item["index"] = index;
                item["command"] = entry.Command;
                item.WriteTo(writer, ExcelToolsBase.JsonOptions);
                executedCount++;
                if (!outcome.Success)
                {
                    firstFailureCategory ??= outcome.Error?["category"]?.GetValue<string>() ?? "OperationFailed";
                    if (stop_on_error || outcome.MayStillBeRunning)
                        break;
                }
                // Keep this complete result, including failure evidence. The budget gates the next dispatch.
                if (executedCount < commands.Length &&
                    writer.BytesCommitted + writer.BytesPending - resultsStart >= ResultBudgetBytes)
                {
                    resultBudgetReached = true;
                    break;
                }
            }

            writer.WriteEndArray();
            writer.WriteBoolean("success", firstFailureCategory is null && !resultBudgetReached);
            writer.WriteNumber("totalCount", commands.Length);
            writer.WriteNumber("executedCount", executedCount);
            writer.WriteBoolean("stoppedEarly", executedCount < commands.Length);
            if (resultBudgetReached)
            {
                writer.WriteString("stopReason", "result-budget");
                writer.WriteString("errorMessage", "The batch reached its 8 MiB result budget. Inspect the returned results and session state before continuing with the remaining commands.");
                writer.WriteString("errorCategory", firstFailureCategory ?? "ResourceLimit");
            }
            else if (firstFailureCategory is not null)
            {
                writer.WriteString("errorMessage", "One or more batch commands failed. The batch does not roll back; individual commands can save or discard edits. Inspect results before continuing.");
                writer.WriteString("errorCategory", firstFailureCategory);
            }
            writer.WriteEndObject();
            writer.Flush();
            return Encoding.UTF8.GetString(buffer.WrittenSpan);
        }, cancellationToken, static (_, _, _, _) => { });

    private static async Task<ResultJson> ExecuteCommandAsync(
        ServiceBridge.ServiceBridge bridge,
        BatchCommandEntry entry,
        (string Tool, string Action, bool RequiresSession)? operation,
        string? sessionId,
        CancellationToken cancellationToken,
        Action<string, string?> validateArguments,
        Action<string, string, long, ToolInvocationResult> trackInvocation)
    {
        var tool = operation?.Tool ?? "other";
        var action = operation?.Action ?? "other";
        var command = operation is { } known
            ? ServiceRegistry.GetMcpCommand(known.Tool, known.Action)
            : entry.Command;
        var stopwatch = Stopwatch.StartNew();
        var invocation = new ToolInvocationResult(ToolInvocationOutcome.Failed, ToolFailureClass.Unclassified);
        try
        {
            var validationFailure = await Task.Run(() => ValidateCommand(entry, operation, command, sessionId, validateArguments),
                cancellationToken).WaitAsync(cancellationToken);
            var outcome = validationFailure ?? await ExcelToolsBase.ForwardResultToServiceAsync(
                bridge, command, sessionId, entry.Args, cancellationToken);
            invocation = ExcelToolsBase.ClassifyToolResponse(tool, action, outcome.ToToolJson());
            return outcome;
        }
        catch (Exception exception)
        {
            invocation = ExcelToolsBase.ClassifyToolException(exception);
            throw;
        }
        finally
        {
            trackInvocation(tool, action, stopwatch.ElapsedMilliseconds, invocation);
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Input validation failures become explicit failed items and set the aggregate MCP error flag. Execution and cancellation are outside this boundary.")]
    private static ResultJson? ValidateCommand(
        BatchCommandEntry entry,
        (string Tool, string Action, bool RequiresSession)? operation,
        string command,
        string? sessionId,
        Action<string, string?> validateArguments)
    {
        try
        {
            if (operation is null)
                throw new ArgumentException($"Command '{entry.Command}' is not available in MCP batches. Use the batch command and argument names in MCP tool descriptions.");
            if (entry.SessionId is not null && string.IsNullOrWhiteSpace(entry.SessionId))
                throw new ArgumentException("sessionId must be a non-empty string.");
            if (entry.Args is { ValueKind: not (JsonValueKind.Object or JsonValueKind.Null) })
                throw new ArgumentException("args must be a JSON object or null.");
            var argsJson = entry.Args?.GetRawText();
            validateArguments(command, argsJson);
            ServiceRegistry.BindLocalInputFileArguments(command, argsJson, callerDirectory: null);
            return null;
        }
        catch (Exception ex) when (ex is ArgumentException or JsonException or IOException or UnauthorizedAccessException)
        {
            return ResultJson.Read(null, entry.Command, delivered: false,
                message: ex.Message, category: "InvalidInput", sessionId: sessionId,
                exceptionType: ex.GetType().Name, innerError: ex.InnerException?.Message);
        }
    }
}
