using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.CLI.Telemetry;
using Sbroenne.ExcelMcp.Service;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Base class for CLI commands that send requests to the service.
/// Handles common validation and execution patterns.
/// </summary>
internal abstract class ServiceCommandBase<TSettings> : AsyncCommand<TSettings>
    where TSettings : CommandSettings
{
    /// <summary>
    /// Gets the session ID from settings.
    /// </summary>
    protected abstract string? GetSessionId(TSettings settings);

    /// <summary>
    /// Gets the action from settings.
    /// </summary>
    protected abstract string? GetAction(TSettings settings);

    /// <summary>
    /// Gets the valid actions for this command.
    /// </summary>
    protected abstract IReadOnlyList<string> ValidActions { get; }

    /// <summary>Gets the service registry category for the generated command.</summary>
    protected abstract string CommandCategory { get; }

    /// <summary>
    /// Routes the action to a service command and args.
    /// </summary>
    protected abstract (string command, object? args) Route(TSettings settings, string action);

    /// <summary>
    /// Whether this command requires a session ID. Default is true.
    /// Override to return false for commands that don't need a session.
    /// </summary>
    protected virtual bool RequiresSession => true;

    /// <summary>
    /// Validates settings and executes the command.
    /// Returns early with error code if validation fails.
    /// </summary>
    public sealed override async Task<int> ExecuteAsync(CommandContext context, TSettings settings, CancellationToken cancellationToken)
    {
        // Action validation
        var rawAction = GetAction(settings);
        if (string.IsNullOrWhiteSpace(rawAction))
        {
            return CliErrorOutput.WriteError("Action is required.");
        }

        var action = rawAction.Trim().ToLowerInvariant();
        if (!ValidActions.Contains(action, StringComparer.OrdinalIgnoreCase))
        {
            var validList = string.Join(", ", ValidActions);
            return CliErrorOutput.WriteError($"Invalid action '{action}'. Valid actions: {validList}");
        }

        var sessionId = GetSessionId(settings);
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            if (RequiresSession && !ExcelMcpService.IsSessionlessCommand(CommandCategory + "." + action))
                return CliErrorOutput.WriteError("Session ID is required. Use --session <id>");
            sessionId = null;
        }

        // Check for --output file path (generated on all CliSettings).
        var outputPath = settings.GetType().GetProperty("OutputPath")?.GetValue(settings) as string;

        string command;
        string? requestArgs;
        try
        {
            (command, var args) = Route(settings, action);
            requestArgs = SerializeRequestArguments(command, args, outputPath);
        }
        catch (ArgumentException ex)
        {
            // Parameter validation failed (e.g., required param missing)
            // Return clean JSON error with exit code 1 instead of unhandled crash
            return CliErrorOutput.WriteException(ex, "InvalidInput");
        }
        catch (PathTooLongException ex)
        {
            return CliErrorOutput.WriteException(ex, "InvalidInput");
        }

        // Connect to CLI daemon service (auto-starts if not running)
        using var client = await CliCommandRuntime.Current.ClientFactory.ConnectAsync(cancellationToken);
        var request = new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = requestArgs
        };
        var response = await CliTelemetry.TrackCommandAsync(
            request,
            () => client.SendAsync(request, cancellationToken));

        // Output result
        if (response.Success)
        {
            var outcome = CliResultOutput.Read(response, command);
            var originalData = outcome.Data as System.Text.Json.Nodes.JsonObject;
            var durationMs = CliCommandRuntime.Current.OperationTimer.ElapsedMilliseconds;
            if (!outcome.Success)
                return CliResultOutput.Write(outcome, command);

            if (!string.IsNullOrEmpty(outputPath))
            {
                return WriteOutputToFile(outcome, originalData, command, durationMs, outputPath);
            }

            return CliResultOutput.Write(outcome, command);
        }
        else
        {
            return CliResultOutput.Write(response, command);
        }
    }

    private static string? SerializeRequestArguments(string command, object? args, string? outputPath)
    {
        if (args is null)
            return null;
        var writesArtifact = command is "workbook.save-as" or "workbook.save-copy-as" or "workbook.export-fixed-format" or "chart.export-image";
        var transfersSheet = command is "sheet.copy-to-file" or "sheet.move-to-file";
        if (string.IsNullOrEmpty(outputPath) || (!writesArtifact && !transfersSheet))
            return JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions);

        var request = JsonSerializer.SerializeToNode(args, ServiceProtocol.JsonOptions)!.AsObject();
        var callerDirectory = Environment.CurrentDirectory;
        var output = Path.GetFullPath(outputPath, callerDirectory);
        if (transfersSheet)
        {
            ValidateOutputPath(request, "sourceFile", "--source-file", callerDirectory, output);
            ValidateOutputPath(request, "targetFile", "--target-file", callerDirectory, output);
        }
        else
        {
            ValidateOutputPath(request, "targetPath", "--target-path", callerDirectory, output);
        }
        return request.ToJsonString(ServiceProtocol.JsonOptions);
    }

    private static void ValidateOutputPath(JsonObject request, string parameter, string option, string callerDirectory, string output)
    {
        var path = request[parameter]!.GetValue<string>();
        var absolute = Path.GetFullPath(path, callerDirectory);
        if (string.Equals(output, absolute, StringComparison.OrdinalIgnoreCase) || ReferToSameFile(output, absolute))
            throw new ArgumentException($"--output writes the command response and must use a different file from {option}.");

        // The daemon's working directory can differ from the caller's.
        if (!Path.IsPathFullyQualified(path))
            request[parameter] = absolute;
    }

    private static bool ReferToSameFile(string outputPath, string targetPath)
    {
        try
        {
            var output = WorkbookFileIdentity.Read(outputPath);
            var target = WorkbookFileIdentity.Read(targetPath);
            return string.Equals(output.Path, target.Path, StringComparison.OrdinalIgnoreCase)
                || output.FileKey is not null && string.Equals(output.FileKey, target.FileKey, StringComparison.OrdinalIgnoreCase);
        }
        catch (Win32Exception)
        {
            // The application will report inaccessible paths when it attempts the write.
            return false;
        }
        catch (IOException)
        {
            return false;
        }
    }

    /// <summary>
    /// Writes the result to a file. For image results containing base64 data,
    /// decodes and writes the binary image. Otherwise writes the JSON text.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "CLI output-file boundary reports the original write/encoding failure as success=false JSON and returns exit code 1.")]
    private static int WriteOutputToFile(ResultJson outcome, System.Text.Json.Nodes.JsonObject? originalData,
        string command, long durationMs, string outputPath)
    {
        try
        {
            Dictionary<string, object?> metadata;
            var base64Data = originalData?["imageBase64"]?.GetValue<string>();
            if (base64Data != null)
            {
                var imageBytes = Convert.FromBase64String(base64Data);
                File.WriteAllBytes(outputPath, imageBytes);
                // Write metadata (without the large base64 payload) to stdout
                using var doc = JsonDocument.Parse(originalData!.ToJsonString());
                metadata = new Dictionary<string, object?>
                {
                    ["success"] = true,
                    ["outputPath"] = outputPath,
                    ["sizeBytes"] = imageBytes.Length
                };
                if (doc.RootElement.TryGetProperty("width", out var w)) metadata["width"] = w.GetInt32();
                if (doc.RootElement.TryGetProperty("height", out var h)) metadata["height"] = h.GetInt32();
                if (doc.RootElement.TryGetProperty("mimeType", out var m)) metadata["mimeType"] = m.GetString();
                if (doc.RootElement.TryGetProperty("sheetName", out var s)) metadata["sheetName"] = s.GetString();
                if (doc.RootElement.TryGetProperty("rangeAddress", out var r)) metadata["rangeAddress"] = r.GetString();
                if (doc.RootElement.TryGetProperty("message", out var message)) metadata["message"] = message.GetString();
            }
            else
            {
                File.WriteAllText(outputPath, outcome.ToCliJson(command, durationMs));
                metadata = new Dictionary<string, object?>
                {
                    ["success"] = true,
                    ["outputPath"] = outputPath
                };
            }
            metadata["warnings"] = outcome.Warnings;
            metadata["partialState"] = outcome.PartialState;
            return CliResultOutput.Write(ResultJson.Read(
                JsonSerializer.Serialize(metadata, ServiceProtocol.JsonOptions), command), command, durationMs);
        }
        catch (Exception ex)
        {
            return CliResultOutput.Write(outcome.WithFailure(command, ex, "OutputWriteFailed",
                $"The operation completed, but writing its output failed: {ex.Message}",
                partialState: JsonSerializer.Serialize(new
                {
                    operationCompleted = true,
                    outputWriteCompleted = false,
                    operationPartialState = outcome.PartialState
                }, ServiceProtocol.JsonOptions)), command, durationMs);
        }
    }
}
