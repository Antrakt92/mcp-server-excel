using System.Reflection;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.Service;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>Applies CLI-specific input policy before dispatch.</summary>
internal static class CliInputPolicy
{
    private static readonly HashSet<string> VbaBooleanOptions = typeof(ServiceRegistry.Vba.CliSettings)
        .GetProperties()
        .Where(property => property.PropertyType == typeof(bool) || property.PropertyType == typeof(bool?))
        .Select(property => property.GetCustomAttribute<CommandOptionAttribute>())
        .OfType<CommandOptionAttribute>()
        .SelectMany(option => option.LongNames.Select(name => "--" + name)
            .Concat(option.ShortNames.Select(name => "-" + name)))
        .ToHashSet(StringComparer.OrdinalIgnoreCase);

    internal static bool IsHelpFlag(string argument) => argument is "--help" or "-h" or "-?";

    internal static (string? Args, IReadOnlyList<(string Path, string Option)> ProtectedPaths) BindArtifactPaths(
        string command, string? argsJson, string callerDirectory)
    {
        (string Name, string Option)[] operands = command switch
        {
            "workbook.save-as" or "workbook.save-copy-as" or "workbook.export-fixed-format" or "chart.export-image" =>
                [("targetPath", "--target-path")],
            "sheet.copy-to-file" or "sheet.move-to-file" =>
                [("sourceFile", "--source-file"), ("targetFile", "--target-file")],
            _ => []
        };
        if (argsJson is null || operands.Length == 0)
            return (argsJson, []);

        if (JsonNode.Parse(argsJson) is not JsonObject request)
            throw new ArgumentException("Command arguments must be a JSON object.", nameof(argsJson));
        var protectedPaths = new List<(string Path, string Option)>(operands.Length);
        var changed = false;
        foreach (var (name, option) in operands)
        {
            if (request[name] is not JsonValue operand
                || !operand.TryGetValue<string>(out var path) || string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException($"{name} must be a non-empty path string.", name);
            }
            var absolute = Path.GetFullPath(path, callerDirectory);
            // The daemon's directory can differ from the calling client's directory.
            if (!Path.IsPathFullyQualified(path))
            {
                request[name] = absolute;
                changed = true;
            }
            protectedPaths.Add((absolute, option));
        }

        if (command == "workbook.save-copy-as" && request["packageCheck"] is JsonObject packageCheck
            && packageCheck["baselinePath"] is JsonValue baselineValue
            && baselineValue.TryGetValue<string>(out var baselinePath) && !string.IsNullOrWhiteSpace(baselinePath))
        {
            protectedPaths.Add((Path.GetFullPath(baselinePath, callerDirectory), "--package-check baselinePath"));
        }
        return (changed ? request.ToJsonString(ServiceProtocol.JsonOptions) : argsJson, protectedPaths);
    }

    internal static string? FindStdinOnlyCommand(IReadOnlyList<string> arguments)
    {
        if (arguments.Count < 2) return null;
        var category = _CliCategoryMetadata.ValidActionsByCommand.Keys.FirstOrDefault(candidate =>
            string.Equals(candidate, arguments[0], StringComparison.OrdinalIgnoreCase));
        if (category is null) return null;
        var actions = _CliCategoryMetadata.ValidActionsByCommand[category];
        for (int index = 1; index < arguments.Count; index++)
        {
            var argument = arguments[index];
            if (argument.StartsWith('-'))
            {
                if (argument.IndexOfAny(['=', ':']) < 0 && !IsHelpFlag(argument)
                    && argument is not ("--version" or "-v"))
                {
                    // Spectre allows a bare boolean flag before the action.
                    if (category != ServiceRegistry.Vba.Category || !VbaBooleanOptions.Contains(argument)
                        || (index + 1 < arguments.Count && bool.TryParse(arguments[index + 1], out _)))
                        index++;
                }
                continue;
            }
            var normalized = argument.Trim().ToLowerInvariant();
            var action = actions.FirstOrDefault(candidate => string.Equals(candidate, normalized, StringComparison.OrdinalIgnoreCase));
            if (action is null) return null;
            var command = category + "." + action;
            return _CliCategoryMetadata.RequiresBatchStdin(command) ? command : null;
        }
        return null;
    }
}
