using System.Reflection;
using Sbroenne.ExcelMcp.Generated;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>Enforces source-declared input policy before Spectre can convert or print argument values.</summary>
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
                if (argument.IndexOfAny(['=', ':']) < 0 && argument is not ("--help" or "-h" or "--version" or "-v"))
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
