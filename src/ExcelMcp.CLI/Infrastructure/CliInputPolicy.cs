using Sbroenne.ExcelMcp.Generated;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>Enforces source-declared input policy before Spectre can convert or print argument values.</summary>
internal static class CliInputPolicy
{
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
                    index++;
                continue;
            }
            var action = actions.FirstOrDefault(candidate => string.Equals(candidate, argument.Trim(), StringComparison.OrdinalIgnoreCase));
            if (action is null) return null;
            var command = category + "." + action;
            return _CliCategoryMetadata.RequiresBatchStdin(command) ? command : null;
        }
        return null;
    }
}
