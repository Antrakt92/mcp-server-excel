using System.Text;
using System.Text.RegularExpressions;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

internal sealed record VbaLabelSource(string ModuleName, bool IsStandardModule, string Source);

internal static class VbaLabelScanner
{
    internal const int MaximumProjectCharacters = 2_000_000;
    private const int MaximumWarnings = 100;
    private const RegexOptions Keywords = RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.NonBacktracking;
    private const string Identifier = @"[\p{L}_][\p{L}\p{Nd}_]*";
    private const string Target = "(?:" + Identifier + @"|[0-9]+)";
    private static readonly Regex ZeroArgumentSub = new(
        @"^\s*(?:(?:Public|Private|Friend|Global|Static)\s+)*Sub\s+(" + Identifier + @")\s*(?:\(\s*\))?\s*(?::|$)", Keywords);
    private static readonly Regex LocalAccess = new(@"^\s*(?:Static\s+)*(?:Private|Friend)\b", Keywords);
    private static readonly Regex ProcedureEnd = new(@"(?:^|:)\s*End\s+(?:Sub|Function|Property)\s*(?::|$)", Keywords);
    private static readonly Regex ConditionalStart = new(@"^\s*#If\b", Keywords);
    private static readonly Regex ConditionalEnd = new(@"^\s*#End\s+If\b", Keywords);
    private static readonly Regex Label = new(@"^\s*(" + Identifier + @")\s*:", Keywords);
    private static readonly Regex Jump = new(
        @"\b(?:GoTo|GoSub|Resume)\s+(" + Target + @"(?:\s*,\s*" + Target + @")*)", Keywords);
    private static readonly Regex QuotedText = new(
        "\"(?:[^\"\\r\\n]|\"\")*\"", RegexOptions.NonBacktracking | RegexOptions.CultureInvariant);

    internal static VbaLabelDiagnostics Analyze(
        IReadOnlyList<VbaLabelSource> modules, CancellationToken cancellationToken = default)
    {
        var globalSubs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var parsedModules = new List<ModuleLabels>(modules.Count);
        foreach (VbaLabelSource module in modules)
        {
            cancellationToken.ThrowIfCancellationRequested();
            parsedModules.Add(ReadModule(module, globalSubs, cancellationToken));
        }

        var result = new VbaLabelDiagnostics();
        foreach (ModuleLabels module in parsedModules)
        {
            foreach (ProcedureLabels procedure in module.Procedures)
            {
                foreach ((int line, string label) in procedure.Labels)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (procedure.JumpTargets.Contains(label)
                        || (!module.LocalSubs.Contains(label) && !globalSubs.Contains(label)))
                    {
                        continue;
                    }

                    if (result.Warnings.Count == MaximumWarnings)
                    {
                        result.HasMore = true;
                        return result;
                    }

                    result.Warnings.Add(new VbaLabelWarning
                    {
                        ModuleName = module.Name,
                        ProcedureName = procedure.Name,
                        Line = line,
                        Label = label
                    });
                }
            }
        }

        return result;
    }

    private static ModuleLabels ReadModule(
        VbaLabelSource module, HashSet<string> globalSubs, CancellationToken cancellationToken)
    {
        var result = new ModuleLabels(module.ModuleName);
        ProcedureLabels? procedure = null;
        int conditionalDepth = 0;
        foreach ((int line, string code) in LogicalLines(module.Source, cancellationToken))
        {
            if (ConditionalStart.IsMatch(code))
            {
                conditionalDepth++;
                continue;
            }
            if (ConditionalEnd.IsMatch(code))
            {
                conditionalDepth = Math.Max(0, conditionalDepth - 1);
                continue;
            }
            if (code.TrimStart().StartsWith('#'))
            {
                continue;
            }

            if (VbaProcedureSource.TryParseProcedureHeader(code) is { } header)
            {
                procedure = new ProcedureLabels(header.Name);
                result.Procedures.Add(procedure);

                // Only unconditional declarations and labels qualify. Jumps in any branch can exempt a label.
                if (conditionalDepth == 0 && ZeroArgumentSub.Match(code) is { Success: true } declaration)
                {
                    string name = declaration.Groups[1].Value;
                    result.LocalSubs.Add(name);
                    if (module.IsStandardModule && !LocalAccess.IsMatch(code))
                    {
                        globalSubs.Add(name);
                    }
                }
            }

            if (procedure is null)
            {
                continue;
            }

            if (conditionalDepth == 0 && Label.Match(code) is { Success: true } label
                && (label.Length == code.Length || code[label.Length] != '='))
            {
                procedure.Labels.Add((line, label.Groups[1].Value));
            }

            string withoutStrings = QuotedText.Replace(code, "\"\"");
            foreach (Match jump in Jump.Matches(withoutStrings))
            {
                procedure.JumpTargets.UnionWith(jump.Groups[1].Value.Split(',', StringSplitOptions.TrimEntries));
            }

            if (ProcedureEnd.IsMatch(withoutStrings))
            {
                procedure = null;
            }
        }

        return result;
    }

    private static IEnumerable<(int Line, string Code)> LogicalLines(string source, CancellationToken cancellationToken)
    {
        using var reader = new StringReader(source);
        var pending = new StringBuilder();
        int startLine = 1;
        for (int line = 1; reader.ReadLine() is { } physical; line++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string code = physical[..VbaModuleSource.CommentStart(physical)].TrimEnd();
            if (pending.Length == 0)
            {
                startLine = line;
            }

            if (code.Length > 1 && code[^1] == '_' && char.IsWhiteSpace(code[^2]))
            {
                pending.Append(code.AsSpan(0, code.Length - 1));
                continue;
            }

            if (pending.Length == 0)
            {
                yield return (line, code);
                continue;
            }

            pending.Append(code);
            yield return (startLine, pending.ToString());
            pending.Clear();
        }
    }

    private sealed record ModuleLabels(string Name)
    {
        public HashSet<string> LocalSubs { get; } = new(StringComparer.OrdinalIgnoreCase);
        public List<ProcedureLabels> Procedures { get; } = [];
    }

    private sealed record ProcedureLabels(string Name)
    {
        public HashSet<string> JumpTargets { get; } = new(StringComparer.OrdinalIgnoreCase);
        public List<(int Line, string Label)> Labels { get; } = [];
    }
}
