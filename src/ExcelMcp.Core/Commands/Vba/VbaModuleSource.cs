using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

internal static class VbaModuleSource
{
    private static readonly Regex ExportAttribute = new(@"^\s*Attribute\s+VB_[A-Za-z0-9_]+\s*=", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    internal static string Normalize(string source)
    {
        ArgumentNullException.ThrowIfNull(source);
        var lines = NormalizeLines(source).Split('\n');
        var retained = new List<string>();
        var seenCode = false;
        foreach (var line in lines)
        {
            var trimmed = line.TrimStart();
            if (line.Contains('\0') || line.Length > 1023 ||
                Regex.IsMatch(trimmed, @"^VERSION\s+\d", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant) ||
                trimmed.Equals("BEGIN", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Unsupported VBA export header, NUL, or source line longer than 1023 characters.", nameof(source));
            if (trimmed.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase))
            {
                if (seenCode || !ExportAttribute.IsMatch(line))
                    throw new ArgumentException("Only leading module export Attribute VB_* metadata can be omitted safely.", nameof(source));
                continue;
            }
            if (!string.IsNullOrWhiteSpace(line)) seenCode = true;
            retained.Add(line);
        }
        return string.Join("\r\n", retained).TrimEnd('\r', '\n');
    }

    internal static string ComputeHash(string source) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(NormalizeLines(source).TrimEnd('\n'))));

    internal static int CountLines(string source)
    {
        var normalized = NormalizeLines(source).TrimEnd('\n');
        return normalized.Length == 0 ? 0 : normalized.Count(character => character == '\n') + 1;
    }

    internal static void ValidateExpectedHash(string? expectedHash)
    {
        if (expectedHash is not null && (expectedHash.Length != 64 || expectedHash.Any(character => !Uri.IsHexDigit(character))))
            throw new ArgumentException("expectedSourceHash must be a 64-character source fingerprint.", nameof(expectedHash));
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "After a source mutation fails, preserve the primary exception and separately transport any read-back failure in VbaSourceEditState. No failure is treated as success and no rollback runs.")]
    internal static string Replace(string moduleName, string? expectedHash, string source,
        Func<string> read, Action delete, Action<string> insert)
    {
        string normalized = Normalize(source);
        ValidateExpectedHash(expectedHash);
        string current = read();
        if (expectedHash is not null && !CryptographicOperations.FixedTimeEquals(
            Convert.FromHexString(ComputeHash(current)), Convert.FromHexString(expectedHash)))
            throw new OperationFailureException(OperationFailureCategory.Conflict, "The module changed after it was read. Read it again before replacing it.");
        string phase = "deleting";
        try
        {
            delete();
            phase = "inserting";
            if (normalized.Length != 0) insert(normalized);
            phase = "verifying";
            string stored = read();
            if (!CanonicalReadBack(normalized).Equals(CanonicalReadBack(stored), StringComparison.Ordinal))
                throw new InvalidOperationException("The source read back from Excel differs from the supplied replacement.");
            return stored;
        }
        catch (Exception ex)
        {
            string? hash = null;
            int? lineCount = null;
            string? readBackError = null;
            try
            {
                string observed = read();
                hash = ComputeHash(observed);
                lineCount = CountLines(observed);
            }
            catch (Exception readFailure)
            {
                readBackError = readFailure.Message;
            }
            throw new VbaSourceEditException("VBA module replacement failed after mutation started. No rollback was attempted; inspect the observed state before continuing.", ex,
                new VbaSourceEditState { ModuleName = moduleName, Phase = phase, SourceHash = hash, LineCount = lineCount, ReadBackError = readBackError });
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Import owns the created-component state. Report the primary error and independent name/source read-back failures without deleting a created component or reporting success.")]
    internal static void Import(string moduleName, string source, IEnumerable<string> existingNames,
        Action create, Action<string> rename, Action<string> insert, Func<string> actualName, Func<string> read)
    {
        string normalized = Normalize(source);
        if (existingNames.Contains(moduleName, StringComparer.OrdinalIgnoreCase))
            throw new OperationFailureException(OperationFailureCategory.Conflict,
                $"Module '{moduleName}' already exists. Use update or replace-module to modify it.");
        string phase = "creating";
        try
        {
            create();
            phase = "naming";
            rename(moduleName);
            phase = "inserting";
            if (normalized.Length != 0) insert(normalized);
            phase = "verifying";
            if (!CanonicalReadBack(normalized).Equals(CanonicalReadBack(read()), StringComparison.Ordinal))
                throw new InvalidOperationException("Imported source read back from Excel differs from the supplied source.");
        }
        catch (Exception ex)
        {
            string observedName = string.Empty;
            string? hash = null;
            int? lineCount = null;
            var readBackErrors = new List<string>();
            try { observedName = actualName(); }
            catch (Exception nameFailure) { readBackErrors.Add($"Component name: {nameFailure.Message}"); }
            try
            {
                string observed = read();
                hash = ComputeHash(observed);
                lineCount = CountLines(observed);
            }
            catch (Exception readFailure) { readBackErrors.Add($"Source: {readFailure.Message}"); }
            throw new VbaSourceEditException(
                $"VBA import of '{moduleName}' failed after component creation started. No component was removed or source restored; inspect the observed state before continuing.", ex,
                new VbaSourceEditState
                {
                    ModuleName = observedName,
                    Phase = phase,
                    SourceHash = hash,
                    LineCount = lineCount,
                    ReadBackError = readBackErrors.Count == 0 ? null : string.Join(" ", readBackErrors)
                });
        }
    }

    // VBE recases identifiers. Literal text and comments must still read back exactly.
    private static string CanonicalReadBack(string source)
    {
        var result = new StringBuilder();
        foreach (string line in NormalizeLines(source).TrimEnd('\n').Split('\n'))
        {
            bool quoted = false;
            bool statementStart = true;
            for (int index = 0; index < line.Length; index++)
            {
                char character = line[index];
                if (!quoted && (character == '\'' || statementStart && index + 3 <= line.Length &&
                    line.AsSpan(index, 3).Equals("Rem", StringComparison.OrdinalIgnoreCase) &&
                    (index + 3 == line.Length || char.IsWhiteSpace(line[index + 3]))))
                {
                    if (character != '\'')
                    {
                        result.Append("REM");
                        index += 3;
                    }
                    result.Append(line.AsSpan(index));
                    break;
                }
                if (!quoted && char.IsLetter(character))
                {
                    int end = index + 1;
                    while (end < line.Length && (char.IsLetterOrDigit(line[end]) || line[end] == '_')) end++;
                    string token = line[index..end];
                    result.Append(token.ToUpperInvariant());
                    statementStart = token.Equals("Then", StringComparison.OrdinalIgnoreCase)
                        || token.Equals("Else", StringComparison.OrdinalIgnoreCase);
                    index = end - 1;
                    continue;
                }
                if (!quoted && statementStart && char.IsDigit(character))
                {
                    int end = index + 1;
                    while (end < line.Length && char.IsDigit(line[end])) end++;
                    result.Append(line.AsSpan(index, end - index));
                    index = end - 1;
                    continue;
                }
                result.Append(quoted ? character : char.ToUpperInvariant(character));
                if (character == '"') quoted = !quoted;
                if (!quoted && !char.IsWhiteSpace(character)) statementStart = character == ':';
            }
            result.Append('\n');
        }
        return result.ToString();
    }

    private static string NormalizeLines(string source) =>
        source.Replace("\r\r\n", "\n", StringComparison.Ordinal)
            .Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');
}
