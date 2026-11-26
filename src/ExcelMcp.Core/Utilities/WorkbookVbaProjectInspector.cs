using System.Security.Cryptography;
using System.Text;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Utilities;

internal static class WorkbookVbaProjectInspector
{
    internal const int MaxProjectBytes = 8 * 1024 * 1024;
    private const int MaxDecodedModuleBytes = 4 * 1024 * 1024;
    private const int MaxDecodedProjectBytes = 16 * 1024 * 1024;
    private const int MaxReportedChanges = 512;
    private const string Other = "other-stream";

    private sealed record Item(string Category, string Name);
    private sealed record Project(Dictionary<Item, string> Hashes, string? References, int ReferenceCount, string? Protection);
    private sealed record Properties(string Metadata, string? Protection, HashSet<string> Forms);

    internal static WorkbookVbaProjectComparison Compare(byte[]? before, string? beforePart, byte[]? after, string? afterPart,
        CancellationToken cancellationToken)
    {
        try
        {
            var left = before is null ? new Project([], null, 0, null) : Inspect(before, cancellationToken);
            var right = after is null ? new Project([], null, 0, null) : Inspect(after, cancellationToken);
            var changes = new List<WorkbookVbaProjectItemChange>();
            int otherChanges = 0;
            foreach (var key in left.Hashes.Keys.Union(right.Hashes.Keys)
                .OrderBy(item => item.Category, StringComparer.Ordinal).ThenBy(item => item.Name, StringComparer.Ordinal))
            {
                cancellationToken.ThrowIfCancellationRequested();
                left.Hashes.TryGetValue(key, out string? previous);
                right.Hashes.TryGetValue(key, out string? current);
                if (previous == current)
                    continue;
                if (key.Category == Other)
                {
                    otherChanges++;
                    continue;
                }
                if (changes.Count == MaxReportedChanges)
                    return Unavailable(beforePart, afterPart, "vba-project-report-limit");
                changes.Add(new(key.Category, key.Name, previous is null ? "added" : current is null ? "removed" : "changed"));
            }

            return new()
            {
                Status = "compared",
                BeforePart = beforePart,
                AfterPart = afterPart,
                References = CompareMetadata(left.References, right.References),
                BeforeReferenceCount = left.ReferenceCount,
                AfterReferenceCount = right.ReferenceCount,
                ProtectionMetadata = (before is not null && left.Protection is null) ||
                    (after is not null && right.Protection is null) ? "unknown" : CompareMetadata(left.Protection, right.Protection),
                OtherChangedStreams = otherChanges,
                Changes = changes
            };
        }
        catch (InvalidDataException)
        {
            // Preserve the completed outer ZIP comparison and omit potentially private parse details.
            return Unavailable(beforePart, afterPart, "vba-project-format-unavailable");
        }
        catch (EndOfStreamException)
        {
            return Unavailable(beforePart, afterPart, "vba-project-truncated");
        }
        catch (DecoderFallbackException)
        {
            return Unavailable(beforePart, afterPart, "vba-project-encoding-unavailable");
        }
    }

    private static WorkbookVbaProjectComparison Unavailable(string? beforePart, string? afterPart, string reason) =>
        new() { Status = "unavailable", BeforePart = beforePart, AfterPart = afterPart, Reason = reason };

    private static string CompareMetadata(string? before, string? after) =>
        before == after ? before is null ? "unknown" : "unchanged" : "changed";

    private static Project Inspect(byte[] bytes, CancellationToken cancellationToken)
    {
        using var compound = OleCompoundFileReader.OpenProject(bytes, cancellationToken);
        byte[] Read(string stream)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return compound.TryReadStream(stream, out var value) ? value :
                throw new InvalidDataException("A required VBA project OLE stream is missing.");
        }

        var directory = VbaProjectDirectory.Read(Read("VBA/dir"), cancellationToken);
        var properties = ReadProperties(Read("PROJECT"), directory.Encoding, cancellationToken);
        byte[] projectInfo = Read("VBA/_VBA_PROJECT");
        if (projectInfo.Length < 7)
            throw new InvalidDataException("The VBA project metadata is incomplete.");

        var hashes = new Dictionary<Item, string>
        {
            [new("project-metadata", "PROJECT")] = properties.Metadata,
            [new("project-metadata", "VBA/dir")] = directory.Metadata,
            [new("project-metadata", "VBA/_VBA_PROJECT")] = VbaProjectDirectory.Hash(projectInfo.AsSpan(0, 7)),
            [new("performance-cache", "VBA/_VBA_PROJECT")] = VbaProjectDirectory.Hash(projectInfo.AsSpan(7))
        };
        var consumed = new HashSet<string>(["PROJECT", "VBA/dir", "VBA/_VBA_PROJECT"], StringComparer.OrdinalIgnoreCase);
        var designerStorage = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        int remaining = MaxDecodedProjectBytes;
        foreach (var module in directory.Modules)
        {
            string path = "VBA/" + module.StreamName;
            byte[] stream = Read(path);
            if (module.SourceOffset > stream.Length || !consumed.Add(path))
                throw new InvalidDataException("The VBA module source offset or stream identity is invalid.");
            byte[] source = VbaProjectDirectory.Decompress(stream.AsSpan(module.SourceOffset),
                Math.Min(MaxDecodedModuleBytes, remaining), cancellationToken);
            remaining -= source.Length;
            if (properties.Forms.Contains(module.Name))
                designerStorage.Add(module.StreamName, module.Name);
            hashes.Add(new("module-source", module.Name),
                VbaProjectDirectory.Hash(Encoding.UTF8.GetBytes(directory.Encoding.GetString(source))));
            hashes.Add(new("module-metadata", module.Name), module.Metadata);
            hashes.Add(new("module-cache", module.Name), VbaProjectDirectory.Hash(stream.AsSpan(0, module.SourceOffset)));
        }

        var detectedForms = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string path in compound.StreamPaths)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (consumed.Contains(path))
                continue;
            int separator = path.IndexOf('/');
            if (separator > 0 && designerStorage.TryGetValue(path[..separator], out var formName))
            {
                // The storage uses the module stream name; report the logical form name.
                string part = path[(separator + 1)..];
                string displayName = formName + "/" + string.Concat(part.Select(character =>
                    char.IsControl(character) || character == '%' ? $"%{(int)character:X4}" : character.ToString()));
                hashes.Add(new("designer-stream", displayName), VbaProjectDirectory.Hash(Read(path)));
                if (!part.Contains('/') && part.StartsWith("\u0003VBFrame", StringComparison.OrdinalIgnoreCase))
                    detectedForms.Add(formName);
                continue;
            }
            string category = path is "PROJECTwm" or "PROJECTlk" ? "project-metadata" :
                IsPerformanceCache(path) ? "performance-cache" : Other;
            hashes.Add(new(category, path), VbaProjectDirectory.Hash(Read(path)));
        }
        if (!detectedForms.SetEquals(properties.Forms))
            throw new InvalidDataException("The VBA project designer storage could not be matched to its module.");
        return new(hashes, directory.References, directory.ReferenceCount, properties.Protection);
    }

    private static bool IsPerformanceCache(string path)
    {
        const string prefix = "VBA/__SRP_";
        return path.Length > prefix.Length && path.Length <= prefix.Length + 25 &&
            path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) && path.Skip(prefix.Length).All(char.IsAsciiHexDigit);
    }

    private static Properties ReadProperties(byte[] bytes, Encoding encoding, CancellationToken cancellationToken)
    {
        const int maxPropertiesBytes = 256 * 1024;
        if (bytes.Length > maxPropertiesBytes)
            throw new InvalidDataException("The VBA project properties exceed the inspection limit.");

        string text = encoding.GetString(bytes);
        using var reader = new StringReader(text);
        using var metadata = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var forms = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var protection = new Dictionary<string, string>(StringComparer.Ordinal);
        bool inProperties = true;
        while (reader.ReadLine() is { } line)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (line.StartsWith('['))
                inProperties = false;
            int equals = line.IndexOf('=');
            string key = inProperties && equals > 0 ? line[..equals] : "";
            if (key is "CMG" or "DPB" or "GC")
            {
                if (!protection.TryAdd(key, line))
                    throw new InvalidDataException("The VBA project contains duplicate protection metadata.");
                continue;
            }
            if (key == "BaseClass")
            {
                string form = line[(equals + 1)..];
                if (!VbaProjectDirectory.IsSafeName(form) ||
                    form.Equals("VBA", StringComparison.OrdinalIgnoreCase) || !forms.Add(form))
                    throw new InvalidDataException("The VBA project contains an unsupported designer identity.");
            }
            metadata.AppendData(Encoding.UTF8.GetBytes(line + "\n"));
        }

        string? protectionHash = protection.Count == 3
            ? VbaProjectDirectory.Hash(Encoding.UTF8.GetBytes(string.Join("\n", protection.OrderBy(pair => pair.Key, StringComparer.Ordinal).Select(pair => pair.Value))))
            : null;
        return new(Convert.ToHexString(metadata.GetHashAndReset()), protectionHash, forms);
    }
}
