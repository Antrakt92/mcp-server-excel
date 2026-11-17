using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Utilities;

/// <summary>Reads selected, bounded metadata from already verified saved workbook packages.</summary>
internal static class WorkbookPackageStructure
{
    private const int MaxRecords = 4096;
    private const int MaxMetadataBytes = 2 * 1024 * 1024;
    private const int MaxMetadataTotalBytes = 16 * 1024 * 1024;
    private const int MaxTextLength = 2048;
    private static readonly XNamespace PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace TransitionalSpreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace StrictSpreadsheet = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    private const string TransitionalRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private const string StrictRelationships = "http://purl.oclc.org/ooxml/officeDocument/relationships/";

    internal static WorkbookPackageStructureReport Compare(
        IReadOnlyDictionary<string, byte[]> beforeMetadata, IReadOnlyCollection<string> beforePartNames,
        IReadOnlyDictionary<string, string?> beforeCodeNames,
        IReadOnlyDictionary<string, byte[]> afterMetadata, IReadOnlyCollection<string> afterPartNames,
        IReadOnlyDictionary<string, string?> afterCodeNames, CancellationToken cancellationToken)
    {
        var before = ReadSnapshot(beforeMetadata, beforePartNames, beforeCodeNames, cancellationToken);
        var after = ReadSnapshot(afterMetadata, afterPartNames, afterCodeNames, cancellationToken);
        return new WorkbookPackageStructureReport
        {
            Before = before,
            After = after,
            ParentStyleMappings = MapParentStyles(before.ParentStyles, after.ParentStyles, cancellationToken)
        };
    }

    private static WorkbookPackageStructureSnapshot ReadSnapshot(
        IReadOnlyDictionary<string, byte[]> metadata, IReadOnlyCollection<string> partNames,
        IReadOnlyDictionary<string, string?> codeNames, CancellationToken cancellationToken)
    {
        long total = 0;
        foreach (var bytes in metadata.Values)
        {
            cancellationToken.ThrowIfCancellationRequested();
            total += bytes.Length;
            if (bytes.Length > MaxMetadataBytes || total > MaxMetadataTotalBytes)
                throw new InvalidDataException("The package structure metadata exceeds its size limit.");
        }

        var parts = partNames.ToHashSet(StringComparer.Ordinal);
        var relationships = ReadRelationships(metadata, parts, cancellationToken);
        var roots = relationships.Where(item => item.OwnerPart.Length == 0 && !item.External &&
            IsRelationship(item.Type, "officeDocument")).ToArray();
        if (roots.Length != 1 || roots[0].TargetExists != true)
            throw new InvalidDataException("The structure report requires one existing internal root workbook relationship.");
        string workbookPart = roots[0].TargetPart!;
        var workbook = ReadXml(metadata, workbookPart, "workbook", cancellationToken);
        var sheetNamespace = workbook.Name.Namespace;
        RequireSpreadsheetNamespace(sheetNamespace);
        var workbookRelationships = relationships.Where(item => item.OwnerPart == workbookPart)
            .ToDictionary(item => item.Id, StringComparer.Ordinal);
        var sheets = ReadSheets(workbook, workbookRelationships, codeNames, cancellationToken);
        var names = ReadDefinedNames(workbook, sheets, cancellationToken);
        var styles = ReadParentStyles(metadata, workbookRelationships.Values, cancellationToken);
        var reachable = ReachableParts(relationships, cancellationToken);
        var unreachable = parts.Where(name => name != WorkbookPackageComparison.ContentTypesEntry &&
                !IsRelationshipsPart(name) && !reachable.Contains(name))
            .Order(StringComparer.Ordinal).ToArray();
        RequireRecordCount(unreachable.Length);

        return new WorkbookPackageStructureSnapshot
        {
            WorkbookPart = workbookPart,
            Sheets = sheets,
            Relationships = relationships,
            DefinedNames = names,
            ParentStyles = styles,
            UnreachableParts = unreachable
        };
    }

    private static List<WorkbookPackageRelationship> ReadRelationships(
        IReadOnlyDictionary<string, byte[]> metadata, HashSet<string> parts, CancellationToken cancellationToken)
    {
        var result = new List<WorkbookPackageRelationship>();
        foreach (string part in metadata.Keys.Where(IsRelationshipsPart).Order(StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            string owner = RelationshipOwner(part);
            var root = ReadXml(metadata, part, "Relationships", cancellationToken);
            if (root.Name.Namespace != PackageRelationships)
                throw new InvalidDataException("The package relationship metadata has an unsupported namespace.");
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (var relationship in root.Elements())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (relationship.Name != PackageRelationships + "Relationship")
                    throw new InvalidDataException("The package relationship metadata contains an unexpected element.");
                string id = RequiredText(relationship, "Id");
                string type = RequiredText(relationship, "Type");
                if (!ids.Add(id))
                    throw new InvalidDataException("The package relationship metadata contains duplicate IDs.");
                string? mode = (string?)relationship.Attribute("TargetMode");
                if (mode is not (null or "Internal" or "External"))
                    throw new InvalidDataException("The package relationship metadata contains an invalid target mode.");
                bool external = mode == "External";
                string target = RequiredText(relationship, "Target");
                string? resolved = external ? null : ResolveTarget(owner, target);
                result.Add(new WorkbookPackageRelationship
                {
                    OwnerPart = owner,
                    OwnerExists = owner.Length == 0 || parts.Contains(owner),
                    Id = id,
                    Type = type,
                    External = external,
                    TargetPart = resolved,
                    TargetExists = external ? null : parts.Contains(resolved!)
                });
                RequireRecordCount(result.Count);
            }
        }
        return result;
    }

    private static List<WorkbookPackageSheet> ReadSheets(XElement workbook,
        Dictionary<string, WorkbookPackageRelationship> relationships,
        IReadOnlyDictionary<string, string?> codeNames, CancellationToken cancellationToken)
    {
        XNamespace ns = workbook.Name.Namespace;
        if (workbook.Elements(ns + "sheets").Take(2).Count() != 1)
            throw new InvalidDataException("The workbook must have one saved sheets collection.");
        var result = new List<WorkbookPackageSheet>();
        var ids = new HashSet<uint>();
        var sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var targets = new HashSet<string>(StringComparer.Ordinal);
        foreach (var sheet in workbook.Elements(ns + "sheets").Elements(ns + "sheet"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            string name = RequiredText(sheet, "name");
            uint sheetId = ReadUnsigned(RequiredText(sheet, "sheetId"));
            string? relationshipId = (string?)sheet.Attribute(XName.Get("id", TransitionalRelationships.TrimEnd('/'))) ??
                (string?)sheet.Attribute(XName.Get("id", StrictRelationships.TrimEnd('/')));
            if (sheetId == 0 || !ids.Add(sheetId) || !sheetNames.Add(name) || string.IsNullOrEmpty(relationshipId) ||
                !relationships.TryGetValue(relationshipId, out var relationship) || relationship.External ||
                relationship.TargetExists != true ||
                !(IsRelationship(relationship.Type, "worksheet") || IsRelationship(relationship.Type, "chartsheet") ||
                  IsRelationship(relationship.Type, "dialogsheet") ||
                  relationship.Type is "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" or
                      "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet") ||
                !targets.Add(relationship.TargetPart!))
                throw new InvalidDataException("The workbook contains an invalid or ambiguous sheet identity or relationship.");
            if (!codeNames.TryGetValue(relationship.TargetPart!, out string? codeName))
                throw new InvalidDataException("The sheet structural metadata was not captured from its declared content type.");
            if (codeName?.Length > MaxTextLength)
                throw new InvalidDataException("A saved sheet codeName exceeds the structure text limit.");
            result.Add(new WorkbookPackageSheet
            {
                Position = result.Count,
                Name = name,
                SheetId = sheetId,
                RelationshipId = relationshipId,
                PartName = relationship.TargetPart!,
                SheetType = relationship.Type[(relationship.Type.LastIndexOf('/') + 1)..],
                CodeName = codeName
            });
            RequireRecordCount(result.Count);
        }
        return result;
    }

    private static List<WorkbookPackageDefinedName> ReadDefinedNames(XElement workbook,
        List<WorkbookPackageSheet> sheets, CancellationToken cancellationToken)
    {
        XNamespace ns = workbook.Name.Namespace;
        if (workbook.Elements(ns + "definedNames").Take(2).Count() > 1)
            throw new InvalidDataException("The workbook contains duplicate defined-name collections.");
        var result = new List<WorkbookPackageDefinedName>();
        var identities = new HashSet<(string Name, int? Scope)>();
        foreach (var name in workbook.Elements(ns + "definedNames").Elements(ns + "definedName"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            string identifier = RequiredText(name, "name");
            string? scope = (string?)name.Attribute("localSheetId");
            int? localSheetId = scope is null ? null : ReadIndex(scope);
            if (localSheetId >= sheets.Count || !identities.Add((identifier.ToUpperInvariant(), localSheetId)))
                throw new InvalidDataException("The workbook contains an invalid or ambiguous defined-name scope.");
            var sheet = localSheetId.HasValue ? sheets[localSheetId.Value] : null;
            result.Add(new WorkbookPackageDefinedName
            {
                Name = identifier,
                LocalSheetId = localSheetId,
                SheetId = sheet?.SheetId,
                SheetPart = sheet?.PartName
            });
            RequireRecordCount(result.Count);
        }
        return result;
    }

    private static List<WorkbookPackageParentStyle> ReadParentStyles(IReadOnlyDictionary<string, byte[]> metadata,
        IEnumerable<WorkbookPackageRelationship> relationships, CancellationToken cancellationToken)
    {
        var styles = relationships.Where(item => IsRelationship(item.Type, "styles")).ToArray();
        if (styles.Length == 0)
            return [];
        if (styles.Length != 1 || styles[0].External || styles[0].TargetExists != true)
            throw new InvalidDataException("The workbook styles relationship is invalid or ambiguous.");
        var root = ReadXml(metadata, styles[0].TargetPart!, "styleSheet", cancellationToken);
        XNamespace ns = root.Name.Namespace;
        RequireSpreadsheetNamespace(ns);
        var parents = ReadStyleRecords(root, "cellStyleXfs", "xf");
        var fonts = ReadStyleRecords(root, "fonts", "font");
        var fills = ReadStyleRecords(root, "fills", "fill");
        var borders = ReadStyleRecords(root, "borders", "border");
        var formats = new Dictionary<int, XElement>();
        foreach (var format in ReadStyleRecords(root, "numFmts", "numFmt"))
        {
            int id = ReadIndex(RequiredText(format, "numFmtId"));
            if (!formats.TryAdd(id, format))
                throw new InvalidDataException("The styles part contains duplicate number-format IDs.");
        }

        var result = new List<WorkbookPackageParentStyle>();
        var hashes = new Dictionary<int, string>();
        var dependencyHashes = new Dictionary<(string Kind, int Index), string>();
        foreach (var style in ReadStyleRecords(root, "cellStyles", "cellStyle"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            string name = RequiredText(style, "name");
            int xfId = ReadIndex(RequiredText(style, "xfId"));
            if (xfId >= parents.Count)
                throw new InvalidDataException("A named cell style refers to a missing parent definition.");
            if (!hashes.TryGetValue(xfId, out string? hash))
            {
                var xf = parents[xfId];
                using var buffer = new CanonicalBuffer();
                using var writer = new BinaryWriter(buffer, Encoding.UTF8, leaveOpen: true);
                WriteCanonical(writer, xf, ["fontId", "fillId", "borderId", "numFmtId"], cancellationToken);
                WriteStyleReference(writer, xf, "fontId", fonts, dependencyHashes, cancellationToken);
                WriteStyleReference(writer, xf, "fillId", fills, dependencyHashes, cancellationToken);
                WriteStyleReference(writer, xf, "borderId", borders, dependencyHashes, cancellationToken);
                string? formatId = (string?)xf.Attribute("numFmtId");
                writer.Write("numFmtId");
                writer.Write(formatId is not null);
                if (formatId is not null)
                {
                    int id = ReadIndex(formatId);
                    if (formats.TryGetValue(id, out var definition))
                    {
                        writer.Write("custom");
                        if (!dependencyHashes.TryGetValue(("numFmtId", id), out string? formatHash))
                        {
                            formatHash = CanonicalHash(definition, ["numFmtId"], cancellationToken);
                            dependencyHashes.Add(("numFmtId", id), formatHash);
                        }
                        writer.Write(formatHash);
                    }
                    else if (id < 164)
                    {
                        writer.Write("builtin");
                        writer.Write(id);
                    }
                    else
                        throw new InvalidDataException("A parent style refers to a missing custom number format.");
                }
                writer.Flush();
                hash = Convert.ToHexStringLower(SHA256.HashData(buffer.GetBuffer().AsSpan(0, (int)buffer.Length)));
                hashes.Add(xfId, hash);
            }
            result.Add(new WorkbookPackageParentStyle { Name = name, XfId = xfId, DefinitionSha256 = hash });
        }
        return result;
    }

    private static List<XElement> ReadStyleRecords(XElement root, string collection, string record)
    {
        XNamespace ns = root.Name.Namespace;
        var containers = root.Elements(ns + collection).Take(2).ToArray();
        if (containers.Length > 1)
            throw new InvalidDataException("The styles part contains duplicate record collections.");
        if (containers.Length == 0)
            return [];
        var records = containers[0].Elements(ns + record).ToList();
        RequireRecordCount(records.Count);
        return records;
    }

    private static void WriteStyleReference(BinaryWriter writer, XElement xf, string attribute,
        List<XElement> records, Dictionary<(string Kind, int Index), string> hashes, CancellationToken cancellationToken)
    {
        string? reference = (string?)xf.Attribute(attribute);
        writer.Write(attribute);
        writer.Write(reference is not null);
        if (reference is null)
            return;
        int index = ReadIndex(reference);
        if (index >= records.Count)
            throw new InvalidDataException("A parent style refers to a missing font, fill or border definition.");
        if (!hashes.TryGetValue((attribute, index), out string? hash))
        {
            hash = CanonicalHash(records[index], [], cancellationToken);
            hashes.Add((attribute, index), hash);
        }
        writer.Write(hash);
    }

    private static List<WorkbookPackageParentStyleMapping> MapParentStyles(
        IReadOnlyList<WorkbookPackageParentStyle> before, IReadOnlyList<WorkbookPackageParentStyle> after,
        CancellationToken cancellationToken)
    {
        var result = new List<WorkbookPackageParentStyleMapping>();
        var afterByName = after.ToLookup(style => style.Name, StringComparer.Ordinal);
        var beforeByDefinition = before.ToLookup(style => (style.Name, style.DefinitionSha256));
        int candidateCount = 0;
        foreach (var style in before)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sameName = afterByName[style.Name].ToArray();
            var matches = sameName.Where(candidate => candidate.DefinitionSha256 == style.DefinitionSha256).ToArray();
            candidateCount += matches.Length;
            if (candidateCount > MaxRecords)
                throw new InvalidDataException("The package parent-style candidates exceed the report record limit.");
            bool ambiguous = matches.Length > 1 || beforeByDefinition[(style.Name, style.DefinitionSha256)].Skip(1).Any();
            result.Add(new WorkbookPackageParentStyleMapping
            {
                Name = style.Name,
                BeforeXfId = style.XfId,
                AfterXfId = !ambiguous && matches.Length == 1 ? matches[0].XfId : null,
                Status = ambiguous ? "ambiguous" : matches.Length == 1 ? "matched" : sameName.Length == 0 ? "removed" : "definition-changed",
                CandidateAfterXfIds = matches.Select(candidate => candidate.XfId).Distinct().Order().ToArray()
            });
        }
        var beforeNames = before.Select(style => style.Name).ToHashSet(StringComparer.Ordinal);
        foreach (var style in after.Where(style => !beforeNames.Contains(style.Name)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            result.Add(new WorkbookPackageParentStyleMapping
            {
                Name = style.Name,
                BeforeXfId = null,
                AfterXfId = style.XfId,
                Status = "added",
                CandidateAfterXfIds = [style.XfId]
            });
        }
        RequireRecordCount(result.Count);
        return result;
    }

    private static HashSet<string> ReachableParts(IReadOnlyList<WorkbookPackageRelationship> relationships,
        CancellationToken cancellationToken)
    {
        var byOwner = relationships.Where(item => !item.External && item.TargetExists == true)
            .ToLookup(item => item.OwnerPart, StringComparer.Ordinal);
        var reachable = new HashSet<string>(StringComparer.Ordinal);
        var pending = new Queue<string>();
        pending.Enqueue("");
        while (pending.TryDequeue(out string? owner))
        {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (var relationship in byOwner[owner])
            {
                if (reachable.Add(relationship.TargetPart!))
                    pending.Enqueue(relationship.TargetPart!);
            }
        }
        return reachable;
    }

    private static XElement ReadXml(IReadOnlyDictionary<string, byte[]> metadata, string part,
        string expectedRoot, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!metadata.TryGetValue(part, out byte[]? bytes))
            throw new InvalidDataException("Required structural metadata was not captured from the package.");
        try
        {
            using var stream = new MemoryStream(bytes, writable: false);
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = MaxMetadataBytes,
                IgnoreWhitespace = true
            };
            // Check depth before constructing an XElement tree used by recursive fingerprinting.
            using (var boundedReader = XmlReader.Create(stream, settings))
            {
                while (boundedReader.Read())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (boundedReader.Depth > 64)
                        throw new InvalidDataException("The package structural metadata exceeds the XML depth limit.");
                }
            }
            stream.Position = 0;
            using var reader = XmlReader.Create(stream, settings);
            var root = XDocument.Load(reader).Root;
            if (root is null || root.Name.LocalName != expectedRoot)
                throw new InvalidDataException("The package structural metadata has an unexpected root element.");
            cancellationToken.ThrowIfCancellationRequested();
            return root;
        }
        catch (XmlException exception)
        {
            throw new InvalidDataException("The package structural metadata XML is invalid.", exception);
        }
    }

    // Expanded XML names can repeat long namespace URIs far beyond the source XML size.
    private sealed class CanonicalBuffer : MemoryStream
    {
        private const int MaxCanonicalBytes = 8 * 1024 * 1024;

        public override void Write(byte[] buffer, int offset, int count)
        {
            EnsureWriteWithinLimit(count);
            base.Write(buffer, offset, count);
        }

        public override void Write(ReadOnlySpan<byte> buffer)
        {
            EnsureWriteWithinLimit(buffer.Length);
            base.Write(buffer);
        }

        public override void WriteByte(byte value)
        {
            EnsureWriteWithinLimit(1);
            base.WriteByte(value);
        }

        private void EnsureWriteWithinLimit(int count)
        {
            if (count > MaxCanonicalBytes - Position)
                throw new InvalidDataException("The canonical style definition exceeds the fingerprint size limit.");
        }
    }

    private static string CanonicalHash(XElement element, IReadOnlyCollection<string> excludedRootAttributes,
        CancellationToken cancellationToken)
    {
        using var buffer = new CanonicalBuffer();
        using var writer = new BinaryWriter(buffer, Encoding.UTF8, leaveOpen: true);
        WriteCanonical(writer, element, excludedRootAttributes, cancellationToken);
        writer.Flush();
        return Convert.ToHexStringLower(SHA256.HashData(buffer.GetBuffer().AsSpan(0, (int)buffer.Length)));
    }

    private static void WriteCanonical(BinaryWriter writer, XElement element, IReadOnlyCollection<string> excludedRootAttributes,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        writer.Write(element.Name.ToString());
        var attributes = element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration &&
            !(attribute.Name.Namespace == XNamespace.None && excludedRootAttributes.Contains(attribute.Name.LocalName)))
            .OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal).ToArray();
        writer.Write(attributes.Length);
        foreach (var attribute in attributes)
        {
            writer.Write(attribute.Name.ToString());
            writer.Write(attribute.Value);
        }
        var nodes = element.Nodes().Where(node => node is XElement || node is XText text && !string.IsNullOrWhiteSpace(text.Value)).ToArray();
        writer.Write(nodes.Length);
        foreach (var node in nodes)
        {
            writer.Write(node is XElement);
            if (node is XElement child)
                WriteCanonical(writer, child, [], cancellationToken);
            else
                writer.Write(((XText)node).Value);
        }
    }

    private static bool IsRelationshipsPart(string part) => part == "_rels/.rels" ||
        part.EndsWith(".rels", StringComparison.Ordinal) &&
        (part.StartsWith("_rels/", StringComparison.Ordinal) || part.Contains("/_rels/", StringComparison.Ordinal));

    private static string RelationshipOwner(string part)
    {
        if (part == "_rels/.rels")
            return "";
        int separator = part.LastIndexOf("/_rels/", StringComparison.Ordinal);
        string owner = separator < 0 ? part[6..^5] : part[..separator] + "/" + part[(separator + 7)..^5];
        if (owner.Length is 0 or > 512 || owner.EndsWith('/') || IsRelationshipsPart(owner))
            throw new InvalidDataException("The package relationships part has an invalid source owner.");
        return owner;
    }

    private static string ResolveTarget(string owner, string target)
    {
        if (target.Length > 512 || target.Contains('\\') || target.Contains('?') || target.Contains('#') ||
            target.Contains(':') || target.StartsWith("//", StringComparison.Ordinal) || target.Any(char.IsControl))
            throw new InvalidDataException("A package relationship contains an unsupported internal target.");
        string baseDirectory = owner.Contains('/') ? owner[..(owner.LastIndexOf('/') + 1)] : "";
        string combined = target.StartsWith('/') ? target[1..] : baseDirectory + target;
        var segments = new List<string>();
        foreach (string segment in combined.Split('/'))
        {
            if (segment == ".")
                continue;
            if (segment == "..")
            {
                if (segments.Count == 0)
                    throw new InvalidDataException("A package relationship escapes the package root.");
                segments.RemoveAt(segments.Count - 1);
                continue;
            }
            if (segment.Length == 0 || segment.EndsWith('.'))
                throw new InvalidDataException("A package relationship contains an ambiguous internal target.");
            for (int index = 0; index < segment.Length; index++)
            {
                char value = segment[index];
                if (char.IsAsciiLetterOrDigit(value) || "-._~!$&'()*+,;=@".Contains(value))
                    continue;
                if (value != '%' || index + 2 >= segment.Length ||
                    !byte.TryParse(segment.AsSpan(index + 1, 2), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out byte decoded) ||
                    decoded is < 32 or 127 or (byte)'/' or (byte)'\\' || char.IsAsciiLetterOrDigit((char)decoded) || "-._~".Contains((char)decoded))
                    throw new InvalidDataException("A package relationship contains an ambiguous encoded internal target.");
                index += 2;
            }
            segments.Add(segment);
        }
        string resolved = string.Join('/', segments);
        if (resolved.Length is 0 or > 512 || resolved == WorkbookPackageComparison.ContentTypesEntry || IsRelationshipsPart(resolved))
            throw new InvalidDataException("A package relationship contains an invalid target part.");
        return resolved;
    }

    private static string RequiredText(XElement element, string attribute)
    {
        string? value = (string?)element.Attribute(attribute);
        if (string.IsNullOrEmpty(value) || value.Length > MaxTextLength || value.Any(char.IsControl))
            throw new InvalidDataException("The package structure contains a missing or invalid identity attribute.");
        return value;
    }

    private static uint ReadUnsigned(string value)
    {
        if (!uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out uint parsed))
            throw new InvalidDataException("The package structure contains an invalid unsigned identity.");
        return parsed;
    }

    private static int ReadIndex(string value)
    {
        uint parsed = ReadUnsigned(value);
        if (parsed > int.MaxValue)
            throw new InvalidDataException("The package structure contains an invalid record index.");
        return (int)parsed;
    }

    private static bool IsRelationship(string type, string suffix) =>
        type == TransitionalRelationships + suffix || type == StrictRelationships + suffix;

    private static void RequireSpreadsheetNamespace(XNamespace ns)
    {
        if (ns != TransitionalSpreadsheet && ns != StrictSpreadsheet)
            throw new InvalidDataException("The package structure uses an unsupported spreadsheet namespace.");
    }

    private static void RequireRecordCount(int count)
    {
        if (count > MaxRecords)
            throw new InvalidDataException("The package structure exceeds the metadata record limit.");
    }
}
