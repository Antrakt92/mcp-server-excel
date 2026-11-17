using System.Buffers.Binary;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Utilities;

/// <summary>Read-only comparison of decompressed entry bytes in saved Excel packages.</summary>
public static class WorkbookPackageComparison
{
    private const int MaxArchiveBytes = 256 * 1024 * 1024;
    private const int MaxDirectoryBytes = 8 * 1024 * 1024;
    private const int MaxExpandedBytes = 1024 * 1024 * 1024;
    private const int MaxPartBytes = 256 * 1024 * 1024;
    private const int MaxMetadataBytes = 1024 * 1024;
    private const int MaxPartCount = 8192;
    private const int MaxReportedChanges = 1024;
    private const int MaxPartNameLength = 512;
    private const ushort DataDescriptor = 0x0008;
    private const ushort Stored = 0;
    private const ushort Deflated = 8;
    internal const string ContentTypesEntry = "[Content_Types].xml";
    private const string RootRelationshipsEntry = "_rels/.rels";
    private static readonly XNamespace ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
    private static readonly XNamespace RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly uint[] Crc32Table = CreateCrc32Table();

    /// <summary>Returns the complete difference and checks it against exact case-sensitive ZIP entry names.</summary>
    public static WorkbookPackageComparisonResult Compare(
        string beforePath, string afterPath, IReadOnlyList<string>? allowedParts, CancellationToken cancellationToken,
        bool includeStructure = false)
    {
        cancellationToken.ThrowIfCancellationRequested();
        string beforeFile = NormalizePackagePath(beforePath);
        string afterFile = NormalizePackagePath(afterPath);
        var allowlist = ValidateAllowlist(allowedParts);
        // Hold both inputs open for the comparison. On Windows this sharing mode excludes concurrent writers and renames.
        using var beforeStream = new FileStream(beforeFile, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var afterStream = new FileStream(afterFile, FileMode.Open, FileAccess.Read, FileShare.Read);
        return ComparePinned(beforeStream, afterStream, beforeFile, afterFile, allowlist, cancellationToken, includeStructure);
    }

    // The checked save keeps these file identities pinned until publication. The paths supply their expected formats.
    internal static WorkbookPackageComparisonResult Compare(
        FileStream beforeStream, FileStream afterStream, string beforePath, string afterPath,
        IReadOnlyList<string>? allowedParts, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return ComparePinned(beforeStream, afterStream, NormalizePackagePath(beforePath), NormalizePackagePath(afterPath),
            ValidateAllowlist(allowedParts), cancellationToken);
    }

    private static WorkbookPackageComparisonResult ComparePinned(
        FileStream beforeStream, FileStream afterStream, string beforeFile, string afterFile,
        HashSet<string> allowlist, CancellationToken cancellationToken, bool includeStructure = false)
    {
        var beforeInventory = ReadInventory(beforeStream, Path.GetExtension(beforeFile), cancellationToken, includeStructure);
        var afterInventory = ReadInventory(afterStream, Path.GetExtension(afterFile), cancellationToken, includeStructure);
        var before = beforeInventory.Hashes;
        var after = afterInventory.Hashes;
        var changes = new List<WorkbookPackagePartChange>();

        foreach (string name in before.Keys.Union(after.Keys, StringComparer.Ordinal).Order(StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            before.TryGetValue(name, out string? beforeHash);
            after.TryGetValue(name, out string? afterHash);
            if (beforeHash == afterHash)
                continue;
            if (changes.Count == MaxReportedChanges)
                throw new InvalidDataException($"The package comparison has too many changes to report (maximum {MaxReportedChanges}).");

            changes.Add(new WorkbookPackagePartChange
            {
                PartName = name,
                ChangeKind = beforeHash is null ? "added" : afterHash is null ? "removed" : "changed",
                BeforeSha256 = beforeHash,
                AfterSha256 = afterHash,
                Allowed = allowlist.Contains(name)
            });
        }

        cancellationToken.ThrowIfCancellationRequested();
        var structure = includeStructure ? WorkbookPackageStructure.Compare(
            beforeInventory.Metadata, before.Keys, beforeInventory.CodeNames,
            afterInventory.Metadata, after.Keys, afterInventory.CodeNames, cancellationToken) : null;
        if (structure is not null && JsonSerializer.SerializeToUtf8Bytes(structure).Length > 4 * 1024 * 1024)
            throw new InvalidDataException("The package structure report exceeds the 4 MiB output limit.");
        return new WorkbookPackageComparisonResult
        {
            Success = true,
            BeforePartCount = before.Count,
            AfterPartCount = after.Count,
            MatchesAllowlist = changes.All(part => part.Allowed),
            Changes = changes,
            Structure = structure
        };
    }

    internal static string NormalizePackagePath(string path)
    {
        string file = FilePathValidation.NormalizeAbsoluteWindowsPath(path);
        string root = Path.GetPathRoot(file)!;
        if (file.StartsWith(@"\\.\", StringComparison.Ordinal) ||
            file.StartsWith(@"\\?\", StringComparison.Ordinal) ||
            file.AsSpan(root.Length).Contains(':') || IsReservedFileName(file) || IsIpcShare(root))
            throw new ArgumentException("Package comparison accepts regular file paths. Device paths and alternate data streams are unsupported.", nameof(path));

        string extension = Path.GetExtension(file);
        if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Package comparison accepts saved .xlsx and .xlsm files.", nameof(path));
        return file;
    }

    private static bool IsIpcShare(string root)
    {
        if (!root.StartsWith(@"\\", StringComparison.Ordinal))
            return false;
        var rootName = root.AsSpan().TrimEnd('\\');
        var share = rootName[(rootName.LastIndexOf('\\') + 1)..].TrimEnd(" .");
        return share.Equals("PIPE", StringComparison.OrdinalIgnoreCase) ||
            share.Equals("MAILSLOT", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsReservedFileName(string path)
    {
        var name = Path.GetFileName(path.AsSpan());
        int extension = name.IndexOf('.');
        var stem = (extension < 0 ? name : name[..extension]).TrimEnd(' ');
        if (stem.Equals("CON", StringComparison.OrdinalIgnoreCase) ||
            stem.Equals("PRN", StringComparison.OrdinalIgnoreCase) ||
            stem.Equals("AUX", StringComparison.OrdinalIgnoreCase) ||
            stem.Equals("NUL", StringComparison.OrdinalIgnoreCase) ||
            stem.Equals("CONIN$", StringComparison.OrdinalIgnoreCase) ||
            stem.Equals("CONOUT$", StringComparison.OrdinalIgnoreCase))
            return true;
        return stem.Length == 4 && (stem[3] is (>= '1' and <= '9') or '¹' or '²' or '³')
            && (stem[..3].Equals("COM", StringComparison.OrdinalIgnoreCase) ||
                stem[..3].Equals("LPT", StringComparison.OrdinalIgnoreCase));
    }

    internal static HashSet<string> ValidateAllowlist(IReadOnlyList<string>? allowedParts)
    {
        if (allowedParts?.Count > MaxReportedChanges)
            throw new ArgumentException($"The allowlist may contain at most {MaxReportedChanges} part names.", nameof(allowedParts));

        var allowlist = new HashSet<string>(StringComparer.Ordinal);
        if (allowedParts is null)
            return allowlist;
        foreach (string name in allowedParts)
        {
            if (!IsValidEntryName(name) || !allowlist.Add(name))
                throw new ArgumentException("The allowlist must contain exact, unique package part names.", nameof(allowedParts));
        }
        return allowlist;
    }

    private sealed record PackageInventory(Dictionary<string, string> Hashes,
        Dictionary<string, byte[]> Metadata, Dictionary<string, string?> CodeNames);

    private static PackageInventory ReadInventory(
        FileStream stream, string extension, CancellationToken cancellationToken, bool includeStructure)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var directory = ReadZipDirectory(stream);
        var entries = ReadZipEntries(stream, directory, cancellationToken);
        var equivalentNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string name in entries.Keys)
        {
            if (!equivalentNames.Add(name))
                throw new InvalidDataException("The package contains an invalid or ambiguous part name.");
        }
        foreach (string name in entries.Keys)
        {
            for (int separator = name.IndexOf('/'); separator >= 0; separator = name.IndexOf('/', separator + 1))
            {
                if (equivalentNames.Contains(name[..separator]))
                    throw new InvalidDataException("The package contains a part name that is also a part directory.");
            }
        }
        var contentTypes = RequireMetadata(entries, ContentTypesEntry);
        var relationships = RequireMetadata(entries, RootRelationshipsEntry);

        var hashes = new Dictionary<string, string>(entries.Count, StringComparer.Ordinal);
        var buffer = new byte[64 * 1024];
        using var digest = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var metadata = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        var codeNames = new Dictionary<string, string?>(StringComparer.Ordinal);
        var typesBytes = new byte[contentTypes.Length];
        hashes.Add(contentTypes.Name, HashEntry(stream, contentTypes, digest, buffer, typesBytes, cancellationToken));
        var types = ReadMetadata(contentTypes, ContentTypesNamespace + "Types", typesBytes);
        (Dictionary<string, string> Overrides, Dictionary<string, string> Defaults)? partTypes =
            includeStructure ? ReadContentTypes(types) : null;
        if (includeStructure)
            metadata.Add(contentTypes.Name, typesBytes);
        long metadataBytes = typesBytes.Length;
        XElement? root = null;
        foreach (var entry in entries.Values.OrderBy(entry => entry.DataOffset))
        {
            if (entry == contentTypes)
                continue;
            string? contentType = partTypes?.Overrides.GetValueOrDefault("/" + entry.Name) ??
                partTypes?.Defaults.GetValueOrDefault(Path.GetExtension(entry.Name).TrimStart('.'));
            bool isMetadata = entry == relationships || includeStructure &&
                (entry.Name.EndsWith(".rels", StringComparison.Ordinal) || IsStructuralMetadata(contentType));
            byte[]? captured = null;
            if (isMetadata)
            {
                if (entry.Length > 2 * MaxMetadataBytes || includeStructure && metadataBytes + entry.Length > 16 * MaxMetadataBytes)
                    throw new InvalidDataException("The package exceeds the structural metadata size limit.");
                captured = new byte[entry.Length];
                metadataBytes += entry.Length;
            }
            hashes.Add(entry.Name, HashEntry(stream, entry, digest, buffer, captured, cancellationToken));
            if (entry == relationships)
                root = ReadMetadata(entry, RelationshipsNamespace + "Relationships", captured!);
            if (includeStructure && captured is not null)
                metadata.Add(entry.Name, captured);
            if (includeStructure && IsSheetContentType(contentType))
                codeNames.Add(entry.Name, ReadSheetCodeName(stream, entry, cancellationToken));
        }

        ValidateWorkbookIdentity(types, root!, equivalentNames, extension);
        return new PackageInventory(hashes, metadata, codeNames);
    }

    private static (Dictionary<string, string> Overrides, Dictionary<string, string> Defaults) ReadContentTypes(XElement types)
    {
        var overrides = new Dictionary<string, string>(StringComparer.Ordinal);
        var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in types.Elements())
        {
            bool isOverride = item.Name == ContentTypesNamespace + "Override";
            if (!isOverride && item.Name != ContentTypesNamespace + "Default")
                throw new InvalidDataException("The package content types contain an unexpected element.");
            string? key = (string?)item.Attribute(isOverride ? "PartName" : "Extension");
            string? type = (string?)item.Attribute("ContentType");
            if (string.IsNullOrEmpty(key) || string.IsNullOrEmpty(type) ||
                !(isOverride ? overrides : defaults).TryAdd(key, type))
                throw new InvalidDataException("The package content types are missing or ambiguous.");
        }
        return (overrides, defaults);
    }

    private static bool IsStructuralMetadata(string? type) => type is
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" or
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml" or
        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";

    private static bool IsSheetContentType(string? type) => type is
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" or
        "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml" or
        "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml" or
        "application/vnd.ms-excel.macrosheet+xml" or
        "application/vnd.ms-excel.intlmacrosheet+xml";

    private static string? ReadSheetCodeName(FileStream file, PackageZipEntry entry, CancellationToken cancellationToken)
    {
        try
        {
            using var payload = new ZipPayloadStream(file, entry.DataOffset, entry.CompressedLength,
                entry.Method == Deflated, cancellationToken);
            using var inflated = entry.Method == Deflated ? new DeflateStream(payload, CompressionMode.Decompress, leaveOpen: true) : null;
            using var reader = XmlReader.Create(inflated ?? (Stream)payload, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = 2 * MaxMetadataBytes,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = true
            });
            reader.MoveToContent();
            string ns = reader.NamespaceURI;
            if (reader.LocalName is not ("worksheet" or "chartsheet" or "dialogsheet" or "macrosheet") || ns is not
                ("http://schemas.openxmlformats.org/spreadsheetml/2006/main" or "http://purl.oclc.org/ooxml/spreadsheetml/main"))
                throw new InvalidDataException("A sheet part has an unexpected XML root.");
            if (reader.IsEmptyElement)
                return null;
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (reader.NodeType != XmlNodeType.Element)
                    continue;
                if (reader.NamespaceURI != ns)
                    throw new InvalidDataException("The sheet has unsupported leading extension metadata.");
                // sheetPr is the first optional child. Stop before reading sheet cells.
                if (reader.Depth != 1 || reader.LocalName != "sheetPr" || reader.NamespaceURI != ns)
                    return null;
                string? codeName = reader.GetAttribute("codeName");
                if (codeName?.Length > 512)
                    throw new InvalidDataException("A sheet CodeName exceeds the structural name limit.");
                return codeName;
            }
            return null;
        }
        catch (XmlException exception)
        {
            throw new InvalidDataException("The sheet structural metadata XML is invalid.", exception);
        }
    }

    private static (int Count, long Start, long End) ReadZipDirectory(FileStream stream)
    {
        const int footerLength = 22;
        const uint footerSignature = 0x06054b50;
        if (stream.Length < footerLength || stream.Length > MaxArchiveBytes)
            throw new InvalidDataException("The package is empty, invalid, or exceeds the ZIP size limit.");

        // Split, ZIP64, and prefixed ZIPs are outside this size-limited format.
        Span<byte> signature = stackalloc byte[4];
        stream.Position = 0;
        stream.ReadExactly(signature);
        if (BinaryPrimitives.ReadUInt32LittleEndian(signature) != 0x04034b50)
            throw new InvalidDataException("Expected a saved, unencrypted workbook ZIP package.");

        var tail = new byte[(int)Math.Min(stream.Length, footerLength + ushort.MaxValue)];
        long tailOffset = stream.Length - tail.Length;
        stream.Position = tailOffset;
        stream.ReadExactly(tail);
        for (int index = tail.Length - footerLength; index >= 0; index--)
        {
            var footer = tail.AsSpan(index, footerLength);
            if (BinaryPrimitives.ReadUInt32LittleEndian(footer) != footerSignature)
                continue;
            int count = BinaryPrimitives.ReadUInt16LittleEndian(footer[10..]);
            long length = BinaryPrimitives.ReadUInt32LittleEndian(footer[12..]);
            long start = BinaryPrimitives.ReadUInt32LittleEndian(footer[16..]);
            int commentLength = BinaryPrimitives.ReadUInt16LittleEndian(footer[20..]);
            if (BinaryPrimitives.ReadUInt32LittleEndian(footer[4..]) != 0 ||
                BinaryPrimitives.ReadUInt16LittleEndian(footer[8..]) != count ||
                count is < 1 or > MaxPartCount || length < 46L * count || length > MaxDirectoryBytes || start < 30 ||
                start + length != tailOffset + index || commentLength != tail.Length - index - footerLength)
                throw new InvalidDataException("The package ZIP directory is unsupported, invalid, or exceeds the part-count limit.");
            return (count, start, start + length);
        }
        throw new InvalidDataException("The package does not have a valid ZIP directory.");
    }

    private sealed record PackageZipEntry(string Name, ushort Flags, ushort Method, uint Crc32,
        uint Length, uint CompressedLength, uint HeaderOffset)
    {
        public long DataOffset { get; set; }
    }

    private static Dictionary<string, PackageZipEntry> ReadZipEntries(
        FileStream stream, (int Count, long Start, long End) directory, CancellationToken cancellationToken)
    {
        var entries = new Dictionary<string, PackageZipEntry>(directory.Count, StringComparer.Ordinal);
        stream.Position = directory.Start;
        Span<byte> header = stackalloc byte[46];
        Span<byte> nameBuffer = stackalloc byte[MaxPartNameLength];
        long compressedBytes = 0;
        long expandedBytes = 0;
        for (int index = 0; index < directory.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (directory.End - stream.Position < header.Length)
                throw new InvalidDataException("The ZIP central directory contains a truncated entry.");
            stream.ReadExactly(header);
            ushort flags = BinaryPrimitives.ReadUInt16LittleEndian(header[8..]);
            ushort method = BinaryPrimitives.ReadUInt16LittleEndian(header[10..]);
            uint crc = BinaryPrimitives.ReadUInt32LittleEndian(header[16..]);
            uint compressed = BinaryPrimitives.ReadUInt32LittleEndian(header[20..]);
            uint expanded = BinaryPrimitives.ReadUInt32LittleEndian(header[24..]);
            int nameLength = BinaryPrimitives.ReadUInt16LittleEndian(header[28..]);
            int extraLength = BinaryPrimitives.ReadUInt16LittleEndian(header[30..]);
            int commentLength = BinaryPrimitives.ReadUInt16LittleEndian(header[32..]);
            uint headerOffset = BinaryPrimitives.ReadUInt32LittleEndian(header[42..]);

            if (BinaryPrimitives.ReadUInt32LittleEndian(header) != 0x02014b50 ||
                BinaryPrimitives.ReadUInt16LittleEndian(header[34..]) != 0 ||
                (flags & ~0x080e) != 0 || method is not (Stored or Deflated) ||
                method == Stored && compressed != expanded || method == Deflated && compressed == 0 ||
                headerOffset >= directory.Start)
                throw new InvalidDataException("The package contains an invalid, encrypted, or unsupported ZIP entry.");
            if (expanded > MaxPartBytes || expanded > MaxExpandedBytes - expandedBytes)
                throw new InvalidDataException("The package exceeds the decompressed size limit.");
            if (compressed > directory.Start - compressedBytes)
                throw new InvalidDataException("The combined compressed ZIP parts exceed the physical package data size.");
            if (nameLength is < 1 or > MaxPartNameLength || nameLength + extraLength + commentLength > directory.End - stream.Position)
                throw new InvalidDataException("The ZIP central directory has an invalid part name or record size.");
            stream.ReadExactly(nameBuffer[..nameLength]);
            string name = Encoding.ASCII.GetString(nameBuffer[..nameLength]);
            if (!IsValidEntryName(name) || !entries.TryAdd(name,
                new PackageZipEntry(name, flags, method, crc, expanded, compressed, headerOffset)))
                throw new InvalidDataException("The package contains an invalid or ambiguous part name.");
            stream.Position += extraLength + commentLength;
            compressedBytes += compressed;
            expandedBytes += expanded;
        }
        if (stream.Position != directory.End)
            throw new InvalidDataException("The ZIP central directory length does not match its entries.");

        long previousEnd = 0;
        foreach (var entry in entries.Values.OrderBy(entry => entry.HeaderOffset))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.HeaderOffset < previousEnd)
                throw new InvalidDataException("The package contains overlapping ZIP entry records.");
            previousEnd = ReadLocalZipEntry(stream, entry, directory.Start);
        }
        return entries;
    }

    private static long ReadLocalZipEntry(FileStream stream, PackageZipEntry entry, long directoryStart)
    {
        Span<byte> header = stackalloc byte[30];
        if (directoryStart - entry.HeaderOffset < header.Length)
            throw new InvalidDataException("A ZIP local entry header overlaps the central directory.");
        stream.Position = entry.HeaderOffset;
        stream.ReadExactly(header);
        if (BinaryPrimitives.ReadUInt32LittleEndian(header) != 0x04034b50 ||
            BinaryPrimitives.ReadUInt16LittleEndian(header[6..]) != entry.Flags ||
            BinaryPrimitives.ReadUInt16LittleEndian(header[8..]) != entry.Method)
            throw new InvalidDataException("A ZIP local entry header conflicts with its central directory entry.");

        int nameLength = BinaryPrimitives.ReadUInt16LittleEndian(header[26..]);
        int extraLength = BinaryPrimitives.ReadUInt16LittleEndian(header[28..]);
        entry.DataOffset = stream.Position + nameLength + extraLength;
        long end = entry.DataOffset + entry.CompressedLength;
        if (nameLength != entry.Name.Length || end > directoryStart)
            throw new InvalidDataException("A ZIP local part name or data extent conflicts with the central directory.");
        Span<byte> name = stackalloc byte[nameLength];
        stream.ReadExactly(name);
        if (!entry.Name.Equals(Encoding.ASCII.GetString(name), StringComparison.Ordinal))
            throw new InvalidDataException("A ZIP local part name conflicts with the central directory.");

        bool hasDescriptor = (entry.Flags & DataDescriptor) != 0;
        if (!MatchesEntrySizes(header[14..], entry, allowZero: hasDescriptor))
            throw new InvalidDataException("A ZIP local entry CRC or size conflicts with the central directory.");
        return hasDescriptor ? ReadDataDescriptor(stream, entry, end, directoryStart) : end;
    }

    private static bool MatchesEntrySizes(ReadOnlySpan<byte> data, PackageZipEntry entry, bool allowZero) =>
        MatchesZipField(data, entry.Crc32, allowZero) &&
        MatchesZipField(data[4..], entry.CompressedLength, allowZero) &&
        MatchesZipField(data[8..], entry.Length, allowZero);

    private static bool MatchesZipField(ReadOnlySpan<byte> data, uint expected, bool allowZero)
    {
        uint actual = BinaryPrimitives.ReadUInt32LittleEndian(data);
        return actual == expected || allowZero && actual == 0;
    }

    private static long ReadDataDescriptor(FileStream stream, PackageZipEntry entry, long position, long directoryStart)
    {
        int available = (int)Math.Min(16, directoryStart - position);
        if (available >= 12)
        {
            Span<byte> descriptor = stackalloc byte[16];
            stream.Position = position;
            stream.ReadExactly(descriptor[..available]);
            if (available == 16 && BinaryPrimitives.ReadUInt32LittleEndian(descriptor) == 0x08074b50 &&
                MatchesEntrySizes(descriptor[4..], entry, allowZero: false))
                return position + 16;
            if (MatchesEntrySizes(descriptor, entry, allowZero: false))
                return position + 12;
        }
        throw new InvalidDataException("A ZIP entry data descriptor conflicts with the central directory.");
    }

    private static bool IsValidEntryName(string? name)
    {
        if (name == ContentTypesEntry)
            return true;
        if (string.IsNullOrEmpty(name) || name.Length > MaxPartNameLength)
            return false;

        foreach (string segment in name.Split('/'))
        {
            if (segment.Length == 0 || segment[^1] == '.')
                return false;
            for (int index = 0; index < segment.Length; index++)
            {
                char character = segment[index];
                if (IsLiteralPartCharacter(character))
                    continue;
                if (character != '%' || index + 2 >= segment.Length ||
                    !byte.TryParse(segment.AsSpan(index + 1, 2), System.Globalization.NumberStyles.AllowHexSpecifier,
                        System.Globalization.CultureInfo.InvariantCulture, out byte decoded) ||
                    decoded is < 32 or 127 or (byte)'/' or (byte)'\\' ||
                    IsUnreservedPartCharacter((char)decoded))
                    return false;
                index += 2;
            }
        }
        return true;
    }

    private static bool IsLiteralPartCharacter(char value) =>
        IsUnreservedPartCharacter(value) || "!$&'()*+,;=:@".Contains(value);

    private static bool IsUnreservedPartCharacter(char value) =>
        char.IsAsciiLetterOrDigit(value) || "-._~".Contains(value);

    private static PackageZipEntry RequireMetadata(Dictionary<string, PackageZipEntry> entries, string name)
    {
        if (!entries.TryGetValue(name, out var entry) || entry.Length is <= 0 or > MaxMetadataBytes)
            throw new InvalidDataException($"The required package part '{name}' is missing or exceeds the metadata size limit.");
        return entry;
    }

    private static void ValidateWorkbookIdentity(XElement types, XElement root,
        HashSet<string> equivalentNames, string extension)
    {
        var officeDocuments = root.Elements(RelationshipsNamespace + "Relationship")
            .Where(element => (string?)element.Attribute("Type") is
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" or
                "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument")
            .Take(2)
            .ToList();
        if (officeDocuments.Count != 1 || (string?)officeDocuments[0].Attribute("TargetMode") is not (null or "Internal"))
            throw new InvalidDataException("The package must have one internal root officeDocument relationship.");

        string target = (string?)officeDocuments[0].Attribute("Target") ?? "";
        if (target.StartsWith('/'))
            target = target[1..];
        else
        {
            int slash = target.IndexOf('/');
            var firstSegment = slash >= 0 ? target.AsSpan(0, slash) : target.AsSpan();
            if (firstSegment.Contains(':'))
                throw new InvalidDataException("The root officeDocument relationship must identify an internal workbook package part.");
        }
        if (target is ContentTypesEntry or RootRelationshipsEntry || !IsValidEntryName(target) ||
            !equivalentNames.Contains(target))
            throw new InvalidDataException("The root officeDocument relationship does not identify a workbook package part.");

        var matchingTypes = types.Elements(ContentTypesNamespace + "Override")
            .Where(element => string.Equals((string?)element.Attribute("PartName"), "/" + target, StringComparison.OrdinalIgnoreCase))
            .Take(2)
            .ToList();
        string expectedType = extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase)
            ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
            : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        if (matchingTypes.Count != 1 ||
            !string.Equals((string?)matchingTypes[0].Attribute("ContentType"), expectedType, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("The package workbook content type does not match its .xlsx or .xlsm extension.");
    }

    private static XElement ReadMetadata(PackageZipEntry entry, XName expectedRoot, byte[] buffer)
    {
        using var stream = new MemoryStream(buffer, 0, (int)entry.Length, writable: false);
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = MaxMetadataBytes
        });
        try
        {
            var root = XDocument.Load(reader).Root;
            if (root?.Name != expectedRoot)
                throw new InvalidDataException($"The package metadata XML for '{entry.Name}' has an unexpected root.");
            return root;
        }
        catch (XmlException exception)
        {
            throw new InvalidDataException($"The package metadata XML for '{entry.Name}' is invalid.", exception);
        }
    }

    private static string HashEntry(FileStream file, PackageZipEntry entry, IncrementalHash digest, byte[] buffer,
        byte[]? metadata, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        using var payload = new ZipPayloadStream(file, entry.DataOffset, entry.CompressedLength,
            entry.Method == Deflated, cancellationToken);
        using var inflated = entry.Method == Deflated ? new DeflateStream(payload, CompressionMode.Decompress, leaveOpen: true) : null;
        Stream stream = inflated ?? (Stream)payload;
        long remaining = entry.Length;
        uint crc = uint.MaxValue;
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int bytesRead = stream.Read(buffer);
            if (bytesRead == 0)
                break;
            remaining -= bytesRead;
            if (remaining < 0)
                throw new InvalidDataException("A package part exceeds its declared decompressed size.");
            if (metadata is not null)
                buffer.AsSpan(0, bytesRead).CopyTo(metadata.AsSpan((int)(entry.Length - remaining - bytesRead)));
            digest.AppendData(buffer, 0, bytesRead);
            for (int index = 0; index < bytesRead; index++)
                crc = Crc32Table[(byte)(crc ^ buffer[index])] ^ (crc >> 8);
        }
        if (remaining != 0)
            throw new InvalidDataException("A package part is shorter than its declared decompressed size.");
        if (payload.HasUnreadBytes)
            throw new InvalidDataException("A package part did not consume its declared compressed size.");
        if (~crc != entry.Crc32)
            throw new InvalidDataException("A package part failed ZIP CRC verification.");
        return Convert.ToHexStringLower(digest.GetHashAndReset());
    }

    // DeflateStream must see the exact compressed slice but no output limit from the ZIP metadata.
    private sealed class ZipPayloadStream : Stream
    {
        private readonly FileStream _file;
        private readonly bool _requireDeflateTerminator;
        private readonly CancellationToken _cancellationToken;
        private long _remaining;

        public ZipPayloadStream(FileStream file, long offset, long length, bool requireDeflateTerminator, CancellationToken cancellationToken)
        {
            _file = file;
            _file.Position = offset;
            _remaining = length;
            _requireDeflateTerminator = requireDeflateTerminator;
            _cancellationToken = cancellationToken;
        }

        public bool HasUnreadBytes => _remaining > 0;
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override int Read(byte[] buffer, int offset, int count) => Read(buffer.AsSpan(offset, count));

        public override int Read(Span<byte> buffer)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            if (buffer.IsEmpty)
                return 0;
            if (_remaining == 0)
            {
                if (_requireDeflateTerminator)
                    throw new InvalidDataException("A package part ended without a complete DEFLATE stream.");
                return 0;
            }
            int read = _file.Read(buffer[..(int)Math.Min(buffer.Length, _remaining)]);
            if (read == 0)
                throw new InvalidDataException("A package part is shorter than its declared compressed size.");
            _remaining -= read;
            return read;
        }
    }

    private static uint[] CreateCrc32Table()
    {
        var table = new uint[256];
        for (uint index = 0; index < table.Length; index++)
        {
            uint value = index;
            for (int bit = 0; bit < 8; bit++)
                value = (value & 1) != 0 ? (value >> 1) ^ 0xedb88320 : value >> 1;
            table[index] = value;
        }
        return table;
    }
}
