using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "WorkbookPackageComparison")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class WorkbookPackageVbaComparisonTests : IDisposable
{
    private const string ProjectPart = "xl/macros/fixture.bin";
    private const uint End = 0xFFFFFFFE;
    private const uint Free = 0xFFFFFFFF;
    private const int SectorSize = 512;
    private readonly List<string> _fixtures = [];

    [Theory]
    [InlineData("source", "module-source", "FixtureForm")]
    [InlineData("designer", "designer-stream", "FixtureForm/f")]
    [InlineData("object", "designer-stream", "FixtureForm/o")]
    [InlineData("module-cache", "module-cache", "FixtureForm")]
    [InlineData("project-cache", "performance-cache", "VBA/_VBA_PROJECT")]
    public void CompareVbaProject_SeparatesEachKindAndKeepsTheExactZipAllowlist(string scenario, string category, string name)
    {
        var beforeBytes = Project();
        var afterBytes = Project(
            source: scenario == "source" ? "changed-source" : "fixture-source",
            form: scenario == "designer" ? "changed-form" : "fixture-form",
            obj: scenario == "object" ? "changed-object" : "fixture-object",
            moduleCache: scenario == "module-cache" ? "larger-cache" : "cache",
            projectCache: scenario == "project-cache" ? (byte)2 : (byte)1);
        string before = Workbook(beforeBytes);
        string after = Workbook(afterBytes);

        Assert.Null(WorkbookPackageComparison.Compare(before, after, [ProjectPart], default).VbaProject);
        var result = WorkbookPackageComparison.Compare(before, after, [ProjectPart], default, includeVbaProject: true);
        Assert.True(result.MatchesAllowlist);
        Assert.Equal(ProjectPart, Assert.Single(result.Changes).PartName);
        var vba = Assert.IsType<WorkbookVbaProjectComparison>(result.VbaProject);
        Assert.Equal("compared", vba.Status);
        Assert.Equal((ProjectPart, ProjectPart), (vba.BeforePart, vba.AfterPart));
        Assert.Equal(("unchanged", "unchanged", 0), (vba.References, vba.ProtectionMetadata, vba.OtherChangedStreams));
        Assert.Equal(new WorkbookVbaProjectItemChange(category, name, "changed"), Assert.Single(vba.Changes));
        string json = JsonSerializer.Serialize(vba);
        Assert.DoesNotContain("changed-source", json);
        Assert.DoesNotContain("changed-form", json);
        Assert.Equal(beforeBytes, ReadVbaPart(before));
        Assert.Equal(afterBytes, ReadVbaPart(after));
    }

    [Fact]
    public void CompareVbaProject_KeepsReferenceAndProtectionValuesOutOfTheReport()
    {
        string before = Workbook(Project());
        string after = Workbook(Project(reference: "different-reference-target", protection: "1234"));
        var vba = WorkbookPackageComparison.Compare(before, after, null, default, includeVbaProject: true).VbaProject!;

        Assert.Equal("compared", vba.Status);
        Assert.Equal("changed", vba.References);
        Assert.Equal(1, vba.BeforeReferenceCount);
        Assert.Equal(1, vba.AfterReferenceCount);
        Assert.Equal("changed", vba.ProtectionMetadata);
        Assert.Empty(vba.Changes);
        string json = JsonSerializer.Serialize(vba);
        Assert.DoesNotContain("different-reference-target", json);
        Assert.DoesNotContain("1234", json);
    }

    [Fact]
    public void CompareVbaProject_ReportsIdenticalAbsentAddedRemovedAndUnavailable()
    {
        string before = Workbook(Project());
        string missing = Workbook(null);
        string unsupported = Workbook([1, 2, 3]);
        string unresolved = Workbook(Project(), external: true);
        string unlinked = Workbook(Project(), relationship: false);

        WorkbookVbaProjectComparison Compare(string left, string right) =>
            WorkbookPackageComparison.Compare(left, right, [ProjectPart], default, includeVbaProject: true).VbaProject!;

        Assert.Equal("identical-bytes", Compare(before, before).Status);
        Assert.Equal("absent", Compare(missing, missing).Status);
        Assert.Equal("absent", Compare(unlinked, unlinked).Status);
        var added = Compare(missing, before);
        Assert.Equal("compared", added.Status);
        Assert.Null(added.BeforePart);
        Assert.Equal(ProjectPart, added.AfterPart);
        Assert.Contains(new WorkbookVbaProjectItemChange("module-source", "FixtureForm", "added"), added.Changes);
        var removed = Compare(before, missing);
        Assert.Contains(new WorkbookVbaProjectItemChange("module-source", "FixtureForm", "removed"), removed.Changes);
        var parseFailure = WorkbookPackageComparison.Compare(before, unsupported, [ProjectPart], default, includeVbaProject: true);
        Assert.True(parseFailure.MatchesAllowlist);
        Assert.Equal("unavailable", parseFailure.VbaProject!.Status);
        Assert.Equal("vba-project-truncated", parseFailure.VbaProject.Reason);
        Assert.Equal("unconfirmed-vba-project-relationship", Compare(before, unresolved).Reason);
    }

    [Fact]
    public void Decompress_ExpandsRepeatedBytesAndChecksTheLimitAndCancellation()
    {
        byte[] compressed = [1, 3, 0xB0, 2, (byte)'A', 1, 0];
        Assert.Equal("AAAAA"u8.ToArray(), VbaProjectDirectory.Decompress(compressed, 5, default));
        Assert.Throws<InvalidDataException>(() => VbaProjectDirectory.Decompress(compressed, 4, default));
        Assert.Throws<OperationCanceledException>(() => VbaProjectDirectory.Decompress(compressed, 5, new(true)));
    }

    [Fact]
    public void Decompress_AcceptsExcelFinalFlagsWithoutMoreTokens()
    {
        byte[] compressed = [1, 9, 0xB0, 0, .. "ABCDEFGH"u8, 0];
        Assert.Equal("ABCDEFGH"u8.ToArray(), VbaProjectDirectory.Decompress(compressed, 8, default));
    }

    [Fact]
    public void CompareVbaProject_UsesLogicalFormNameAndDecodesSourceCodePage()
    {
        const string streamName = "FormStorage";
        string before = Workbook(Project(source: "'é", streamName: streamName, includeConstants: false));
        string reencoded = Workbook(Project(source: "'é", streamName: streamName, codePage: 65001));
        string changed = Workbook(Project(source: "'й", streamName: streamName, codePage: 1251, form: "changed-form"));
        WorkbookVbaProjectComparison Compare(string candidate) =>
            WorkbookPackageComparison.Compare(before, candidate, null, default, includeVbaProject: true).VbaProject!;

        var sameText = Compare(reencoded);
        Assert.Equal("compared", sameText.Status);
        Assert.Contains(sameText.Changes, change => change.Category == "project-metadata");
        Assert.DoesNotContain(sameText.Changes, change => change.Category == "module-source");
        var differentText = Compare(changed);
        Assert.Equal("compared", differentText.Status);
        Assert.Contains(new WorkbookVbaProjectItemChange("module-source", "FixtureForm", "changed"), differentText.Changes);
        Assert.Contains(new WorkbookVbaProjectItemChange("designer-stream", "FixtureForm/f", "changed"), differentText.Changes);
    }

    private static byte[] Project(string source = "fixture-source", string form = "fixture-form", string obj = "fixture-object",
        string moduleCache = "cache", byte projectCache = 1, string reference = "fixture-reference", string protection = "0000",
        string streamName = "FixtureForm", ushort codePage = 1252, bool includeConstants = true)
    {
        byte[] cache = Encoding.ASCII.GetBytes(moduleCache);
        byte[] properties = Encoding.ASCII.GetBytes($"Name=\"FixtureProject\"\r\nBaseClass=FixtureForm\r\nCMG=\"{protection}\"\r\nDPB=\"2468\"\r\nGC=\"8642\"\r\n");
        byte[] dir = DirectoryRecords(cache.Length, reference, streamName, codePage, includeConstants);
        var encoding = codePage == 65001 ? Encoding.UTF8 : CodePagesEncodingProvider.Instance.GetEncoding(codePage)!;
        byte[] module = [.. cache, .. Compressed(encoding.GetBytes(source))];

        // A test-only CFB with one FAT sector, three directory sectors and a mini-stream for all streams.
        byte[][] streams =
        [
            properties, dir, module, [0xCC, 0x61, 0, 0, 0, 0, 0, projectCache],
            Encoding.ASCII.GetBytes(form), Encoding.ASCII.GetBytes(obj), [1]
        ];
        int[] miniStarts = new int[streams.Length];
        int miniCount = 0;
        for (int i = 0; i < streams.Length; i++)
        {
            miniStarts[i] = miniCount;
            miniCount += (streams[i].Length + 63) / 64;
        }
        int rootBytes = miniCount * 64;
        int rootSectors = (rootBytes + SectorSize - 1) / SectorSize;
        byte[] bytes = new byte[(6 + rootSectors) * SectorSize];
        byte[] signature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        signature.CopyTo(bytes, 0);
        void Header16(int offset, ushort value) => BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(offset), value);
        void Header32(int offset, uint value) => BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(offset), value);
        Header16(24, 0x003E);
        Header16(26, 3);
        Header16(28, 0xFFFE);
        Header16(30, 9);
        Header16(32, 6);
        Header32(44, 1);
        Header32(48, 1);
        Header32(56, 4096);
        Header32(60, 4);
        Header32(64, 1);
        Header32(68, End);
        bytes.AsSpan(76, SectorSize - 76).Fill(0xFF);
        Header32(76, 0);
        bytes.AsSpan(SectorSize, SectorSize).Fill(0xFF);
        void Fat(int index, uint value) => Header32(SectorSize + 4 * index, value);
        Fat(0, 0xFFFFFFFD);
        Fat(1, 2);
        Fat(2, 3);
        Fat(3, End);
        Fat(4, End);
        for (int index = 0; index < rootSectors; index++)
            Fat(5 + index, index == rootSectors - 1 ? End : (uint)(6 + index));

        Span<byte> entries = bytes.AsSpan(2 * SectorSize, 3 * SectorSize);
        WriteEntry(entries, 0, "Root Entry", 5, child: 1, start: 5, length: rootBytes);
        WriteEntry(entries, 1, "PROJECT", 2, left: 2, right: 6, start: (uint)miniStarts[0], length: streams[0].Length);
        WriteEntry(entries, 2, "VBA", 1, child: 4);
        WriteEntry(entries, 3, "dir", 2, start: (uint)miniStarts[1], length: streams[1].Length);
        WriteEntry(entries, 4, streamName, 2, left: 3, right: 5, start: (uint)miniStarts[2], length: streams[2].Length);
        WriteEntry(entries, 5, "_VBA_PROJECT", 2, start: (uint)miniStarts[3], length: streams[3].Length);
        WriteEntry(entries, 6, streamName, 1, child: 7);
        WriteEntry(entries, 7, "f", 2, left: 9, right: 8, start: (uint)miniStarts[4], length: streams[4].Length);
        WriteEntry(entries, 8, "o", 2, start: (uint)miniStarts[5], length: streams[5].Length);
        WriteEntry(entries, 9, "\u0003VBFrame", 2, start: (uint)miniStarts[6], length: streams[6].Length);
        bytes.AsSpan(5 * SectorSize, SectorSize).Fill(0xFF);
        for (int i = 0; i < streams.Length; i++)
        {
            int start = miniStarts[i];
            int end = i + 1 == streams.Length ? miniCount : miniStarts[i + 1];
            for (int mini = start; mini < end; mini++)
                Header32(5 * SectorSize + 4 * mini, mini == end - 1 ? End : (uint)(mini + 1));
            streams[i].CopyTo(bytes, 6 * SectorSize + 64 * start);
        }
        return bytes;
    }

    private static void WriteEntry(Span<byte> directory, int index, string name, byte type,
        int left = -1, int right = -1, int child = -1, uint start = Free, int length = 0)
    {
        Span<byte> entry = directory.Slice(128 * index, 128);
        int nameLength = Encoding.Unicode.GetBytes(name, entry) + 2;
        BinaryPrimitives.WriteUInt16LittleEndian(entry[64..], (ushort)nameLength);
        entry[66] = type;
        entry[67] = 1;
        BinaryPrimitives.WriteInt32LittleEndian(entry[68..], left);
        BinaryPrimitives.WriteInt32LittleEndian(entry[72..], right);
        BinaryPrimitives.WriteInt32LittleEndian(entry[76..], child);
        BinaryPrimitives.WriteUInt32LittleEndian(entry[116..], start);
        BinaryPrimitives.WriteInt32LittleEndian(entry[120..], length);
    }

    private static byte[] DirectoryRecords(int sourceOffset, string reference, string streamName, ushort codePage, bool includeConstants)
    {
        using var content = new MemoryStream();
        using var writer = new BinaryWriter(content);
        void Write(ushort id, byte[] value)
        {
            writer.Write(id);
            writer.Write(value.Length);
            writer.Write(value);
        }
        Write(0x0001, [3, 0, 0, 0]);
        Write(0x0002, [9, 4, 0, 0]);
        Write(0x0014, [9, 4, 0, 0]);
        Write(0x0003, BitConverter.GetBytes(codePage));
        Write(0x0004, "FixtureProject"u8.ToArray());
        Write(0x0005, []);
        Write(0x0040, []);
        Write(0x0006, []);
        Write(0x003D, []);
        Write(0x0007, [0, 0, 0, 0]);
        Write(0x0008, [0, 0, 0, 0]);
        Write(0x0009, [0, 0, 0, 0]);
        writer.Write((ushort)1);
        if (includeConstants)
        {
            Write(0x000C, []);
            Write(0x003C, []);
        }
        byte[] referenceBytes = Encoding.ASCII.GetBytes(reference);
        writer.Write((ushort)0x000D);
        writer.Write(referenceBytes.Length + 10);
        writer.Write(referenceBytes.Length);
        writer.Write(referenceBytes);
        writer.Write(new byte[6]);
        Write(0x000F, [1, 0]);
        Write(0x0013, [0, 0]);
        Write(0x0019, "FixtureForm"u8.ToArray());
        Write(0x0047, Encoding.Unicode.GetBytes("FixtureForm"));
        Write(0x001A, Encoding.ASCII.GetBytes(streamName));
        Write(0x0032, Encoding.Unicode.GetBytes(streamName));
        Write(0x001C, []);
        Write(0x0048, []);
        Write(0x0031, BitConverter.GetBytes(sourceOffset));
        Write(0x001E, [0, 0, 0, 0]);
        Write(0x002C, [0, 0]);
        Write(0x0022, []);
        Write(0x002B, []);
        Write(0x0010, []);
        return Compressed(content.ToArray());
    }

    private static byte[] Compressed(byte[] expanded)
    {
        using var content = new MemoryStream();
        using var writer = new BinaryWriter(content);
        writer.Write((byte)1);
        int payloadLength = expanded.Length + (expanded.Length + 7) / 8;
        writer.Write((ushort)(0xB000 | (payloadLength - 1)));
        for (int position = 0; position < expanded.Length; position += 8)
        {
            writer.Write((byte)0);
            writer.Write(expanded, position, Math.Min(8, expanded.Length - position));
        }
        return content.ToArray();
    }

    private string Workbook(byte[]? project, bool external = false, bool relationship = true)
    {
        string path = Path.Combine(Path.GetTempPath(), $"ExcelMcp_VbaPackage_{Guid.NewGuid():N}.xlsm");
        _fixtures.Add(path);
        using var file = new FileStream(path, FileMode.CreateNew);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create);
        void Add(string part, string payload)
        {
            using var entry = archive.CreateEntry(part).Open();
            entry.Write(Encoding.UTF8.GetBytes(payload));
        }

        Add("[Content_Types].xml", $"""
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
              <Override PartName="/{ProjectPart}" ContentType="application/vnd.ms-office.vbaProject"/>
            </Types>
            """);
        Add("_rels/.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
            </Relationships>
            """);
        Add("xl/workbook.xml", "<workbook/>");
        if (project is not null)
        {
            if (relationship)
                Add("xl/_rels/workbook.xml.rels", $"""
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="vba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject"
                        Target="macros/fixture.bin" TargetMode="{(external ? "External" : "Internal")}"/>
                    </Relationships>
                    """);
            using var entry = archive.CreateEntry(ProjectPart).Open();
            entry.Write(project);
        }
        return path;
    }

    private static byte[] ReadVbaPart(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        using var stream = archive.GetEntry(ProjectPart)!.Open();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    public void Dispose()
    {
        foreach (string fixture in _fixtures)
            File.Delete(fixture);
    }
}
