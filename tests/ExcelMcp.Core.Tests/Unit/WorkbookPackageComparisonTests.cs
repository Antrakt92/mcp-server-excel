using System.Buffers.Binary;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.ServiceClient;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "WorkbookPackageComparison")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class WorkbookPackageComparisonTests : IDisposable
{
    private readonly List<string> _fixtures = [];

    [Fact]
    public void Compare_ReturnsExactSortedChangesAndAllowlistStatusWithoutWriting()
    {
        var before = Workbook(".xlsm", ("xl/worksheets/sheet1.xml", "<old/>"), ("xl/vbaProject.bin", "same"),
            ("xl/removed.bin", "old binary"));
        var after = Workbook(".xlsm", ("xl/vbaProject.bin", "same"), ("xl/worksheets/sheet1.xml", "<new/>"),
            ("xl/added.bin", "new binary"));
        byte[] beforeBytes = File.ReadAllBytes(before);
        byte[] afterBytes = File.ReadAllBytes(after);

        var result = WorkbookPackageComparison.Compare(before, after, ["xl/added.bin", "XL/removed.bin"], default);

        Assert.True(result.Success);
        Assert.False(result.MatchesAllowlist);
        Assert.Equal(6, result.BeforePartCount);
        Assert.Equal(6, result.AfterPartCount);
        Assert.Collection(result.Changes,
            part =>
            {
                Assert.Equal("xl/added.bin", part.PartName);
                Assert.Equal("added", part.ChangeKind);
                Assert.Null(part.BeforeSha256);
                Assert.Equal(Hash("new binary"), part.AfterSha256);
                Assert.True(part.Allowed);
            },
            part =>
            {
                Assert.Equal("xl/removed.bin", part.PartName);
                Assert.Equal("removed", part.ChangeKind);
                Assert.Equal(Hash("old binary"), part.BeforeSha256);
                Assert.Null(part.AfterSha256);
                Assert.False(part.Allowed);
            },
            part =>
            {
                Assert.Equal("xl/worksheets/sheet1.xml", part.PartName);
                Assert.Equal("changed", part.ChangeKind);
                Assert.Equal(Hash("<old/>"), part.BeforeSha256);
                Assert.Equal(Hash("<new/>"), part.AfterSha256);
                Assert.False(part.Allowed);
            });
        Assert.Equal(beforeBytes, File.ReadAllBytes(before));
        Assert.Equal(afterBytes, File.ReadAllBytes(after));

        var allowed = WorkbookPackageComparison.Compare(before, after,
            ["xl/added.bin", "xl/removed.bin", "xl/worksheets/sheet1.xml"], default);
        Assert.True(allowed.MatchesAllowlist);
        Assert.All(allowed.Changes, part => Assert.True(part.Allowed));
    }

    [Fact]
    public void Compare_TransportRoundTripKeepsExplicitNullHashesForAbsentParts()
    {
        var before = Workbook(".xlsx", ("xl/removed.bin", "old"));
        var after = Workbook(".xlsx", ("xl/added.bin", "new"));
        var result = WorkbookPackageComparison.Compare(before, after, null, default);

        string json = JsonSerializer.Serialize(result, ServiceProtocol.JsonOptions);
        using var document = JsonDocument.Parse(json);
        var changes = document.RootElement.GetProperty("changes");
        Assert.Equal(JsonValueKind.Null, changes[0].GetProperty("beforeSha256").ValueKind);
        Assert.Equal(JsonValueKind.Null, changes[1].GetProperty("afterSha256").ValueKind);

        var readback = Assert.IsType<WorkbookPackageComparisonResult>(
            JsonSerializer.Deserialize<WorkbookPackageComparisonResult>(json, ServiceProtocol.JsonOptions));
        Assert.True(readback.Success);
        Assert.False(readback.MatchesAllowlist);
        Assert.Equal(Hash("new"), readback.Changes[0].AfterSha256);
        Assert.Equal(Hash("old"), readback.Changes[1].BeforeSha256);
    }

    [Fact]
    public void Compare_PinnedStreamsReadFromStartWithoutReopeningOrClosingHandles()
    {
        string before = Workbook(".xlsx", ("xl/marker.bin", "old"));
        string after = Workbook(".xlsx", ("xl/marker.bin", "new"));
        using var beforeStream = new FileStream(before, FileMode.Open, FileAccess.Read, FileShare.None);
        using var afterStream = new FileStream(after, FileMode.Open, FileAccess.Read, FileShare.None);
        beforeStream.Position = 3;
        afterStream.Position = 7;

        var comparison = WorkbookPackageComparison.Compare(beforeStream, afterStream, before, after,
            ["xl/marker.bin"], default);

        Assert.True(comparison.MatchesAllowlist);
        var change = Assert.Single(comparison.Changes);
        Assert.Equal("xl/marker.bin", change.PartName);
        Assert.Equal(Hash("old"), change.BeforeSha256);
        Assert.Equal(Hash("new"), change.AfterSha256);

        Assert.Throws<InvalidDataException>(() => WorkbookPackageComparison.Compare(
            beforeStream, afterStream, before, Path.ChangeExtension(after, ".xlsm"), null, default));
        beforeStream.Position = 0;
        afterStream.Position = 0;
        Assert.Equal(0x50, beforeStream.ReadByte());
        Assert.Equal(0x50, afterStream.ReadByte());
    }

    [Fact]
    public void Compare_UsesPayloadBytesAndTreatsCaseOnlyRenameAsTwoChanges()
    {
        var before = Workbook(".xlsx", ("xl/marker.bin", "marker payload"));
        var repacked = Workbook(".xlsx", CompressionLevel.NoCompression, ("xl/marker.bin", "marker payload"));
        var renamed = Workbook(".xlsx", ("xl/Marker.bin", "marker payload"));

        var same = WorkbookPackageComparison.Compare(before, repacked, null, default);
        Assert.True(same.MatchesAllowlist);
        Assert.Empty(same.Changes);

        var result = WorkbookPackageComparison.Compare(before, renamed, ["xl/marker.bin"], default);
        Assert.False(result.MatchesAllowlist);
        Assert.Collection(result.Changes,
            part =>
            {
                Assert.Equal(("xl/Marker.bin", "added", false), (part.PartName, part.ChangeKind, part.Allowed));
                Assert.Equal(Hash("marker payload"), part.AfterSha256);
            },
            part =>
            {
                Assert.Equal(("xl/marker.bin", "removed", true), (part.PartName, part.ChangeKind, part.Allowed));
                Assert.Equal(Hash("marker payload"), part.BeforeSha256);
            });
    }

    [Theory]
    [InlineData(22)]
    [InlineData(27)]
    public void Compare_IgnoresFooterSignatureInsideZipComment(ushort commentLength)
    {
        string before = Workbook(".xlsx", ("xl/marker.bin", "unchanged"));
        string after = Workbook(".xlsx", ("xl/marker.bin", "unchanged"));
        byte[] bytes = File.ReadAllBytes(after);
        byte[] commented = new byte[bytes.Length + commentLength];
        bytes.CopyTo(commented, 0);
        BinaryPrimitives.WriteUInt16LittleEndian(commented.AsSpan(bytes.Length - sizeof(ushort)), commentLength);
        BinaryPrimitives.WriteUInt32LittleEndian(commented.AsSpan(bytes.Length), 0x06054b50);
        File.WriteAllBytes(after, commented);

        var result = WorkbookPackageComparison.Compare(before, after, null, default);
        Assert.True(result.Success);
        Assert.True(result.MatchesAllowlist);
        Assert.Equal(4, result.BeforePartCount);
        Assert.Equal(4, result.AfterPartCount);
        Assert.Empty(result.Changes);
    }

    [Fact]
    public void Compare_IgnoresPlausibleFooterPointingToInvalidDirectoryInsideZipComment()
    {
        string before = Workbook(".xlsx", ("xl/marker.bin", "unchanged"));
        string after = Workbook(".xlsx", ("xl/marker.bin", "unchanged"));
        byte[] bytes = File.ReadAllBytes(after);
        const int falseDirectoryLength = 46;
        const int footerLength = 22;
        byte[] commented = new byte[bytes.Length + falseDirectoryLength + footerLength];
        bytes.CopyTo(commented, 0);
        BinaryPrimitives.WriteUInt16LittleEndian(commented.AsSpan(bytes.Length - sizeof(ushort)),
            falseDirectoryLength + footerLength);
        var falseFooter = commented.AsSpan(bytes.Length + falseDirectoryLength);
        BinaryPrimitives.WriteUInt32LittleEndian(falseFooter, 0x06054b50);
        BinaryPrimitives.WriteUInt16LittleEndian(falseFooter[8..], 1);
        BinaryPrimitives.WriteUInt16LittleEndian(falseFooter[10..], 1);
        BinaryPrimitives.WriteUInt32LittleEndian(falseFooter[12..], falseDirectoryLength);
        BinaryPrimitives.WriteUInt32LittleEndian(falseFooter[16..], (uint)bytes.Length);
        File.WriteAllBytes(after, commented);

        var result = WorkbookPackageComparison.Compare(before, after, null, default);
        Assert.True(result.Success);
        Assert.True(result.MatchesAllowlist);
        Assert.Equal(result.BeforePartCount, result.AfterPartCount);
        Assert.Empty(result.Changes);
    }

    [Fact]
    public void Compare_RejectsTwoStructurallyValidZipDirectories()
    {
        string workbook = Workbook(".xlsx", ("xl/marker.bin", "unchanged"));
        byte[] bytes = File.ReadAllBytes(workbook);
        const int footerLength = 22;
        int footerOffset = bytes.Length - footerLength;
        int originalDirectoryStart = (int)BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(footerOffset + 16));
        var secondDirectory = bytes.AsSpan(originalDirectoryStart);
        byte[] ambiguous = new byte[bytes.Length + secondDirectory.Length];
        bytes.CopyTo(ambiguous, 0);
        secondDirectory.CopyTo(ambiguous.AsSpan(bytes.Length));
        BinaryPrimitives.WriteUInt16LittleEndian(ambiguous.AsSpan(footerOffset + 20), (ushort)secondDirectory.Length);
        BinaryPrimitives.WriteUInt32LittleEndian(ambiguous.AsSpan(ambiguous.Length - 6), (uint)bytes.Length);
        File.WriteAllBytes(workbook, ambiguous);

        var error = Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(workbook, workbook, null, default));
        Assert.Contains("ambiguous ZIP directories", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Compare_PreservesDistinctLiteralAndEscapedReservedPartNames()
    {
        string before = Workbook(".xlsx", ("xl/custom+part.xml", "same"),
            ("xl/custom%2Bpart.xml", "old"), ("xl/50%25.xml", "same"));
        string after = Workbook(".xlsx", ("xl/custom+part.xml", "same"),
            ("xl/custom%2Bpart.xml", "new"), ("xl/50%25.xml", "same"));

        var result = WorkbookPackageComparison.Compare(before, after, ["xl/custom+part.xml"], default);

        Assert.False(result.MatchesAllowlist);
        var change = Assert.Single(result.Changes);
        Assert.Equal("xl/custom%2Bpart.xml", change.PartName);
        Assert.Equal(Hash("old"), change.BeforeSha256);
        Assert.Equal(Hash("new"), change.AfterSha256);
        Assert.False(change.Allowed);
    }

    [Theory]
    [InlineData("xl/WORKBOOK.xml")]
    [InlineData("xl/workbook.xml")]
    [InlineData("xl/%77orkbook.xml")]
    [InlineData("xl/a%2fb.xml")]
    [InlineData("xl//a.xml")]
    [InlineData("xl/../a.xml")]
    [InlineData("xl/a./b.xml")]
    [InlineData("xl\\a.xml")]
    [InlineData("xl")]
    public void Compare_RejectsAmbiguousOrInvalidPartNames(string name)
    {
        var before = Workbook(".xlsx");
        var after = Workbook(".xlsx", (name, "ambiguous"));

        var error = Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(before, after, null, default));
        Assert.Contains("part name", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Compare_RejectsMissingRootRelationshipAndWrongWorkbookContentType()
    {
        var before = Workbook(".xlsx");
        var plainZip = Archive(".xlsx", CompressionLevel.Optimal,
            [("[Content_Types].xml", ContentTypes(".xlsx")), ("xl/workbook.xml", "<workbook/>")]);
        var conflictingExtension = Archive(".xlsx", CompressionLevel.Optimal, WorkbookEntries(".xlsm"));

        Assert.Contains("required package part", Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(before, plainZip, null, default)).Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("content type", Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(before, conflictingExtension, null, default)).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("//xl/workbook.xml")]
    [InlineData("/[Content_Types].xml")]
    [InlineData("/_rels/.rels")]
    [InlineData("mailto:book.xml")]
    public void Compare_RefusesNonPartRootRelationshipTargets(string target)
    {
        var before = Workbook(".xlsx");
        string altered = Relationships.Replace("Target=\"xl/workbook.xml\"", $"Target=\"{target}\"", StringComparison.Ordinal);
        var after = Archive(".xlsx", CompressionLevel.Optimal,
            [("[Content_Types].xml", ContentTypes(".xlsx")), ("_rels/.rels", altered), ("xl/workbook.xml", "<workbook/>"),
                ("mailto:book.xml", "<workbook/>")]);

        Assert.Contains("root officeDocument relationship", Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(before, after, null, default)).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(@"\\.\pipe\fixture.xlsx")]
    [InlineData(@"\\localhost\Pipe\fixture.xlsx")]
    [InlineData(@"\\localhost\MAILSLOT\fixture.xlsx")]
    [InlineData(@"C:\fixture.xlsm:other.xlsx")]
    [InlineData(@"C:\work\NUL.xlsx")]
    [InlineData(@"C:\work\CON .xlsx")]
    [InlineData(@"C:\work\COM¹.xlsm")]
    [InlineData(@"C:\work\lpt3.package.xlsx")]
    public void Compare_RefusesDeviceOrAlternateStreamPathsBeforeOpening(string path)
    {
        var workbook = Workbook(".xlsx");

        Assert.Contains("regular file paths", Assert.Throws<ArgumentException>(() =>
            WorkbookPackageComparison.Compare(workbook, path, null, default)).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Compare_RejectsDtdAndCorruptedEntry()
    {
        var before = Workbook(".xlsx");
        var withDtd = Archive(".xlsx", CompressionLevel.Optimal,
            [("[Content_Types].xml", "<!DOCTYPE Types []>" + ContentTypes(".xlsx")),
                ("_rels/.rels", Relationships), ("xl/workbook.xml", "<workbook/>")]);
        var corrupt = Workbook(".xlsx", CompressionLevel.NoCompression, ("xl/unique.bin", "unique payload"));
        var bytes = File.ReadAllBytes(corrupt);
        int offset = bytes.AsSpan().IndexOf("unique payload"u8);
        Assert.True(offset >= 0);
        bytes[offset] ^= 1;
        File.WriteAllBytes(corrupt, bytes);

        Assert.Contains("metadata XML", Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(before, withDtd, null, default)).Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("CRC", Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(before, corrupt, null, default)).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Compare_RejectsDeflatedDataHiddenByForgedPrefixSizeAndCrc()
    {
        const string part = "xl/marker.bin";
        string before = Workbook(".xlsx", (part, "a"));
        string forged = Workbook(".xlsx", (part, "a" + new string('b', 256)));
        byte[] bytes = File.ReadAllBytes(forged);
        int central = FindCentralEntry(bytes, part);
        int local = (int)BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(central + 42));
        Assert.Equal(8, BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(central + 10)));
        const uint crcForA = 0xe8b7be43;
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(central + 16), crcForA);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(central + 24), 1);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(local + 14), crcForA);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(local + 22), 1);
        File.WriteAllBytes(forged, bytes);

        Assert.Contains("exceeds its declared decompressed size", Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(before, forged, null, default)).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Compare_RejectsDeflatedDataMissingFinalBlockEvenWithMatchingPayloadSizeAndCrc()
    {
        const string part = "xl/marker.bin";
        string before = Workbook(".xlsx", (part, "small payload"));
        string forged = Workbook(".xlsx", (part, "small payload"));
        byte[] bytes = File.ReadAllBytes(forged);
        int central = FindCentralEntry(bytes, part);
        int local = (int)BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(central + 42));
        int localNameLength = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(local + 26));
        int extraLength = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(local + 28));
        int payload = local + 30 + localNameLength + extraLength;
        Assert.Equal(8, BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(central + 10)));
        Assert.Equal(1, bytes[payload] & 1);
        bytes[payload] &= 0xfe;
        File.WriteAllBytes(forged, bytes);

        Assert.Contains("complete DEFLATE stream", Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(before, forged, null, default)).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Compare_RejectsTwoCentralEntriesAliasingOnePhysicalRecord()
    {
        const string first = "xl/marka.bin";
        const string second = "xl/markb.bin";
        string before = Workbook(".xlsx");
        string forged = Workbook(".xlsx", (first, "same"), (second, "same"));
        byte[] bytes = File.ReadAllBytes(forged);
        int firstCentral = FindCentralEntry(bytes, first);
        int secondCentral = FindCentralEntry(bytes, second);
        bytes.AsSpan(firstCentral + 42, 4).CopyTo(bytes.AsSpan(secondCentral + 42, 4));
        File.WriteAllBytes(forged, bytes);

        Assert.Contains("overlapping", Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(before, forged, null, default)).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Compare_RejectsConflictingLocalAndCentralPartNames()
    {
        const string part = "xl/marker.bin";
        string before = Workbook(".xlsx", (part, "same"));
        string forged = Workbook(".xlsx", (part, "same"));
        byte[] bytes = File.ReadAllBytes(forged);
        int central = FindCentralEntry(bytes, part);
        int local = (int)BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(central + 42));
        bytes[local + 30] = (byte)'z';
        File.WriteAllBytes(forged, bytes);

        Assert.Contains("local part name", Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(before, forged, null, default)).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Compare_RejectsMoreChangesThanCanBeReportedAndObservesCancellation()
    {
        var before = Workbook(".xlsx");
        var after = Workbook(".xlsx", Enumerable.Range(0, 1025)
            .Select(index => ($"xl/extra{index}.bin", "added")).ToArray());
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Contains("too many changes", Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(before, after, null, default)).Message, StringComparison.OrdinalIgnoreCase);
        Assert.Throws<OperationCanceledException>(() =>
            WorkbookPackageComparison.Compare(before, after, null, cancellation.Token));
    }

    private const string Relationships = """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """;

    private static string ContentTypes(string extension)
    {
        string workbookType = extension == ".xlsm"
            ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
            : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        return $"""
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Override PartName="/xl/workbook.xml" ContentType="{workbookType}"/>
            </Types>
            """;
    }

    private static (string Name, string Content)[] WorkbookEntries(string extension) =>
        [("[Content_Types].xml", ContentTypes(extension)), ("_rels/.rels", Relationships),
            ("xl/workbook.xml", "<workbook/>")];

    private string Workbook(string extension, params (string Name, string Content)[] entries) =>
        Workbook(extension, CompressionLevel.Optimal, entries);

    private string Workbook(string extension, CompressionLevel compression, params (string Name, string Content)[] entries) =>
        Archive(extension, compression, WorkbookEntries(extension).Concat(entries));

    private string Archive(string extension, CompressionLevel compression, IEnumerable<(string Name, string Content)> entries)
    {
        string path = Path.Combine(Path.GetTempPath(), $"ExcelMcp_Package_{Guid.NewGuid():N}{extension}");
        _fixtures.Add(path);
        using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create);
        foreach (var (name, content) in entries)
        {
            using var value = archive.CreateEntry(name, compression).Open();
            value.Write(Encoding.UTF8.GetBytes(content));
        }
        return path;
    }

    private static string Hash(string content) => Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(content)));

    private static int FindCentralEntry(byte[] bytes, string part)
    {
        // These test fixtures are ZIP32 archives with an empty file comment.
        int offset = (int)BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(bytes.Length - 6));
        byte[] name = Encoding.ASCII.GetBytes(part);
        while (BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset)) == 0x02014b50)
        {
            int nameLength = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset + 28));
            if (bytes.AsSpan(offset + 46, nameLength).SequenceEqual(name))
                return offset;
            int extraLength = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset + 30));
            int commentLength = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset + 32));
            offset += 46 + nameLength + extraLength + commentLength;
        }
        throw new InvalidOperationException($"The test archive has no central entry for {part}.");
    }

    public void Dispose()
    {
        foreach (var fixture in _fixtures)
            File.Delete(fixture);
    }
}
