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
public sealed class WorkbookPackageStructureTests : IDisposable
{
    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string StrictSpreadsheetNamespace = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    private const string DocumentRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private const string StrictDocumentRelationships = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
    private const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    private readonly List<string> _fixtures = [];

    [Fact]
    public void CompareStructure_MovedSheetsRetainIdentityAcrossPartRenumberingAndScopedNameChanges()
    {
        SheetFixture[] beforeSheets =
        [
            new("Data", 17, "xl/worksheets/sheet1.xml", "SheetData"),
            new("Chart", 23, "xl/chartsheets/sheet1.xml", "ChartData", "chartsheet")
        ];
        SheetFixture[] afterSheets =
        [
            new("Chart", 23, "xl/chartsheets/sheet7.xml", "ChartData", "chartsheet"),
            new("Data", 17, "xl/worksheets/sheet2.xml", "SheetData")
        ];
        string before = Workbook(beforeSheets,
            names: "<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"0\">Data!$A$1:$B$4</definedName>",
            extraParts:
            [
                ("xl/worksheets/_rels/sheet1.xml.rels", Relationships(
                    Relationship("comments", "comments", "../comments1.xml") +
                    Relationship("vml", "vmlDrawing", "../drawings/vmlDrawing1.vml"))),
                ("xl/comments1.xml", "<comments/>"),
                ("xl/drawings/vmlDrawing1.vml", "<xml/>"),
                ("xl/vbaProject.bin", "synthetic-vba-marker")
            ], macro: true);
        string after = Workbook(afterSheets,
            names: "<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"1\">Data!$A$1:$B$4</definedName>",
            extraParts:
            [
                ("xl/worksheets/_rels/sheet2.xml.rels", Relationships(
                    Relationship("comments", "comments", "../comments4.xml") +
                    Relationship("vml", "vmlDrawing", "../drawings/vmlDrawing4.vml"))),
                ("xl/comments4.xml", "<comments/>"),
                ("xl/drawings/vmlDrawing4.vml", "<xml/>"),
                ("xl/vbaProject.bin", "synthetic-vba-marker")
            ], macro: true);
        byte[] beforeBytes = File.ReadAllBytes(before);
        byte[] afterBytes = File.ReadAllBytes(after);

        var result = WorkbookPackageComparison.Compare(before, after, null, default, includeStructure: true);
        var structure = Assert.IsType<WorkbookPackageStructureReport>(result.Structure);

        Assert.True(result.Success);
        Assert.False(result.MatchesAllowlist);
        Assert.Equal("xl/workbook.xml", structure.Before.WorkbookPart);
        Assert.Equal("xl/workbook.xml", structure.After.WorkbookPart);
        Assert.Collection(structure.Before.Sheets,
            sheet => AssertSheet(sheet, 0, "Data", 17, "sheet1", "xl/worksheets/sheet1.xml", "SheetData"),
            sheet => AssertSheet(sheet, 1, "Chart", 23, "sheet2", "xl/chartsheets/sheet1.xml", "ChartData"));
        Assert.Collection(structure.After.Sheets,
            sheet => AssertSheet(sheet, 0, "Chart", 23, "sheet1", "xl/chartsheets/sheet7.xml", "ChartData"),
            sheet => AssertSheet(sheet, 1, "Data", 17, "sheet2", "xl/worksheets/sheet2.xml", "SheetData"));
        var beforeFilter = Assert.Single(structure.Before.DefinedNames);
        Assert.Equal("_xlnm._FilterDatabase", beforeFilter.Name);
        Assert.Equal(0, beforeFilter.LocalSheetId);
        Assert.Equal(17u, beforeFilter.SheetId);
        Assert.Equal("xl/worksheets/sheet1.xml", beforeFilter.SheetPart);
        var afterFilter = Assert.Single(structure.After.DefinedNames);
        Assert.Equal(1, afterFilter.LocalSheetId);
        Assert.Equal(17u, afterFilter.SheetId);
        Assert.Equal("xl/worksheets/sheet2.xml", afterFilter.SheetPart);
        var comment = Assert.Single(structure.After.Relationships, relationship => relationship.Id == "comments");
        Assert.Equal("xl/worksheets/sheet2.xml", comment.OwnerPart);
        Assert.Equal(DocumentRelationships + "comments", comment.Type);
        Assert.Equal("xl/comments4.xml", comment.TargetPart);
        Assert.True(comment.OwnerExists);
        Assert.True(comment.TargetExists);
        Assert.False(comment.External);
        var vml = Assert.Single(structure.After.Relationships, relationship => relationship.Id == "vml");
        Assert.Equal("xl/drawings/vmlDrawing4.vml", vml.TargetPart);
        Assert.DoesNotContain(vml.TargetPart!, structure.After.UnreachableParts);
        Assert.DoesNotContain(comment.TargetPart!, structure.After.UnreachableParts);
        Assert.Contains(result.Changes, change => change.PartName == "xl/worksheets/sheet1.xml" && change.ChangeKind == "removed");
        Assert.Contains(result.Changes, change => change.PartName == "xl/worksheets/sheet2.xml" && change.ChangeKind == "added");
        Assert.Empty(structure.ParentStyleMappings);
        Assert.Equal(beforeBytes, File.ReadAllBytes(before));
        Assert.Equal(afterBytes, File.ReadAllBytes(after));
    }

    [Fact]
    public void CompareStructure_RemovedHyperlinkStyleDoesNotShiftRemainingNamedStyleIdentity()
    {
        string before = Workbook(styles: Styles(
            "<xf numFmtId=\"0\"/><xf numFmtId=\"0\" applyFont=\"1\"/><xf numFmtId=\"0\" applyAlignment=\"1\"/>",
            "<cellStyle name=\"Normal\" xfId=\"0\"/><cellStyle name=\"Hyperlink\" xfId=\"1\"/><cellStyle name=\"Alert\" xfId=\"2\"/>"));
        string after = Workbook(styles: Styles(
            "<xf numFmtId=\"0\"/><xf numFmtId=\"0\" applyAlignment=\"1\"/>",
            "<cellStyle name=\"Normal\" xfId=\"0\"/><cellStyle name=\"Alert\" xfId=\"1\"/>"));

        var result = WorkbookPackageComparison.Compare(before, after, null, default, includeStructure: true);
        var structure = Assert.IsType<WorkbookPackageStructureReport>(result.Structure);

        Assert.Equal(3, structure.Before.ParentStyles.Count);
        Assert.Equal(2, structure.After.ParentStyles.Count);
        Assert.Collection(structure.ParentStyleMappings,
            mapping => AssertMapping(mapping, "Normal", 0, 0, "matched", [0]),
            mapping => AssertMapping(mapping, "Hyperlink", 1, null, "removed", []),
            mapping => AssertMapping(mapping, "Alert", 2, 1, "matched", [1]));
        var beforeAlert = Assert.Single(structure.Before.ParentStyles, style => style.Name == "Alert");
        var afterAlert = Assert.Single(structure.After.ParentStyles, style => style.Name == "Alert");
        Assert.Equal(2, beforeAlert.XfId);
        Assert.Equal(1, afterAlert.XfId);
        Assert.Equal(beforeAlert.DefinitionSha256, afterAlert.DefinitionSha256);
        Assert.Matches("^[0-9a-f]{64}$", beforeAlert.DefinitionSha256);
        Assert.False(result.MatchesAllowlist);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CompareStructure_DuplicateStyleMatchesAreExplicitlyAmbiguous(bool duplicateBefore)
    {
        const string uniqueNames = "<cellStyle name=\"Normal\" xfId=\"0\"/>";
        const string duplicateNames = "<cellStyle name=\"Normal\" xfId=\"0\"/><cellStyle name=\"Normal\" xfId=\"1\"/>";
        string before = Workbook(styles: Styles("<xf/><xf/>", duplicateBefore ? duplicateNames : uniqueNames));
        string after = Workbook(styles: Styles("<xf/><xf/>", duplicateBefore ? uniqueNames : duplicateNames));

        var structure = WorkbookPackageComparison.Compare(before, after, null, default, includeStructure: true).Structure!;

        Assert.Equal(duplicateBefore ? 2 : 1, structure.ParentStyleMappings.Count);
        int[] expectedCandidates = duplicateBefore ? [0] : [0, 1];
        Assert.All(structure.ParentStyleMappings, mapping =>
        {
            Assert.Equal("Normal", mapping.Name);
            Assert.Equal("ambiguous", mapping.Status);
            Assert.Null(mapping.AfterXfId);
            Assert.Equal(expectedCandidates, mapping.CandidateAfterXfIds);
        });
    }

    [Fact]
    public void CompareStructure_SameNameWithDifferentDefinitionsMatchesOnlyTheExactDefinition()
    {
        const string matching = "<xf applyAlignment=\"1\"/>";
        const string changed = "<xf applyFont=\"1\"/>";
        const string other = "<xf applyBorder=\"1\"/>";
        const string names = "<cellStyle name=\"Repeated\" xfId=\"0\"/><cellStyle name=\"Repeated\" xfId=\"1\"/>";
        string before = Workbook(styles: Styles(matching + changed, names));
        string after = Workbook(styles: Styles(other + matching, names));

        var structure = WorkbookPackageComparison.Compare(before, after, null, default, includeStructure: true).Structure!;

        Assert.Collection(structure.ParentStyleMappings,
            mapping => AssertMapping(mapping, "Repeated", 0, 1, "matched", [1]),
            mapping => AssertMapping(mapping, "Repeated", 1, null, "definition-changed", []));
    }

    [Fact]
    public void CompareStructure_ReferencedFontContentChangesInvalidateOtherwiseIdenticalParentStyle()
    {
        const string parents = "<xf fontId=\"0\"/>";
        const string names = "<cellStyle name=\"Normal\" xfId=\"0\"/>";
        string before = Workbook(styles: Styles(parents, names, "<font><name val=\"Calibri\"/></font>"));
        string after = Workbook(styles: Styles(parents, names, "<font><name val=\"Arial\"/></font>"));

        var structure = WorkbookPackageComparison.Compare(before, after, null, default, includeStructure: true).Structure!;

        Assert.NotEqual(Assert.Single(structure.Before.ParentStyles).DefinitionSha256,
            Assert.Single(structure.After.ParentStyles).DefinitionSha256);
        AssertMapping(Assert.Single(structure.ParentStyleMappings), "Normal", 0, null, "definition-changed", []);
    }

    [Fact]
    public void CompareStructure_RenumberedFontReferencesMatchByResolvedDefinition()
    {
        const string names = "<cellStyle name=\"Normal\" xfId=\"0\"/>";
        string before = Workbook(styles: Styles("<xf fontId=\"0\"/>", names,
            "<font><name val=\"Calibri\"/></font><font><name val=\"Arial\"/></font>"));
        string after = Workbook(styles: Styles("<xf fontId=\"1\"/>", names,
            "<font><name val=\"Arial\"/></font><font><name val=\"Calibri\"/></font>"));

        var structure = WorkbookPackageComparison.Compare(before, after, null, default, includeStructure: true).Structure!;

        Assert.Equal(Assert.Single(structure.Before.ParentStyles).DefinitionSha256,
            Assert.Single(structure.After.ParentStyles).DefinitionSha256);
        AssertMapping(Assert.Single(structure.ParentStyleMappings), "Normal", 0, 0, "matched", [0]);
    }

    [Fact]
    public void CompareStructure_AddedStyleHasNoInventedBeforeIndex()
    {
        string before = Workbook();
        string after = Workbook(styles: Styles("<xf/>", "<cellStyle name=\"NewStyle\" xfId=\"0\"/>"));

        var structure = WorkbookPackageComparison.Compare(before, after, null, default, includeStructure: true).Structure!;

        Assert.Empty(structure.Before.ParentStyles);
        AssertMapping(Assert.Single(structure.ParentStyleMappings), "NewStyle", null, 0, "added", [0]);
    }

    [Fact]
    public void CompareStructure_StrictNamespaceAndWorkbookScopedNamesAreSupportedWithoutReturningValues()
    {
        string workbook = Workbook(strict: true,
            names: "<definedName name=\"GlobalScope\">SYNTHETIC_FORMULA_MARKER</definedName>");

        var result = WorkbookPackageComparison.Compare(workbook, workbook, null, default, includeStructure: true);
        var structure = Assert.IsType<WorkbookPackageStructureReport>(result.Structure);

        Assert.True(result.MatchesAllowlist);
        Assert.Empty(result.Changes);
        AssertSheet(Assert.Single(structure.Before.Sheets), 0, "Data", 9, "sheet1", "xl/worksheets/sheet1.xml", "SheetData");
        var name = Assert.Single(structure.Before.DefinedNames);
        Assert.Equal("GlobalScope", name.Name);
        Assert.Null(name.LocalSheetId);
        Assert.Null(name.SheetId);
        Assert.Null(name.SheetPart);
        Assert.Contains(structure.Before.Relationships, relationship => relationship.Type == StrictDocumentRelationships + "worksheet");
        string json = JsonSerializer.Serialize(result);
        Assert.DoesNotContain("SYNTHETIC_FORMULA_MARKER", json, StringComparison.Ordinal);
        Assert.DoesNotContain("SYNTHETIC_CELL_MARKER", json, StringComparison.Ordinal);
    }

    [Fact]
    public void CompareStructure_GraphCyclesMissingOwnersAndExternalLinksRemainBoundedAndObservable()
    {
        string workbook = Workbook(additionalRelationships:
            Relationship("cycle", "customXml", "cycle.xml") +
            "<Relationship Id=\"external\" Type=\"urn:synthetic:external\" TargetMode=\"External\" Target=\"https://invalid.example/?token=SYNTHETIC_EXTERNAL_MARKER\"/>",
            extraParts:
            [
                ("xl/cycle.xml", "<cycle/>"),
                ("xl/_rels/cycle.xml.rels", Relationships(Relationship("back", "customXml", "workbook.xml"))),
                ("xl/_rels/missing.xml.rels", Relationships(Relationship("orphan", "customXml", "unreachable.xml"))),
                ("xl/unreachable.xml", "<unreachable/>"),
                ("xl/unreferenced.bin", "synthetic-unreferenced")
            ]);

        var result = WorkbookPackageComparison.Compare(workbook, workbook, null, default, includeStructure: true);
        var snapshot = result.Structure!.Before;

        string[] expectedUnreachable = ["xl/unreachable.xml", "xl/unreferenced.bin"];
        Assert.Equal(expectedUnreachable, snapshot.UnreachableParts);
        var external = Assert.Single(snapshot.Relationships, relationship => relationship.Id == "external");
        Assert.Equal("xl/workbook.xml", external.OwnerPart);
        Assert.True(external.OwnerExists);
        Assert.True(external.External);
        Assert.Null(external.TargetPart);
        Assert.Null(external.TargetExists);
        var orphan = Assert.Single(snapshot.Relationships, relationship => relationship.Id == "orphan");
        Assert.Equal("xl/missing.xml", orphan.OwnerPart);
        Assert.False(orphan.OwnerExists);
        Assert.True(orphan.TargetExists);
        Assert.Equal("xl/unreachable.xml", orphan.TargetPart);
        var root = Assert.Single(snapshot.Relationships, relationship => relationship.OwnerPart.Length == 0);
        Assert.True(root.OwnerExists);
        Assert.Equal("xl/workbook.xml", root.TargetPart);
        Assert.DoesNotContain("SYNTHETIC_EXTERNAL_MARKER", JsonSerializer.Serialize(result), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("../../outside.xml")]
    [InlineData("worksheets/%2e%2e/outside.xml")]
    [InlineData("worksheets/a%2fb.xml")]
    [InlineData("https://invalid.example/outside.xml")]
    [InlineData("scheme:part.xml")]
    [InlineData("//invalid.example/outside.xml")]
    [InlineData("worksheets/sheet1.xml#fragment")]
    public void CompareStructure_RejectsAmbiguousOrEscapingInternalTargets(string target)
    {
        string workbook = Workbook(additionalRelationships: Relationship("bad", "customXml", target));

        Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(workbook, workbook, null, default, includeStructure: true));
        Assert.True(WorkbookPackageComparison.Compare(workbook, workbook, null, default).MatchesAllowlist);
    }

    [Theory]
    [InlineData("assets/notes:1.xml")]
    [InlineData("/xl/assets/notes:1.xml")]
    public void CompareStructure_ResolvesLiteralColonInInternalPartName(string target)
    {
        const string part = "xl/assets/notes:1.xml";
        string workbook = Workbook(additionalRelationships: Relationship("notes", "customXml", target),
            extraParts: [(part, "<notes/>")]);

        var snapshot = WorkbookPackageComparison.Compare(workbook, workbook, null, default, includeStructure: true).Structure!.Before;
        var relationship = Assert.Single(snapshot.Relationships, item => item.Id == "notes");
        Assert.False(relationship.External);
        Assert.True(relationship.TargetExists);
        Assert.Equal(part, relationship.TargetPart);
        Assert.DoesNotContain(part, snapshot.UnreachableParts);
    }

    [Fact]
    public void CompareStructure_RejectsDuplicateRelationshipIdsBeforeCreatingDictionary()
    {
        string workbook = Workbook(additionalRelationships: Relationship("sheet1", "customXml", "other.xml"));

        var exception = Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(workbook, workbook, null, default, includeStructure: true));

        Assert.Contains("duplicate", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CompareStructure_ResolvesRelationshipsForWorkbookLocatedAtPackageRoot()
    {
        string workbook = Workbook(extraParts:
        [
            ("[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Override PartName="/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/sheets/data.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                </Types>
                """),
            ("_rels/.rels", Relationships(Relationship("office", "officeDocument", "workbook.xml"))),
            ("workbook.xml", $"<workbook xmlns=\"{SpreadsheetNamespace}\" xmlns:r=\"{DocumentRelationships.TrimEnd('/')}\"><sheets><sheet name=\"RootData\" sheetId=\"31\" r:id=\"rootSheet\"/></sheets></workbook>"),
            ("_rels/workbook.xml.rels", Relationships(Relationship("rootSheet", "worksheet", "sheets/data.xml"))),
            ("sheets/data.xml", $"<worksheet xmlns=\"{SpreadsheetNamespace}\"><sheetPr codeName=\"RootDataCode\"/><sheetData/></worksheet>")
        ]);

        var snapshot = WorkbookPackageComparison.Compare(workbook, workbook, null, default, includeStructure: true).Structure!.Before;

        Assert.Equal("workbook.xml", snapshot.WorkbookPart);
        AssertSheet(Assert.Single(snapshot.Sheets), 0, "RootData", 31, "rootSheet", "sheets/data.xml", "RootDataCode");
        var relationship = Assert.Single(snapshot.Relationships, item => item.Id == "rootSheet");
        Assert.Equal("workbook.xml", relationship.OwnerPart);
        Assert.True(relationship.OwnerExists);
        Assert.Equal("sheets/data.xml", relationship.TargetPart);
        Assert.True(relationship.TargetExists);
        Assert.DoesNotContain("_rels/workbook.xml.rels", snapshot.UnreachableParts);
    }

    [Theory]
    [InlineData("2")]
    [InlineData("-1")]
    public void CompareStructure_RejectsDefinedNameScopesOutsideWorkbookTabPositions(string scope)
    {
        string workbook = Workbook(names: $"<definedName name=\"Scoped\" localSheetId=\"{scope}\">Data!$A$1</definedName>");

        Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(workbook, workbook, null, default, includeStructure: true));
    }

    [Fact]
    public void CompareStructure_DefaultModeLeavesPreviouslyAcceptedOpaqueWorkbookXmlUnparsed()
    {
        string workbook = Workbook(extraParts: [("xl/workbook.xml", "<opaque-workbook/>")]);

        var result = WorkbookPackageComparison.Compare(workbook, workbook, null, default);

        Assert.True(result.Success);
        Assert.True(result.MatchesAllowlist);
        Assert.Null(result.Structure);
        Assert.DoesNotContain("Structure", JsonSerializer.Serialize(result), StringComparison.Ordinal);
        Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(workbook, workbook, null, default, includeStructure: true));
    }

    [Fact]
    public void CompareStructure_RejectsOversizedMetadataWithoutWeakeningDefaultComparison()
    {
        string workbook = Workbook(names: "<definedName name=\"Large\">" + new string('x', 2 * 1024 * 1024) + "</definedName>");

        var exception = Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(workbook, workbook, null, default, includeStructure: true));

        Assert.Contains("limit", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(WorkbookPackageComparison.Compare(workbook, workbook, null, default).MatchesAllowlist);
    }

    [Fact]
    public void CompareStructure_ProhibitsDtdInCapturedWorkbookMetadata()
    {
        string workbook = Workbook(extraParts:
        [
            ("xl/workbook.xml", $"<!DOCTYPE workbook [<!ENTITY example 'SYNTHETIC_ENTITY_MARKER'>]><workbook xmlns=\"{SpreadsheetNamespace}\"><definedNames><definedName name=\"Entity\">&example;</definedName></definedNames></workbook>")
        ]);

        Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(workbook, workbook, null, default, includeStructure: true));
        Assert.True(WorkbookPackageComparison.Compare(workbook, workbook, null, default).MatchesAllowlist);
    }

    [Fact]
    public void CompareStructure_RejectsDeepMetadataBeforeRecursiveFingerprinting()
    {
        string nested = string.Concat(Enumerable.Repeat("<extension>", 70)) +
            string.Concat(Enumerable.Repeat("</extension>", 70));
        string workbook = Workbook(styles: Styles($"<xf>{nested}</xf>", "<cellStyle name=\"Deep\" xfId=\"0\"/>"));

        var exception = Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(workbook, workbook, null, default, includeStructure: true));

        Assert.Contains("depth", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(WorkbookPackageComparison.Compare(workbook, workbook, null, default).MatchesAllowlist);
    }

    [Fact]
    public void CompareStructure_BoundsCanonicalStyleSizeWhenNamespaceExpansionExceedsSourceSize()
    {
        string namespaceUri = "urn:synthetic:" + new string('n', 16 * 1024);
        string children = string.Concat(Enumerable.Repeat("<p:node/>", 520));
        string parent = $"<xf xmlns:p=\"{namespaceUri}\">{children}</xf>";
        string workbook = Workbook(styles: Styles(parent, "<cellStyle name=\"Expanded\" xfId=\"0\"/>"));

        var exception = Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(workbook, workbook, null, default, includeStructure: true));

        Assert.Contains("fingerprint size limit", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(WorkbookPackageComparison.Compare(workbook, workbook, null, default).MatchesAllowlist);
    }

    [Theory]
    [InlineData("<x:metadata/>")]
    [InlineData("<mc:AlternateContent><mc:Choice Requires=\"x\"><sheetPr codeName=\"ChoiceCode\"/></mc:Choice><mc:Fallback/></mc:AlternateContent>")]
    public void CompareStructure_RejectsUnsupportedLeadingSheetMetadataInsteadOfReportingAbsentCodeName(string leadingMetadata)
    {
        string sheet = $"<worksheet xmlns=\"{SpreadsheetNamespace}\" xmlns:x=\"urn:synthetic:extension\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x\">{leadingMetadata}<sheetPr codeName=\"SavedCodeName\"/><sheetData/></worksheet>";
        string workbook = Workbook(extraParts: [("xl/worksheets/sheet1.xml", sheet)]);

        var exception = Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(workbook, workbook, null, default, includeStructure: true));

        Assert.Contains("unsupported", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(WorkbookPackageComparison.Compare(workbook, workbook, null, default).MatchesAllowlist);
    }

    [Fact]
    public void CompareStructure_BoundsCumulativeParentStyleCandidatesBeforeSerializingReport()
    {
        const int styleCount = 65;
        string parents = string.Concat(Enumerable.Repeat("<xf/>", styleCount));
        string names = string.Concat(Enumerable.Range(0, styleCount)
            .Select(index => $"<cellStyle name=\"Repeated\" xfId=\"{index}\"/>"));
        string workbook = Workbook(styles: Styles(parents, names));

        var exception = Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(workbook, workbook, null, default, includeStructure: true));

        Assert.Contains("limit", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(WorkbookPackageComparison.Compare(workbook, workbook, null, default).MatchesAllowlist);
    }

    [Fact]
    public void CompareStructure_CancelledCallPreservesBothSavedInputs()
    {
        string before = Workbook();
        string after = Workbook(names: "<definedName name=\"Global\">42</definedName>");
        byte[] beforeBytes = File.ReadAllBytes(before);
        byte[] afterBytes = File.ReadAllBytes(after);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            WorkbookPackageComparison.Compare(before, after, null, cancellation.Token, includeStructure: true));

        Assert.Equal(beforeBytes, File.ReadAllBytes(before));
        Assert.Equal(afterBytes, File.ReadAllBytes(after));
    }

    private sealed record SheetFixture(string Name, uint SheetId, string PartName, string? CodeName,
        string Kind = "worksheet");

    [Fact]
    public void CompareStructure_RejectsReportWhoseEscapedJsonExceedsOutputLimit()
    {
        // The XML stays below the metadata limit. JSON escapes each plus sign as six bytes.
        string type = "urn:bounded:" + new string('+', 1024);
        string relationships = string.Concat(Enumerable.Range(0, 400).Select(index =>
            $"<Relationship Id=\"external{index}\" Type=\"{type}\" TargetMode=\"External\" Target=\"https://invalid.example\"/>"));
        string workbook = Workbook(additionalRelationships: relationships);

        var error = Assert.Throws<InvalidDataException>(() =>
            WorkbookPackageComparison.Compare(workbook, workbook, null, default, includeStructure: true));
        Assert.Contains("4 MiB output limit", error.Message, StringComparison.Ordinal);
        Assert.True(WorkbookPackageComparison.Compare(workbook, workbook, null, default).MatchesAllowlist);
    }

    [Fact]
    public void CompareStructure_ResolvesEquivalentPartNamesAndKeepsPhysicalSpelling()
    {
        string workbook = Workbook(extraParts:
        [
            ("[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Override PartName="/XL/WORKBOOK.XML" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/XL/WORKSHEETS/SHEET1.XML" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                  <Default Extension="RELS" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                </Types>
                """),
            ("_rels/.rels", Relationships(Relationship("office", "officeDocument", "XL/WORKBOOK.XML"))),
            ("xl/_rels/workbook.xml.rels", Relationships(Relationship("sheet1", "worksheet", "/XL/WorkSheets/Sheet1.xml"))),
            ("XL/WORKSHEETS/_RELS/SHEET1.XML.RELS", Relationships(Relationship("notes", "comments", "../COMMENTS1.XML"))),
            ("xl/Comments1.xml", "<comments/>")
        ]);

        var result = WorkbookPackageComparison.Compare(workbook, workbook, null, default, includeStructure: true);
        var snapshot = result.Structure!.Before;
        Assert.True(result.MatchesAllowlist);
        Assert.Empty(result.Changes);
        Assert.Equal("xl/workbook.xml", snapshot.WorkbookPart);
        AssertSheet(Assert.Single(snapshot.Sheets), 0, "Data", 9, "sheet1", "xl/worksheets/sheet1.xml", "SheetData");
        var notes = Assert.Single(snapshot.Relationships, relationship => relationship.Id == "notes");
        Assert.Equal("xl/worksheets/sheet1.xml", notes.OwnerPart);
        Assert.True(notes.OwnerExists);
        Assert.Equal("xl/Comments1.xml", notes.TargetPart);
        Assert.True(notes.TargetExists);
        Assert.Empty(snapshot.UnreachableParts);
    }

    [Fact]
    public void CompareStructure_DoesNotTreatNestedRelsFolderFilesAsRelationships()
    {
        const string opaquePart = "xl/_rels/custom/notes.xml.rels";
        string workbook = Workbook(extraParts:
        [
            ("xl/custom/notes.xml", "<notes/>"),
            (opaquePart, Relationships(Relationship("phantom", "comments", "fake.xml")))
        ]);

        var snapshot = WorkbookPackageComparison.Compare(workbook, workbook, null, default, includeStructure: true).Structure!.Before;
        Assert.DoesNotContain(snapshot.Relationships, relationship => relationship.Id == "phantom");
        Assert.Contains(opaquePart, snapshot.UnreachableParts);
    }

    private string Workbook(SheetFixture[]? sheets = null, string? names = null, string? styles = null,
        string additionalRelationships = "", bool strict = false,
        (string Part, string Content)[]? extraParts = null, bool macro = false)
    {
        sheets ??= [new("Data", 9, "xl/worksheets/sheet1.xml", "SheetData")];
        string spreadsheet = strict ? StrictSpreadsheetNamespace : SpreadsheetNamespace;
        string relationshipTypes = strict ? StrictDocumentRelationships : DocumentRelationships;
        string sheetElements = string.Concat(sheets.Select((sheet, index) =>
            $"<sheet name=\"{sheet.Name}\" sheetId=\"{sheet.SheetId}\" r:id=\"sheet{index + 1}\"/>"));
        string sheetRelationships = string.Concat(sheets.Select((sheet, index) =>
            Relationship($"sheet{index + 1}", sheet.Kind, "/" + sheet.PartName, strict)));
        if (styles is not null)
            sheetRelationships += Relationship("styles", "styles", "styles.xml", strict);
        string workbookXml = $"<workbook xmlns=\"{spreadsheet}\" xmlns:r=\"{relationshipTypes.TrimEnd('/')}\"><sheets>{sheetElements}</sheets><definedNames>{names}</definedNames></workbook>";
        string mainType = macro
            ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
            : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        string overrides = $"<Override PartName=\"/xl/workbook.xml\" ContentType=\"{mainType}\"/>";
        overrides += string.Concat(sheets.Select(sheet =>
            $"<Override PartName=\"/{sheet.PartName}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.{sheet.Kind}+xml\"/>"));
        if (styles is not null)
            overrides += "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>";

        var parts = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["[Content_Types].xml"] = $"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">{overrides}</Types>",
            ["_rels/.rels"] = Relationships(Relationship("office", "officeDocument", "xl/workbook.xml", strict)),
            ["xl/workbook.xml"] = workbookXml,
            ["xl/_rels/workbook.xml.rels"] = Relationships(sheetRelationships + additionalRelationships)
        };
        foreach (var sheet in sheets)
        {
            string properties = sheet.CodeName is null ? "" : $"<sheetPr codeName=\"{sheet.CodeName}\"/>";
            string cells = sheet.Kind == "worksheet"
                ? "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>SYNTHETIC_CELL_MARKER</t></is></c></row></sheetData>"
                : "";
            parts.Add(sheet.PartName, $"<{sheet.Kind} xmlns=\"{spreadsheet}\">{properties}{cells}</{sheet.Kind}>");
        }
        if (styles is not null)
            parts.Add("xl/styles.xml", styles);
        foreach (var (part, content) in extraParts ?? [])
            parts[part] = content;

        string path = Path.Combine(Path.GetTempPath(), $"ExcelMcp_Structure_{Guid.NewGuid():N}{(macro ? ".xlsm" : ".xlsx")}");
        _fixtures.Add(path);
        using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create);
        foreach (var part in parts)
        {
            using var output = archive.CreateEntry(part.Key, CompressionLevel.Optimal).Open();
            output.Write(Encoding.UTF8.GetBytes(part.Value));
        }
        return path;
    }

    private static string Relationships(string entries) =>
        $"<Relationships xmlns=\"{PackageRelationships}\">{entries}</Relationships>";

    private static string Relationship(string id, string kind, string target, bool strict = false) =>
        $"<Relationship Id=\"{id}\" Type=\"{(strict ? StrictDocumentRelationships : DocumentRelationships)}{kind}\" Target=\"{target}\"/>";

    private static string Styles(string parents, string names, string fonts = "") =>
        $"<styleSheet xmlns=\"{SpreadsheetNamespace}\"><fonts>{fonts}</fonts><cellStyleXfs>{parents}</cellStyleXfs><cellStyles>{names}</cellStyles></styleSheet>";

    private static void AssertSheet(WorkbookPackageSheet sheet, int position, string name, uint id,
        string relationship, string part, string codeName)
    {
        Assert.Equal(position, sheet.Position);
        Assert.Equal(name, sheet.Name);
        Assert.Equal(id, sheet.SheetId);
        Assert.Equal(relationship, sheet.RelationshipId);
        Assert.Equal(part, sheet.PartName);
        Assert.Equal(codeName, sheet.CodeName);
    }

    private static void AssertMapping(WorkbookPackageParentStyleMapping mapping, string name, int? before,
        int? after, string status, int[] candidates)
    {
        Assert.Equal(name, mapping.Name);
        Assert.Equal(before, mapping.BeforeXfId);
        Assert.Equal(after, mapping.AfterXfId);
        Assert.Equal(status, mapping.Status);
        Assert.Equal(candidates, mapping.CandidateAfterXfIds);
    }

    public void Dispose()
    {
        foreach (string fixture in _fixtures)
            File.Delete(fixture);
    }
}
