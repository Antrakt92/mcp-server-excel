using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "VBA")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class VbaModuleSourceTests
{
    [Fact]
    public void Normalize_RemovesExportMetadataAfterBrokenNewlines_AndPreservesCode()
    {
        const string source = "Attribute VB_Name = \"Module1\"\r\r\nOption Explicit\r\r\nPublic Sub Test()\r\r\n  Debug.Print \"Attribute VB_Name = literal\"\r\r\nEnd Sub\r\r\n";
        Assert.Equal("Option Explicit\r\nPublic Sub Test()\r\n  Debug.Print \"Attribute VB_Name = literal\"\r\nEnd Sub", VbaModuleSource.Normalize(source));
    }

    [Theory]
    [InlineData("VERSION 1.0 CLASS\nBEGIN\nEND")]
    [InlineData("Option Explicit\0")]
    [InlineData("Attribute InvalidWhatever = 1\nSub A()\nEnd Sub")]
    [InlineData("Sub A()\nAttribute VB_Name = \"Oops\"\nEnd Sub")]
    public void Normalize_RejectsUnsupportedExportOrEmbeddedMetadataBeforeMutation(string source) =>
        Assert.Throws<ArgumentException>(() => VbaModuleSource.Normalize(source));

    [Fact]
    public void Normalize_DoesNotChangeContinuationOrComment()
    {
        const string code = "' Attribute VB_Name = comment\nPublic Function F( _\n x As String) As String\nF = x\nEnd Function";
        Assert.Equal(code.Replace("\n", "\r\n", StringComparison.Ordinal), VbaModuleSource.Normalize(code));
        Assert.Equal(VbaModuleSource.ComputeHash(code), VbaModuleSource.ComputeHash(code.Replace("\n", "\r\n", StringComparison.Ordinal)));
    }

    [Fact]
    public void Replace_StaleFingerprint_DoesNotDelete()
    {
        var writes = 0;
        Assert.Throws<OperationFailureException>(() => VbaModuleSource.Replace("M", new string('0', 64), "new", () => "old", () => writes++, _ => writes++));
        Assert.Equal(0, writes);
    }

    [Fact]
    public void Replace_InsertFailure_ReportsReadBackAndNeverRollsBack()
    {
        var current = "old";
        var inserts = 0;
        var failure = Assert.Throws<VbaSourceEditException>(() => VbaModuleSource.Replace("M", VbaModuleSource.ComputeHash(current), "new", () => current,
            () => current = "", _ => { inserts++; throw new InvalidOperationException("synthetic insert failure"); }));
        Assert.Equal(1, inserts);
        Assert.Equal("", current);
        Assert.Equal("inserting", failure.State.Phase);
        Assert.Equal(VbaModuleSource.ComputeHash(""), failure.State.SourceHash);
        Assert.Equal(0, failure.State.LineCount);
        Assert.Equal("synthetic insert failure", failure.InnerException!.Message);
    }

    [Fact]
    public void Replace_ChangedReadBackFailsAndPreservesObservedState()
    {
        var current = "old";
        var failure = Assert.Throws<VbaSourceEditException>(() => VbaModuleSource.Replace("M", VbaModuleSource.ComputeHash(current), "new", () => current,
            () => current = "", _ => current = "unexpected"));
        Assert.Equal("verifying", failure.State.Phase);
        Assert.Equal(VbaModuleSource.ComputeHash("unexpected"), failure.State.SourceHash);
        Assert.Equal("unexpected", current);
    }

    [Theory]
    [InlineData("public sub F()\n  Debug.Print \"Literal\" ' Comment\nend sub", true)]
    [InlineData("Public Sub F()\n  Debug.Print \"literal\" ' Comment\nEnd Sub", false)]
    [InlineData("Public Sub F()\n  Debug.Print \"Literal\" ' comment\nEnd Sub", false)]
    [InlineData("Public Sub F()\n  Debug.Print \"Literal\" ' Comment\nEnd Sub", true)]
    public void Replace_AcceptsEditorRecasingButPreservesLiteralsAndComments(string stored, bool accepted)
    {
        const string source = "Public Sub F()\n  Debug.Print \"Literal\" ' Comment\nEnd Sub";
        var current = "old";
        if (accepted)
            Assert.Equal(stored, VbaModuleSource.Replace("M", null, source, () => current, () => current = "", _ => current = stored));
        else
            Assert.Throws<VbaSourceEditException>(() => VbaModuleSource.Replace("M", null, source, () => current, () => current = "", _ => current = stored));
    }

    [Fact]
    public void Replace_PrimaryAndReadBackErrorsRemainAvailable()
    {
        var reads = 0;
        var failure = Assert.Throws<VbaSourceEditException>(() => VbaModuleSource.Replace("M", VbaModuleSource.ComputeHash("old"), "new",
            () => ++reads == 1 ? "old" : throw new InvalidOperationException("readback failure"),
            () => throw new InvalidOperationException("delete failure"), _ => { }));
        Assert.Equal("delete failure", failure.InnerException!.Message);
        Assert.Equal("readback failure", failure.State.ReadBackError);
        Assert.Null(failure.State.SourceHash);
    }

    [Theory]
    [InlineData("10 Rem Keep Exact", "10 Rem keep exact")]
    [InlineData("If ready Then Rem Keep Exact", "If ready Then Rem keep exact")]
    [InlineData("If ready Then x = 1 Else Rem Keep Exact", "If ready Then x = 1 Else Rem keep exact")]
    public void Replace_PreservesNumberedAndConditionalRemComments(string source, string changed)
    {
        var current = "old";
        Assert.Throws<VbaSourceEditException>(() => VbaModuleSource.Replace("M", null, source,
            () => current, () => current = "", _ => current = changed));
    }

    [Fact]
    public void Import_CaseInsensitiveConflict_DoesNotCreateOrRename()
    {
        int writes = 0;
        var failure = Assert.Throws<OperationFailureException>(() => VbaModuleSource.Import("tools", "new", ["Tools"],
            () => writes++, _ => writes++, _ => writes++, () => "Tools", () => "old"));
        Assert.Equal(OperationFailureCategory.Conflict, failure.ErrorCategory);
        Assert.Equal(0, writes);
    }

    [Fact]
    public void Import_RenameFailure_ReportsCreatedActualNameAndEmptySource()
    {
        int creates = 0;
        int inserts = 0;
        var failure = Assert.Throws<VbaSourceEditException>(() => VbaModuleSource.Import("Target", "new", [],
            () => creates++, _ => throw new InvalidOperationException("rename failure"), _ => inserts++,
            () => "Module1", () => string.Empty));
        Assert.Equal(1, creates);
        Assert.Equal(0, inserts);
        Assert.Equal("rename failure", failure.InnerException!.Message);
        Assert.Equal("Module1", failure.State.ModuleName);
        Assert.Equal("naming", failure.State.Phase);
        Assert.Equal(VbaModuleSource.ComputeHash(""), failure.State.SourceHash);
        Assert.Equal(0, failure.State.LineCount);
        Assert.Null(failure.State.ReadBackError);
    }

    [Fact]
    public void Import_PartialInsertionFailure_ReportsActualRemainingSourceWithoutBody()
    {
        string name = "Module1";
        string source = "";
        var failure = Assert.Throws<VbaSourceEditException>(() => VbaModuleSource.Import("Target", "new", [],
            () => { }, value => name = value,
            _ => { source = "partial private source"; throw new InvalidOperationException("insert failure"); },
            () => name, () => source));
        Assert.Equal("Target", failure.State.ModuleName);
        Assert.Equal("inserting", failure.State.Phase);
        Assert.Equal(VbaModuleSource.ComputeHash(source), failure.State.SourceHash);
        Assert.Equal(1, failure.State.LineCount);
        Assert.DoesNotContain(source, failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Import_CreationFailure_PreservesPrimaryAndBothReadBackFailures()
    {
        var failure = Assert.Throws<VbaSourceEditException>(() => VbaModuleSource.Import("Target", "new", [],
            () => throw new InvalidOperationException("create failure"), _ => { }, _ => { },
            () => throw new InvalidOperationException("name unavailable"),
            () => throw new InvalidOperationException("source unavailable")));
        Assert.Equal("create failure", failure.InnerException!.Message);
        Assert.Equal("creating", failure.State.Phase);
        Assert.Equal("", failure.State.ModuleName);
        Assert.Null(failure.State.SourceHash);
        Assert.Null(failure.State.LineCount);
        Assert.Contains("name unavailable", failure.State.ReadBackError);
        Assert.Contains("source unavailable", failure.State.ReadBackError);
    }
}
