using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "VBA")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class VbaProcedureEditTests
{
    private const string Source = "Sub F()\nDebug.Print 1\nEnd Sub";

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void InvalidPhysicalSource_DoesNotDeleteOrInsert(bool nul)
    {
        int writes = 0;
        string invalid = "Sub F()\n" + (nul ? "'bad\0value" : "'" + new string('x', 1023)) + "\nEnd Sub";
        Assert.Throws<ArgumentException>(() => VbaProcedureSource.Replace("M", invalid,
            () => writes++, _ => writes++, () => Source, () => Source));
        Assert.Equal(0, writes);
    }

    [Fact]
    public void PostInsertCancellation_PreservesObservedModuleAndPrimaryCancellation()
    {
        using var cancellation = new CancellationTokenSource();
        string current = "old";
        int inserts = 0;
        var failure = Assert.Throws<VbaSourceEditException>(() => VbaProcedureSource.Replace("M", Source,
            () => current = "", source => { inserts++; current = source; cancellation.Cancel(); },
            () => current, () => "Option Explicit\n" + current, cancellation.Token));
        Assert.IsType<OperationCanceledException>(failure.InnerException);
        Assert.Equal("M", failure.State.ModuleName);
        Assert.Equal("verifying", failure.State.Phase);
        Assert.Equal(VbaModuleSource.ComputeHash("Option Explicit\n" + current), failure.State.SourceHash);
        Assert.Equal(4, failure.State.LineCount);
        Assert.Null(failure.State.ReadBackError);
        Assert.Equal(1, inserts);
        Assert.Equal(Source.Replace("\n", "\r\n", StringComparison.Ordinal), current);
    }

    [Fact]
    public void FailedVerification_ReportsChangedSourceAndIndependentReadFailure()
    {
        string current = "old";
        var failure = Assert.Throws<VbaSourceEditException>(() => VbaProcedureSource.Replace("M", Source,
            () => current = "", _ => current = "unexpected", () => current,
            () => throw new InvalidOperationException("observed read failed")));
        Assert.Contains("differs", failure.InnerException!.Message, StringComparison.Ordinal);
        Assert.Equal("verifying", failure.State.Phase);
        Assert.Equal("observed read failed", failure.State.ReadBackError);
        Assert.Null(failure.State.SourceHash);
        Assert.Equal("unexpected", current);
    }

    [Fact]
    public void PartialDeleteFailure_IsReportedBeforeInsertionWithoutRollback()
    {
        int inserts = 0;
        var primary = new InvalidOperationException("delete failed");
        var failure = Assert.Throws<VbaSourceEditException>(() => VbaProcedureSource.Replace("M", Source,
            () => throw primary, _ => inserts++, () => Source, () => "remaining"));
        Assert.Same(primary, failure.InnerException);
        Assert.Equal("deleting", failure.State.Phase);
        Assert.Equal(VbaModuleSource.ComputeHash("remaining"), failure.State.SourceHash);
        Assert.Equal(0, inserts);
    }

    [Fact]
    public void Replacement_NormalizesBrokenNewlinesAndAcceptsEditorRecasing()
    {
        string current = "old";
        string returned = VbaProcedureSource.Replace("M", Source.Replace("\n", "\r\r\n", StringComparison.Ordinal),
            () => current = "", source => current = source.ToUpperInvariant(), () => current, () => current);
        Assert.Equal("SUB F()\r\nDEBUG.PRINT 1\r\nEND SUB", returned);
    }
}
