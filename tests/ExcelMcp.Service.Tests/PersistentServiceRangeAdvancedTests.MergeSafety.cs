using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceRangeAdvancedTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task MergeCells_RejectsOccupiedNonAnchorAndPreservesWorkbook(bool blankFormula)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheetName, "A1", [["anchor"]]));
        if (blankFormula)
            RequireSuccess(_commands.SetFormulas(batch, sheetName, "B1", [["=\"\""]]));
        else
            RequireSuccess(_commands.SetValues(batch, sheetName, "B1", [["do not discard"]]));
        var before = RequireSuccess(_commands.GetFormulas(batch, sheetName, "A1:B1"));

        var response = await _fixture.SendForFailureAsync("rangeformat.merge-cells",
            new { sheetName, rangeAddress = "A1:B1" });

        Assert.Equal("Conflict", response.ErrorCategory);
        Assert.Contains("$B$1", response.ErrorMessage, StringComparison.Ordinal);
        var after = RequireSuccess(_commands.GetFormulas(batch, sheetName, "A1:B1"));
        Assert.Equal(before.Formulas[0], after.Formulas[0]);
        Assert.Equal(before.Values[0], after.Values[0]);
        Assert.False(RequireSuccess(_commands.GetMergeInfo(batch, sheetName, "A1:B1")).IsMerged);
    }

    [Fact]
    public void MergeCells_ExplicitAllowDiscardsOnlyTheOtherCellsInTheRectangle()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheetName, "A1:C1", [["anchor", "discard", "outside"]]));

        var response = _fixture.Send("rangeformat.merge-cells",
            new { sheetName, rangeAddress = "A1:B1", overwritePolicy = "allow" });

        Assert.True(response.Success, response.ErrorMessage);
        Assert.Equal(["$A$1:$B$1"], RequireSuccess(_commands.GetMergeInfo(batch, sheetName, "A1:C1")).MergedRanges);
        AssertCells(sheetName, "A1:C1", [["anchor", null, "outside"]]);
    }

    [Fact]
    public async Task MergeCells_RejectsDisjointAndPartialExistingMergeGeometry()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheetName, "B1", [["existing"]]));
        RequireSuccess(_commands.MergeCells(batch, sheetName, "B1:C1"));
        var disjoint = await _fixture.SendForFailureAsync("rangeformat.merge-cells",
            new { sheetName, rangeAddress = "A1:A2,C3:C4", overwritePolicy = "allow" });
        var partial = await _fixture.SendForFailureAsync("rangeformat.merge-cells",
            new { sheetName, rangeAddress = "A1:B1", overwritePolicy = "allow" });

        Assert.Equal("InvalidInput", disjoint.ErrorCategory);
        Assert.Contains("single rectangular range", disjoint.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal("Conflict", partial.ErrorCategory);
        Assert.Contains("$B$1:$C$1", partial.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(["$B$1:$C$1"], RequireSuccess(_commands.GetMergeInfo(batch, sheetName, "A1:D4")).MergedRanges);
        AssertCells(sheetName, "A1:C1", [[null, "existing", null]]);
    }

    [Fact]
    public void MergeCells_PreservesTheOnlyOccupiedAnchorAndIsRepeatable()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheetName, "A1", [["only content"]]));

        RequireSuccess(_commands.MergeCells(batch, sheetName, "A1:B2"));
        RequireSuccess(_commands.MergeCells(batch, sheetName, "A1:B2"));

        Assert.Equal(["$A$1:$B$2"], RequireSuccess(_commands.GetMergeInfo(batch, sheetName, "A1:B2")).MergedRanges);
        AssertCells(sheetName, "A1:B2", [["only content", null], [null, null]]);
    }
}
