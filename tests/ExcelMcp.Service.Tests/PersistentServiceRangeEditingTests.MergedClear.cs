using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceRangeEditingTests
{
    [Fact]
    public async Task ClearContents_PartialMerge_RejectsWithoutExpandingRequestedRange()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheet, "A2:E2", [["left", "keep", null, null, "right"]]));
        RequireSuccess(_commands.MergeCells(batch, sheet, "B2:D2"));

        var rejected = await _fixture.SendForFailureAsync("range.clear-contents", new
        {
            sheetName = sheet,
            rangeAddress = "B2"
        });
        Assert.False(rejected.Success);
        Assert.False(string.IsNullOrWhiteSpace(rejected.ErrorMessage));
        var unchanged = RequireSuccess(_commands.GetValues(batch, sheet, "A2:E2"));
        Assert.Equal(new object?[] { "left", "keep", null, null, "right" }, unchanged.Values[0]);
        var merge = RequireSuccess(_commands.GetMergeInfo(batch, sheet, "A2:E2"));
        Assert.True(merge.IsMerged);
        Assert.Equal(["$B$2:$D$2"], merge.MergedRanges);

        RequireSuccess(_commands.ClearContents(batch, sheet, "B2:D2"));
        var cleared = RequireSuccess(_commands.GetValues(batch, sheet, "A2:E2"));
        Assert.Equal(new object?[] { "left", null, null, null, "right" }, cleared.Values[0]);
        var preserved = RequireSuccess(_commands.GetMergeInfo(batch, sheet, "A2:E2"));
        Assert.True(preserved.IsMerged);
        Assert.Equal(merge.MergedRanges, preserved.MergedRanges);
    }
}
