using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for range hyperlinks operations
/// </summary>
public sealed partial class PersistentServiceRangeHyperlinksTests
{
    // === HYPERLINK OPERATIONS TESTS ===

    [Fact]
    public void AddHyperlink_CreatesHyperlink()
    {
        // Arrange & Act
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var result = _commands.AddHyperlink(
            batch,
            sheetName,
            "A1",
            "https://www.example.com/",
            "Example Site",
            "Click to visit");

        // Assert
        Assert.True(result.Success);

        // Verify hyperlink exists
        var hyperlinkResult = _commands.GetHyperlink(batch, sheetName, "A1");
        Assert.True(hyperlinkResult.Success);
        var hyperlink = Assert.Single(hyperlinkResult.Hyperlinks);
        Assert.Equal("https://www.example.com/", hyperlink.Address);
        Assert.Equal("A1", hyperlink.CellAddress);
        Assert.Equal("Example Site", hyperlink.DisplayText);
        Assert.Equal("Click to visit", hyperlink.ScreenTip);
        Assert.False(hyperlink.IsInternal);
        var cells = _commands.GetValues(batch, sheetName, "A1");
        Assert.True(cells.Success, cells.ErrorMessage);
        Assert.Equal("Example Site", Assert.Single(Assert.Single(cells.Values)));
    }

    [Fact]
    public void RemoveHyperlink_DeletesHyperlink()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        Assert.True(_commands.AddHyperlink(batch, sheetName, "A1",
            "https://www.example.com/", "Retained text", "Original tooltip").Success);
        Assert.True(_commands.AddHyperlink(batch, sheetName, "B1",
            "https://other.example.com/", "Untouched text", "Untouched tooltip").Success);

        // Act
        var result = _commands.RemoveHyperlink(batch, sheetName, "A1");

        // Assert
        Assert.True(result.Success);

        var hyperlinkResult = _commands.GetHyperlink(batch, sheetName, "A1");
        Assert.True(hyperlinkResult.Success, hyperlinkResult.ErrorMessage);
        Assert.Empty(hyperlinkResult.Hyperlinks);
        var cells = _commands.GetValues(batch, sheetName, "A1:B1");
        Assert.True(cells.Success, cells.ErrorMessage);
        Assert.Equal(["Retained text", "Untouched text"], Assert.Single(cells.Values));
        var remaining = _commands.ListHyperlinks(batch, sheetName);
        Assert.True(remaining.Success, remaining.ErrorMessage);
        var untouched = Assert.Single(remaining.Hyperlinks);
        Assert.Equal("B1", untouched.CellAddress);
        Assert.Equal("https://other.example.com/", untouched.Address);
        Assert.Equal("Untouched text", untouched.DisplayText);
        Assert.Equal("Untouched tooltip", untouched.ScreenTip);
    }

    [Fact]
    public void ListHyperlinks_ReturnsAllHyperlinks()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        Assert.True(_commands.AddHyperlink(batch, sheetName, "A1", "https://site1.com/", "First").Success);
        Assert.True(_commands.AddHyperlink(batch, sheetName, "B2", "https://site2.com/", "Second").Success);
        Assert.True(_commands.AddHyperlink(batch, sheetName, "C3", "https://site3.com/", "Third").Success);

        // Act
        var result = _commands.ListHyperlinks(batch, sheetName);

        // Assert
        Assert.True(result.Success);
        Assert.Equal(3, result.Hyperlinks.Count);
        Assert.Equal(sheetName, result.SheetName);
        var hyperlinks = result.Hyperlinks.OrderBy(link => link.CellAddress, StringComparer.Ordinal).ToArray();
        Assert.Equal(["A1", "B2", "C3"], hyperlinks.Select(link => link.CellAddress));
        Assert.Equal(["https://site1.com/", "https://site2.com/", "https://site3.com/"],
            hyperlinks.Select(link => link.Address));
        Assert.Equal(["First", "Second", "Third"], hyperlinks.Select(link => link.DisplayText));
        Assert.All(hyperlinks, link => Assert.False(link.IsInternal));
    }

    [Fact]
    public void AddHyperlink_InternalTarget_RoundTripsSubAddress()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var add = _commands.AddHyperlink(
            batch,
            sheetName,
            "A1",
            url: null,
            displayText: "Jump",
            tooltip: "Go to target",
            subAddress: $"'{sheetName}'!D5");
        var get = _commands.GetHyperlink(batch, sheetName, "A1");

        Assert.True(add.Success, add.ErrorMessage);
        Assert.True(get.Success, get.ErrorMessage);
        var hyperlink = Assert.Single(get.Hyperlinks);
        Assert.True(hyperlink.IsInternal);
        Assert.Equal($"'{sheetName}'!D5", hyperlink.SubAddress);
        Assert.Equal("Jump", hyperlink.DisplayText);
        Assert.Equal("Go to target", hyperlink.ScreenTip);
        Assert.Equal("A1", hyperlink.CellAddress);
        Assert.Equal(string.Empty, hyperlink.Address);
    }

    [Fact]
    public void UpdateHyperlink_ChangesTargetAndDisplayMetadata()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.AddHyperlink(batch, sheetName, "A1", "https://old.example.com/", "Old").Success);
        Assert.True(_commands.AddHyperlink(batch, sheetName, "B1", "https://other.example.com/", "Untouched").Success);
        var before = _commands.GetHyperlink(batch, sheetName, "A1");
        Assert.True(before.Success, before.ErrorMessage);
        Assert.Equal("https://old.example.com/", Assert.Single(before.Hyperlinks).Address);

        string tooltip = new('T', 255);
        var update = _commands.UpdateHyperlink(
            batch,
            sheetName,
            "A1",
            url: "https://new.example.com/",
            subAddress: "section",
            displayText: "New",
            tooltip: tooltip);
        var get = _commands.GetHyperlink(batch, sheetName, "A1");

        Assert.True(update.Success, update.ErrorMessage);
        Assert.True(get.Success, get.ErrorMessage);
        var hyperlink = Assert.Single(get.Hyperlinks);
        Assert.Equal("https://new.example.com/", hyperlink.Address);
        Assert.Equal("section", hyperlink.SubAddress);
        Assert.Equal("New", hyperlink.DisplayText);
        Assert.Equal(tooltip, hyperlink.ScreenTip);
        Assert.Equal("A1", hyperlink.CellAddress);
        var cells = _commands.GetValues(batch, sheetName, "A1:B1");
        Assert.True(cells.Success, cells.ErrorMessage);
        Assert.Equal(["New", "Untouched"], Assert.Single(cells.Values));
        var untouched = _commands.GetHyperlink(batch, sheetName, "B1");
        Assert.True(untouched.Success, untouched.ErrorMessage);
        Assert.Equal("https://other.example.com/", Assert.Single(untouched.Hyperlinks).Address);
    }

    [Fact]
    public void UpdateHyperlink_PreservesLabelAndTooltipAcrossInternalAndExternalTargets()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        const string url = "https://example.com/guide";
        string internalTarget = $"'{sheetName}'!D5";
        Assert.True(_commands.AddHyperlink(batch, sheetName, "A1", url,
            displayText: "Original label", tooltip: "Original tooltip").Success);

        Assert.True(_commands.UpdateHyperlink(batch, sheetName, "A1", url: string.Empty, subAddress: internalTarget).Success);
        var internalRead = _commands.GetHyperlink(batch, sheetName, "A1");
        Assert.True(internalRead.Success, internalRead.ErrorMessage);
        var internalLink = Assert.Single(internalRead.Hyperlinks);
        Assert.True(internalLink.IsInternal);
        Assert.Equal(internalTarget, internalLink.SubAddress);
        Assert.Equal("Original label", internalLink.DisplayText);
        Assert.Equal("Original tooltip", internalLink.ScreenTip);
        Assert.Equal("A1", Assert.Single(_commands.ListHyperlinks(batch, sheetName).Hyperlinks).CellAddress);

        Assert.True(_commands.UpdateHyperlink(batch, sheetName, "A1", url: url, subAddress: string.Empty).Success);
        var externalRead = _commands.GetHyperlink(batch, sheetName, "A1");
        Assert.True(externalRead.Success, externalRead.ErrorMessage);
        var externalLink = Assert.Single(externalRead.Hyperlinks);
        Assert.False(externalLink.IsInternal);
        Assert.Equal(url, externalLink.Address);
        Assert.True(string.IsNullOrEmpty(externalLink.SubAddress));
        Assert.Equal("Original label", externalLink.DisplayText);
        Assert.Equal("Original tooltip", externalLink.ScreenTip);
        Assert.Equal("A1", Assert.Single(_commands.ListHyperlinks(batch, sheetName).Hyperlinks).CellAddress);
    }

    [Fact]
    public void ListHyperlinks_IncludesInternalTargetMetadata()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.AddHyperlink(batch, sheetName, "A1", url: null,
            displayText: "Jump", tooltip: "Go to B2", subAddress: $"'{sheetName}'!B2").Success);

        var list = _commands.ListHyperlinks(batch, sheetName);

        Assert.True(list.Success, list.ErrorMessage);
        var hyperlink = Assert.Single(list.Hyperlinks);
        Assert.True(hyperlink.IsInternal);
        Assert.Equal($"'{sheetName}'!B2", hyperlink.SubAddress);
        Assert.Equal("A1", hyperlink.CellAddress);
        Assert.Equal(string.Empty, hyperlink.Address);
        Assert.Equal("Jump", hyperlink.DisplayText);
        Assert.Equal("Go to B2", hyperlink.ScreenTip);
    }

    [Theory]
    [InlineData("rangelink.add-hyperlink", "A2")]
    [InlineData("rangelink.update-hyperlink", "A1")]
    public async Task Hyperlink_OversizedTooltipPreservesLinksAndCells(string command, string cellAddress)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.AddHyperlink(batch, sheetName, "A1", "https://old.example.com/",
            displayText: "Original", tooltip: "Original tooltip").Success);
        const string rangeAddress = "A1:A2";
        var sheet = new { sheetName };
        var range = new { sheetName, rangeAddress };
        string? linksBefore = _fixture.Send("rangelink.list-hyperlinks", sheet).Result;
        string? cellsBefore = _fixture.Send("range.get-values", range).Result;

        var rejected = await _fixture.SendForFailureAsync(command, new
        {
            sheetName,
            cellAddress,
            url = "https://new.example.com/",
            displayText = "Rejected text",
            tooltip = new string('T', 256)
        });

        Assert.Equal("InvalidInput", rejected.ErrorCategory);
        Assert.Contains("cannot exceed 255 characters", rejected.ErrorMessage);
        Assert.Equal(linksBefore, _fixture.Send("rangelink.list-hyperlinks", sheet).Result);
        Assert.Equal(cellsBefore, _fixture.Send("range.get-values", range).Result);
    }

    [Fact]
    public void UpdateHyperlink_CannotRemoveOnlyTarget()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var originalTarget = $"'{sheetName}'!B2";
        Assert.True(_commands.AddHyperlink(batch, sheetName, "A1", url: null,
            displayText: "Original", tooltip: "Original tooltip", subAddress: originalTarget).Success);
        Assert.True(_commands.AddHyperlink(batch, sheetName, "B1", "https://other.example.com/", "Untouched").Success);
        var before = _commands.ListHyperlinks(batch, sheetName);
        Assert.True(before.Success, before.ErrorMessage);
        var cellsBefore = _commands.GetValues(batch, sheetName, "A1:B1");
        Assert.True(cellsBefore.Success, cellsBefore.ErrorMessage);

        var exception = Assert.Throws<ArgumentException>(() =>
            _commands.UpdateHyperlink(batch, sheetName, "A1", subAddress: string.Empty,
                displayText: "Rejected text", tooltip: "Rejected tooltip"));
        Assert.Contains("must retain either an external address or an internal subAddress",
            exception.Message, StringComparison.Ordinal);
        var get = _commands.GetHyperlink(batch, sheetName, "A1");

        Assert.True(get.Success, get.ErrorMessage);
        var hyperlink = Assert.Single(get.Hyperlinks);
        Assert.True(hyperlink.IsInternal);
        Assert.Equal(originalTarget, hyperlink.SubAddress);
        Assert.Equal("Original", hyperlink.DisplayText);
        Assert.Equal("Original tooltip", hyperlink.ScreenTip);
        var after = _commands.ListHyperlinks(batch, sheetName);
        Assert.True(after.Success, after.ErrorMessage);
        Assert.Equal(JsonSerializer.Serialize(before.Hyperlinks), JsonSerializer.Serialize(after.Hyperlinks));
        var cellsAfter = _commands.GetValues(batch, sheetName, "A1:B1");
        Assert.True(cellsAfter.Success, cellsAfter.ErrorMessage);
        Assert.Equal(JsonSerializer.Serialize(cellsBefore.Values), JsonSerializer.Serialize(cellsAfter.Values));
    }

    [Theory]
    [InlineData("rangelink.add-hyperlink", "A1:B1")]
    [InlineData("rangelink.add-hyperlink", "A1,C1")]
    [InlineData("rangelink.update-hyperlink", "A1:B1")]
    [InlineData("rangelink.update-hyperlink", "A1,C1")]
    [InlineData("rangelink.get-hyperlink", "A1:B1")]
    [InlineData("rangelink.get-hyperlink", "A1,C1")]
    public async Task SingleCellHyperlinkCommand_RejectsMultipleCellsWithoutChangingThem(
        string command, string cellAddress)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.AddHyperlink(batch, sheetName, "A1", "https://first.example.com/",
            "First", "First tip"));
        RequireSuccess(_commands.AddHyperlink(batch, sheetName, "B1", "https://second.example.com/",
            "Second", "Second tip"));
        RequireSuccess(_commands.AddHyperlink(batch, sheetName, "C1", "https://third.example.com/",
            "Third", "Third tip"));

        var sheet = new { sheetName };
        var range = new { sheetName, rangeAddress = "A1:C1" };
        string? linksBefore = _fixture.Send("rangelink.list-hyperlinks", sheet).Result;
        string? cellsBefore = _fixture.Send("range.get-values", range).Result;
        Assert.NotNull(linksBefore);
        Assert.NotNull(cellsBefore);

        object args = command == "rangelink.get-hyperlink"
            ? new { sheetName, cellAddress }
            : new
            {
                sheetName,
                cellAddress,
                url = "https://replacement.example.com/",
                displayText = "Replacement",
                tooltip = "Replacement tip"
            };
        var rejected = await _fixture.SendForFailureAsync(command, args);
        Assert.Equal("InvalidInput", rejected.ErrorCategory);
        Assert.Contains("exactly one cell", rejected.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(linksBefore, _fixture.Send("rangelink.list-hyperlinks", sheet).Result);
        Assert.Equal(cellsBefore, _fixture.Send("range.get-values", range).Result);
    }
}
