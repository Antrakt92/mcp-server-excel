using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for worksheet comments.
/// </summary>
public sealed partial class PersistentServiceSheetTests
{
    [Fact]
    public void SetComment_RoundsTripThroughSheetAndCanBeCleared()
    {
        var batch = _fixture.BatchToken;
        var sheetName = $"Comments_{Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sheetName);

        var initialComment = _sheetCommands.GetComment(batch, sheetName, "A1");
        RequireSuccess(initialComment);
        Assert.False(initialComment.HasComment);

        var setResult = _sheetCommands.SetComment(batch, sheetName, "A1", "Quarterly update");
        RequireSuccess(setResult);

        var readResult = _sheetCommands.GetComment(batch, sheetName, "A1");
        RequireSuccess(readResult);
        Assert.True(readResult.HasComment);
        Assert.Equal("Quarterly update", readResult.Text);

        var clearResult = _sheetCommands.ClearComment(batch, sheetName, "A1");
        RequireSuccess(clearResult);

        var clearedResult = _sheetCommands.GetComment(batch, sheetName, "A1");
        RequireSuccess(clearedResult);
        Assert.False(clearedResult.HasComment);
        Assert.Null(clearedResult.Text);
    }

    [Theory]
    [InlineData("worksheetstyle.set-comment", "A1:B1")]
    [InlineData("worksheetstyle.set-comment", "A1,C1")]
    [InlineData("worksheetstyle.get-comment", "A1:B1")]
    [InlineData("worksheetstyle.get-comment", "A1,C1")]
    [InlineData("worksheetstyle.clear-comment", "A1:B1")]
    [InlineData("worksheetstyle.clear-comment", "A1,C1")]
    public async Task LegacyNoteCommand_RejectsMultipleCellsWithoutChangingNotes(
        string command, string cellAddress)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_sheetCommands.SetComment(batch, sheetName, "A1", "First note"));
        RequireSuccess(_sheetCommands.SetComment(batch, sheetName, "B1", "Second note"));
        RequireSuccess(_sheetCommands.SetComment(batch, sheetName, "C1", "Third note"));

        object args = command == "worksheetstyle.set-comment"
            ? new { sheetName, cellAddress, text = "Replacement note" }
            : new { sheetName, cellAddress };
        var result = await _fixture.SendForFailureAsync(command, args);
        Assert.Equal("InvalidInput", result.ErrorCategory);
        Assert.Contains("exactly one cell", result.ErrorMessage, StringComparison.Ordinal);

        var first = RequireSuccess(_sheetCommands.GetComment(batch, sheetName, "A1"));
        var second = RequireSuccess(_sheetCommands.GetComment(batch, sheetName, "B1"));
        var third = RequireSuccess(_sheetCommands.GetComment(batch, sheetName, "C1"));
        Assert.True(first.HasComment);
        Assert.True(second.HasComment);
        Assert.True(third.HasComment);
        Assert.Equal("First note", first.Text);
        Assert.Equal("Second note", second.Text);
        Assert.Equal("Third note", third.Text);
    }
}
