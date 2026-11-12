using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceRangeSearchTests
{
    [Theory]
    [InlineData(true, "cat", "bird")]
    [InlineData(false, "cat cat", "bird cat")]
    public void Replace_FirstOccurrence_LeavesRemainingMatchesUnchanged(bool matchEntireCell, string text, string expected)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheet, "A1:B2",
            [[text, text], [text, "outside"]]).Success);

        var replaced = _commands.Replace(batch, sheet, "A1:B2", "cat", "bird", new ReplaceOptions
        {
            MatchEntireCell = matchEntireCell,
            ReplaceAll = false
        });

        Assert.True(replaced.Success, replaced.ErrorMessage);
        var read = _commands.GetValues(batch, sheet, "A1:B2");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Collection(read.Values,
            row => Assert.Equal([expected, text], row),
            row => Assert.Equal([text, "outside"], row));
    }

    [Fact]
    public void Replace_FirstOccurrence_PreservesTheRestOfTheFormula()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetFormulas(batch, sheet, "A1:A2",
            [[@"=""cat""&""cat"""], [@"=""cat"""]]).Success);

        var replaced = _commands.Replace(batch, sheet, "A1:A2", "cat", "bird", new ReplaceOptions
        {
            ReplaceAll = false
        });

        Assert.True(replaced.Success, replaced.ErrorMessage);
        var read = _commands.GetFormulas(batch, sheet, "A1:A2");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal([@"=""bird""&""cat""", @"=""cat"""], read.Formulas.Select(row => Assert.Single(row)));
        Assert.Equal(["birdcat", "cat"], read.Values.Select(row => Assert.Single(row)));
    }

    [Fact]
    public void Replace_FirstOccurrence_UsesRowOrderAndMatchCaseAcrossAreas()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheet, "A1:B2",
            [["CAT", "cat"], ["cat", "cat"]]).Success);

        var replaced = _commands.Replace(batch, sheet, "A2:B2,A1:B1", "cat", "bird", new ReplaceOptions
        {
            MatchCase = true,
            ReplaceAll = false
        });

        Assert.True(replaced.Success, replaced.ErrorMessage);
        var read = _commands.GetValues(batch, sheet, "A1:B2");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Collection(read.Values,
            row => Assert.Equal(["CAT", "bird"], row),
            row => Assert.Equal(["cat", "cat"], row));
    }

    [Theory]
    [InlineData("123", "567")]
    [InlineData("12?", "567")]
    public void Replace_FirstOccurrence_WholeCellMatchesNumericValues(string pattern, string replacement)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheet, "A1:A2", [[123], [123]]).Success);

        var replaced = _commands.Replace(batch, sheet, "A1:A2", pattern, replacement, new ReplaceOptions
        {
            MatchEntireCell = true,
            ReplaceAll = false
        });

        Assert.True(replaced.Success, replaced.ErrorMessage);
        var read = _commands.GetValues(batch, sheet, "A1:A2");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal([567, 123], read.Values.Select(row => Assert.IsType<int>(Assert.Single(row))));
    }

    [Fact]
    public void Replace_FirstOccurrence_ResolvesEscapedWildcardAndPreservesLiteralText()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheet, "A1:A2", [["'=x*x*"], ["*"]]).Success);

        var replaced = _commands.Replace(batch, sheet, "A1:A2", "~*", "+", new ReplaceOptions
        {
            ReplaceAll = false
        });

        Assert.True(replaced.Success, replaced.ErrorMessage);
        var read = _commands.GetValues(batch, sheet, "A1:A2");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal(["=x+x*", "*"], read.Values.Select(row => Assert.Single(row)));
    }

    [Fact]
    public void Replace_FirstOccurrence_NoMatchLeavesCellsUnchanged()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheet, "A1", [["CAT"]]).Success);

        var replaced = _commands.Replace(batch, sheet, "A1", "cat", "bird", new ReplaceOptions
        {
            MatchCase = true,
            ReplaceAll = false
        });

        Assert.True(replaced.Success, replaced.ErrorMessage);
        Assert.Equal("CAT", _commands.GetValues(batch, sheet, "A1").Values[0][0]);
    }

    [Theory]
    [InlineData("cat", "c?t", "literal")]
    [InlineData(123, "12", "matchEntireCell")]
    public void Replace_FirstOccurrence_RejectsUnsupportedPartialMatchBeforeWriting(
        object value, string pattern, string expectedError)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheet, "A1:A2", [[value], [value]]).Success);

        var error = Assert.ThrowsAny<Exception>(() =>
            _commands.Replace(batch, sheet, "A1:A2", pattern, "bird", new ReplaceOptions { ReplaceAll = false }));

        Assert.Contains(expectedError, error.Message, StringComparison.OrdinalIgnoreCase);
        var read = _commands.GetValues(batch, sheet, "A1:A2");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal([value, value], read.Values.Select(row => Assert.Single(row)));
    }
}
