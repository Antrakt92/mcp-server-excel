using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceRangeNumberFormatTests
{
    [Theory]
    [InlineData("A1:A10001", "cell read limit")]
    [InlineData("A1,A3", "rectangular")]
    public async Task GetNumberFormats_RejectsOversizedOrDisjointScopeWithoutChangingCells(string scope, string expected)
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheet, "A1", [[123d]]));
        RequireSuccess(_commands.SetNumberFormat(_fixture.BatchToken, sheet, "A1", "0.00"));
        var response = await _fixture.SendForFailureAsync("range.get-number-formats", new { sheetName = sheet, rangeAddress = scope });
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Contains(expected, response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Null(response.Result);
        Assert.Equal("0.00", RequireSuccess(_commands.GetNumberFormats(_fixture.BatchToken, sheet, "A1")).Formats[0][0]);
        RangeReadAssertions.Number(123d, RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheet, "A1")).Values[0][0]);
    }

    [Fact]
    public async Task GetNumberFormats_RejectsLongRepeatedFormatBudgetAndAllowsSmallerScope()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        string format = "0\"private-format-marker\"";
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheet, "A1", [[123d]]));
        RequireSuccess(_commands.SetNumberFormat(_fixture.BatchToken, sheet, "A1:A10000", format));
        var response = await _fixture.SendForFailureAsync("range.get-number-formats", new { sheetName = sheet, rangeAddress = "A1:A10000" });
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Contains("byte budget", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("private-format-marker", response.ErrorMessage, StringComparison.Ordinal);
        Assert.Null(response.Result);
        Assert.Equal(format, RequireSuccess(_commands.GetNumberFormats(_fixture.BatchToken, sheet, "A1")).Formats[0][0]);
        RangeReadAssertions.Number(123d, RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheet, "A1")).Values[0][0]);
    }
}
