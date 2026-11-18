using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceRangeValidationTests
{
    [Fact]
    public void ValidateRange_OversizedFormula_PreservesExistingRule()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.ValidateRange(
            batch, sheetName, "A1", "whole", "between", "1", "10",
            true, "Original title", "Original message",
            true, "warning", "Original error", "Original error message",
            false, null));
        var before = RequireSuccess(_commands.GetValidation(batch, sheetName, "A1"));

        var failure = Record.Exception(() => _fixture.Send("rangeformat.validate-range", new
        {
            sheetName,
            rangeAddress = "A1",
            validationType = "list",
            formula1 = string.Join(',', Enumerable.Repeat("choice", 37))
        }));
        var after = RequireSuccess(_commands.GetValidation(batch, sheetName, "A1"));

        Assert.Equal(JsonSerializer.Serialize(before), JsonSerializer.Serialize(after));
        var invalidInput = Assert.IsType<ArgumentException>(failure);
        Assert.Contains("[InvalidInput/", invalidInput.Message, StringComparison.Ordinal);
        Assert.Contains("255", invalidInput.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("list", null, "ignored")]
    [InlineData("custom", "", "ignored")]
    [InlineData("whole", null, "")]
    [InlineData("decimal", null, null)]
    [InlineData("time", "", null)]
    public void ValidateRange_MissingRequiredFormula_PreservesExistingRule(
        string type, string? formula1, string? formula2)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.ValidateRange(
            batch, sheet, "A1", "whole", "between", "1", "10",
            true, "Original title", "Original message",
            true, "warning", "Original error", "Original error message",
            false, null));
        var before = RequireSuccess(_commands.GetValidation(batch, sheet, "A1"));
        Assert.True(before.HasValidation);

        var failure = Record.Exception(() => _fixture.Send("rangeformat.validate-range", new
        {
            sheetName = sheet,
            rangeAddress = "A1",
            validationType = type,
            formula1,
            formula2
        }));
        var after = RequireSuccess(_commands.GetValidation(batch, sheet, "A1"));

        Assert.Equal(JsonSerializer.Serialize(before), JsonSerializer.Serialize(after));
        var invalidInput = Assert.IsType<ArgumentException>(failure);
        Assert.Contains("[InvalidInput/", invalidInput.Message, StringComparison.Ordinal);
        Assert.Contains("formula1", invalidInput.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("unsupported", "between", "stop")]
    [InlineData("whole", "unsupported", "stop")]
    [InlineData("whole", "between", "unsupported")]
    public void ValidateRange_InvalidOptions_PreserveExistingRule(
        string validationType, string validationOperator, string errorStyle)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.ValidateRange(
            batch, sheet, "A1", "whole", "between", "1", "10",
            true, "Original title", "Original message",
            true, "warning", "Original error", "Original error message",
            false, null));
        var before = _commands.GetValidation(batch, sheet, "A1");
        Assert.True(before.Success, before.ErrorMessage);
        Assert.True(before.HasValidation);

        Assert.Throws<ArgumentException>(() => _commands.ValidateRange(
            batch, sheet, "A1", validationType, validationOperator, "2", "8",
            true, "Replacement title", "Replacement message",
            true, errorStyle, "Replacement error", "Replacement error message",
            true, null));

        var after = _commands.GetValidation(batch, sheet, "A1");
        Assert.True(after.Success, after.ErrorMessage);
        Assert.Equal(JsonSerializer.Serialize(before), JsonSerializer.Serialize(after));
    }
}
