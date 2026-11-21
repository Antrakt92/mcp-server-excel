using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceRangeValidationTests
{
    [Theory]
    [InlineData("A1", false)]
    [InlineData("A1:A3", false)]
    [InlineData("A1:A3", true)]
    public async Task ValidateRange_ExcelRejectsFormula_PreservesExistingAndBlankRules(
        string rangeAddress, bool firstCellBlank)
    {
        var sheet = CreateValidationReplacementScenario(rangeAddress, firstCellBlank);
        var originalRules = ReadReplacementRules(sheet);
        var originalValues = ReadReplacementValues(sheet);

        var response = await _fixture.SendForFailureAsync("rangeformat.validate-range", new
        {
            sheetName = sheet,
            rangeAddress,
            validationType = "custom",
            formula1 = "=("
        });

        Assert.Equal("ComInterop", response.ErrorCategory);
        Assert.Equal("0x800A03EC", response.HResult);
        Assert.Equal(originalRules, ReadReplacementRules(sheet));
        Assert.Equal(originalValues, ReadReplacementValues(sheet));
    }

    [Theory]
    [InlineData("A1", false)]
    [InlineData("A1:A3", false)]
    [InlineData("A1:A3", true)]
    public void ValidateRange_ValidReplacement_UpdatesOnlySelectedCellsAndClearsOldMessages(
        string rangeAddress, bool firstCellBlank)
    {
        var sheet = CreateValidationReplacementScenario(rangeAddress, firstCellBlank);
        var originalRules = ReadReplacementRules(sheet);
        var originalValues = ReadReplacementValues(sheet);

        RequireSuccess(_commands.ValidateRange(_fixture.BatchToken, sheet, rangeAddress,
            "whole", "between", "2", "8", false, null, null, false, "warning", null, null, true, null));

        var selectedCount = rangeAddress == "A1" ? 1 : 3;
        for (var row = 1; row <= selectedCount; row++)
        {
            var rule = RequireSuccess(_commands.GetValidation(_fixture.BatchToken, sheet, $"A{row}"));
            Assert.True(rule.HasValidation);
            Assert.Equal("whole", rule.ValidationType);
            Assert.Equal("between", rule.ValidationOperator);
            Assert.Equal("2", rule.Formula1);
            Assert.Equal("8", rule.Formula2);
            Assert.Equal("warning", rule.ErrorStyle);
            Assert.True(rule.IgnoreBlank);
            Assert.False(rule.ShowInputMessage);
            Assert.False(rule.ShowErrorAlert);
            Assert.Empty(rule.InputTitle!);
            Assert.Empty(rule.InputMessage!);
            Assert.Empty(rule.ErrorTitle!);
            Assert.Empty(rule.ValidationErrorMessage!);
        }
        Assert.Equal(originalRules.Skip(selectedCount), ReadReplacementRules(sheet).Skip(selectedCount));
        Assert.Equal(originalValues, ReadReplacementValues(sheet));
    }

    private string CreateValidationReplacementScenario(string target, bool firstCellBlank)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheet, "A1:D4",
            [[5, null, null, 5], [6, null, null, 6], [7, null, null, 7], [8, null, null, null]]));
        RequireSuccess(_commands.ValidateRange(batch, sheet, "A1:A4", "list", null,
            "=$D$1:$D$3", null, true, "Original input", "Choose a number",
            true, "warning", "Original alert", "Use the list", false, false));

        if (target != "A1")
        {
            RequireSuccess(_commands.ValidateRange(batch, sheet, "A2", "whole", "between", "1", "20",
                true, "Whole number", "Choose an integer", true, "stop", "Original whole rule", "Use 1 to 20", true, null));
            RequireSuccess(_commands.RemoveValidation(batch, sheet, "A3"));
            if (firstCellBlank)
                RequireSuccess(_commands.RemoveValidation(batch, sheet, "A1"));
        }

        var first = RequireSuccess(_commands.GetValidation(batch, sheet, "A1"));
        Assert.Equal(!firstCellBlank, first.HasValidation);
        var second = RequireSuccess(_commands.GetValidation(batch, sheet, "A2"));
        Assert.Equal(target == "A1" ? "list" : "whole", second.ValidationType);
        var third = RequireSuccess(_commands.GetValidation(batch, sheet, "A3"));
        Assert.Equal(target == "A1", third.HasValidation);
        Assert.Equal("list", RequireSuccess(_commands.GetValidation(batch, sheet, "A4")).ValidationType);
        return sheet;
    }

    private string[] ReadReplacementRules(string sheet) =>
        Enumerable.Range(1, 4).Select(row =>
        {
            var address = $"A{row}";
            var rule = RequireSuccess(_commands.GetValidation(_fixture.BatchToken, sheet, address));
            bool? dropdown = rule.ValidationType == "list" ? ReadDropdown(sheet, address) : null;
            return JsonSerializer.Serialize(new { rule, dropdown });
        }).ToArray();

    private string ReadReplacementValues(string sheet) =>
        JsonSerializer.Serialize(RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheet, "A1:D4")).Values);
}
