using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceRangeReadBudgetTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public async Task LongTextReads_RejectBudgetOverflowWithoutReturningPrivateValues()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        string text = "private-marker" + new string('x', 31_980);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:AF1",
            [Enumerable.Repeat<object?>(text, 32).ToList()]));
        foreach (var command in new[] { "range.get-values", "range.get-formulas", "range.get-used-range", "range.get-current-region", "range.get-values-page" })
        {
            object args = command switch
            {
                "range.get-used-range" => new { sheetName },
                "range.get-current-region" => new { sheetName, cellAddress = "A1" },
                "range.get-values-page" => new { sheetName, rangeAddress = "A1:AF1", pageRows = 1, pageColumns = 32 },
                _ => new { sheetName, rangeAddress = "A1:AF1" }
            };
            var response = await _fixture.SendForFailureAsync(command, args);
            Assert.True(response.ErrorCategory == "InvalidInput",
                $"{command}: expected InvalidInput; category={response.ErrorCategory}; error={response.ErrorMessage}; HRESULT={response.HResult}; inner={response.InnerError}");
            Assert.Contains("byte budget", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("private-marker", response.ErrorMessage, StringComparison.Ordinal);
            Assert.Null(response.Result);
        }
        var page = RequireSuccess(_commands.GetValuesPage(_fixture.BatchToken, sheetName, "A1:AF1", 0, 0, 1, 4));
        Assert.Equal(4, Assert.Single(page.Values).Count);
        Assert.All(page.Values[0], value => Assert.Equal(text, value));
        Assert.True(page.HasMore);
        Assert.Equal(4, page.NextColumnOffset);
    }

    [Fact]
    public void Tiles_AssembleCompleteRowsAndLastTileWithFormulaNotationAndErrorCoordinates()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var seeds = Enumerable.Range(0, 31).Select(row => new List<object?> { row * 1000 }).ToList();
        var formulas = Enumerable.Range(0, 31).Select(_ => Enumerable.Repeat<object?>("=RC[-1]+1", 37).ToList()).ToList();
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheet, "B3:B33", seeds));
        RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, sheet, "C3:AM33", formulas,
            referenceStyle: FormulaReferenceStyle.R1C1));
        var values = RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheet, "C3:AM33"));
        var page = RequireSuccess(_commands.GetValuesPage(_fixture.BatchToken, sheet, "C3:AM33", 0, 0, 31, 37));
        var relative = RequireSuccess(_commands.GetFormulas(_fixture.BatchToken, sheet, "C3:AM33", FormulaReferenceStyle.R1C1));
        var absolute = RequireSuccess(_commands.GetFormulas(_fixture.BatchToken, sheet, "C3:AM33"));
        Assert.Equal((31, 37), (values.RowCount, values.ColumnCount));
        for (int row = 0; row < 31; row++)
        {
            Assert.Equal(37, values.Values[row].Count);
            Assert.Equal(37, relative.Formulas[row].Count);
            for (int column = 0; column < 37; column++)
            {
                RangeReadAssertions.Number(row * 1000 + column + 1, values.Values[row][column]);
                RangeReadAssertions.Number(row * 1000 + column + 1, page.Values[row][column]);
                RangeReadAssertions.Number(row * 1000 + column + 1, relative.Values[row][column]);
                Assert.Equal("=RC[-1]+1", relative.Formulas[row][column]);
            }
        }
        Assert.Equal("=B3+1", absolute.Formulas[0][0]);
        Assert.Equal("=AL33+1", absolute.Formulas[30][36]);
        RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, sheet, "AM33", [["=RC[-1]/0"]],
            overwritePolicy: OverwritePolicy.Allow, referenceStyle: FormulaReferenceStyle.R1C1));
        var error = RequireSuccess(_commands.GetFormulas(_fixture.BatchToken, sheet, "C3:AM33", FormulaReferenceStyle.R1C1));
        Assert.Equal("#DIV/0!", error.Values[30][36]);
        Assert.Equal("AM33", Assert.Single(error.CellErrors).CellAddress);
        Assert.Equal("=RC[-1]/0", error.CellErrors[0].Formula);
    }
}
