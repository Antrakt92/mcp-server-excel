using System.Runtime.ExceptionServices;
using Sbroenne.ExcelMcp.Core.Commands.Calculation;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceRangeBoundedReadsTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void DiscoveryAndPages_PreserveCanonicalErrorsAndActualCellAddresses()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, sheet, "C3:D4",
            [["=1/0", -2146826281d], ["=NA()", 9]]));
        var exact = RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheet, "C3:D4"));
        var used = RequireSuccess(_commands.GetUsedRange(_fixture.BatchToken, sheet));
        var region = RequireSuccess(_commands.GetCurrentRegion(_fixture.BatchToken, sheet, "C3"));
        var page = RequireSuccess(_commands.GetValuesPage(_fixture.BatchToken, sheet, "C3:D4", 0, 0, 2, 2));
        foreach (var result in new[] { used, region, page })
        {
            Assert.Equal(exact.Values[0][0], result.Values[0][0]);
            Assert.Equal("#DIV/0!", result.Values[0][0]);
            Assert.Equal("#N/A", result.Values[1][0]);
            RangeReadAssertions.Number(-2146826281d, result.Values[0][1]);
            Assert.Equal(["C3", "C4"], result.CellErrors.Select(error => error.CellAddress));
            Assert.Equal(exact.CellErrors.Select(error => error.Formula), result.CellErrors.Select(error => error.Formula));
        }
        Assert.False(page.HasMore);
        Assert.Null(page.NextRowOffset);
        Assert.Null(page.NextColumnOffset);
    }

    [Fact]
    public void BoundsAndPaging_ReturnExactScopeAndContinuationWithoutExpandingIt()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheet, "C3:E5", [[1, 2, 3], [4, 5, 6], [7, 8, 9]]));
        var bounds = RequireSuccess(_commands.GetUsedRangeInfo(_fixture.BatchToken, sheet));
        Assert.Equal("$C$3:$E$5", bounds.RangeAddress);
        Assert.Equal((3, 3, 9L), (bounds.RowCount, bounds.ColumnCount, bounds.CellCount));
        var first = RequireSuccess(_commands.GetValuesPage(_fixture.BatchToken, sheet, "C3:E5", 0, 0, 2, 2));
        Assert.Equal("$C$3:$D$4", first.RangeAddress);
        Assert.Equal("$C$3:$E$5", first.ScopeAddress);
        Assert.Equal((3, 3, 2, 2), (first.TotalRowCount, first.TotalColumnCount, first.RowCount, first.ColumnCount));
        Assert.True(first.HasMore);
        Assert.Equal(0, first.NextRowOffset);
        Assert.Equal(2, first.NextColumnOffset);
        RangeReadAssertions.Number(5d, first.Values[1][1]);
        var right = RequireSuccess(_commands.GetValuesPage(_fixture.BatchToken, sheet, "C3:E5", 0, 2, 2, 2));
        Assert.Equal("$E$3:$E$4", right.RangeAddress);
        Assert.Equal((2, 0), (right.NextRowOffset, right.NextColumnOffset));
        RangeReadAssertions.Number(6d, right.Values[1][0]);
        var last = RequireSuccess(_commands.GetValuesPage(_fixture.BatchToken, sheet, "C3:E5", 2, 2, 2, 2));
        Assert.Equal("$E$5", last.RangeAddress);
        RangeReadAssertions.Number(9d, last.Values[0][0]);
        Assert.False(last.HasMore);
        Assert.Null(last.NextRowOffset);
        Assert.Null(last.NextColumnOffset);
        RangeReadAssertions.Number(1d, RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheet, "C3")).Values[0][0]);
    }

    [Fact]
    public async Task OversizedReads_FailBeforeReturningPartialValues()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [["preserve"]]));
        foreach (var command in new[] { "range.get-values", "range.get-formulas", "rangeformat.get-format" })
        {
            var response = await _fixture.SendForFailureAsync(command, new { sheetName, rangeAddress = "A:XFD" });
            Assert.Equal("InvalidInput", response.ErrorCategory);
            Assert.Contains("limit", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        }
        Assert.Equal("preserve", RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A1")).Values[0][0]);
    }

    [Fact]
    public async Task Bounds_CanInspectFormattingInflatedUsedRangeWhileEagerReadIsRejected()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [["keep"]]));
        RequireSuccess(_commands.SetNumberFormat(_fixture.BatchToken, sheetName, "XFD1048576", "0.00"));
        var bounds = RequireSuccess(_commands.GetUsedRangeInfo(_fixture.BatchToken, sheetName));
        Assert.Equal((1_048_576, 16_384, 17_179_869_184L), (bounds.RowCount, bounds.ColumnCount, bounds.CellCount));
        var response = await _fixture.SendForFailureAsync("range.get-used-range", new { sheetName });
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Contains("limit", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        var page = RequireSuccess(_commands.GetValuesPage(_fixture.BatchToken, sheetName, "A1:XFD1048576", 1_048_575, 16_383, 1, 1));
        Assert.Equal("$XFD$1048576", page.RangeAddress);
        Assert.Null(page.Values[0][0]);
        Assert.False(page.HasMore);
        Assert.Equal("keep", RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A1")).Values[0][0]);
    }

    [Fact]
    public async Task Pages_RejectInvalidOffsetsAndDisjointScopesWithoutChangingContent()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:B1", [[1, 2]]));
        var outside = await _fixture.SendForFailureAsync("range.get-values-page",
            new { sheetName, rangeAddress = "A1:B1", rowOffset = 1, pageRows = 1, pageColumns = 1 });
        Assert.Equal("InvalidInput", outside.ErrorCategory);
        var disjoint = await _fixture.SendForFailureAsync("range.get-values-page",
            new { sheetName, rangeAddress = "A1,C1", pageRows = 1, pageColumns = 1 });
        Assert.Equal("InvalidInput", disjoint.ErrorCategory);
        Assert.Contains("rectangular", disjoint.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        var values = RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A1:B1"));
        RangeReadAssertions.Number(1d, values.Values[0][0]);
        RangeReadAssertions.Number(2d, values.Values[0][1]);
    }

    [Fact]
    public void FormatSummary_ReportsUniformAndMixedPropertiesWithoutChangingContent()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:B1", [[1, 2]]));
        _fixture.Send("rangeformat.format", new { sheetName, rangeAddresses = (string[])["A1:B1"], formatOptions = new { bold = true } });
        var uniform = RequireSuccess(_commands.GetFormatSummary(_fixture.BatchToken, sheetName, "A1:B1"));
        Assert.Equal(2L, uniform.CellCount);
        Assert.NotNull(uniform.Stored);
        Assert.True(uniform.Stored.Font.Bold);
        Assert.Null(uniform.Displayed);
        _fixture.Send("rangeformat.format", new { sheetName, rangeAddresses = (string[])["B1"], formatOptions = new { bold = false } });
        var mixed = RequireSuccess(_commands.GetFormatSummary(_fixture.BatchToken, sheetName, "A1:B1"));
        Assert.NotNull(mixed.Stored);
        Assert.Null(mixed.Stored.Font.Bold);
        Assert.Contains("font.bold", mixed.Stored.MixedFields);
        var displayed = RequireSuccess(_commands.GetFormatSummary(_fixture.BatchToken, sheetName, "A1", Sbroenne.ExcelMcp.Core.Commands.Range.FormatView.Both));
        Assert.NotNull(displayed.Stored);
        Assert.NotNull(displayed.Displayed);
        Assert.True(displayed.Stored.Font.Bold);
        Assert.True(displayed.Displayed.Font.Bold);
        RangeReadAssertions.Number(2d, RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "B1")).Values[0][0]);
    }

    [Fact]
    public void ErrorReads_PreserveDirtyManualCalculationCacheUntilExplicitCalculation()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var calculation = _fixture.CreateCommands<ICalculationModeCommands>();
        var original = RequireSuccess(calculation.GetSettings(_fixture.BatchToken));
        var failure = Record.Exception(() =>
        {
            RequireSuccess(calculation.SetSettings(_fixture.BatchToken, CalculationMode.Manual));
            RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[0d]]));
            RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B1", [["=1/A1"]]));
            RequireSuccess(calculation.Calculate(_fixture.BatchToken, CalculationScope.Range, sheetName, "B1"));
            Assert.Equal("#DIV/0!", RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "B1")).Values[0][0]);
            RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[1d]], overwritePolicy: OverwritePolicy.Allow));

            // The source now permits 1, but the cached error must survive every read.
            foreach (var read in new Func<Sbroenne.ExcelMcp.Core.Models.RangeValueResult>[]
            {
                () => _commands.GetValues(_fixture.BatchToken, sheetName, "A1:B1"),
                () => _commands.GetUsedRange(_fixture.BatchToken, sheetName),
                () => _commands.GetCurrentRegion(_fixture.BatchToken, sheetName, "A1"),
                () => _commands.GetValuesPage(_fixture.BatchToken, sheetName, "A1:B1", 0, 0, 1, 2)
            })
            {
                var result = RequireSuccess(read());
                RangeReadAssertions.Number(1d, result.Values[0][0]);
                Assert.Equal("#DIV/0!", result.Values[0][1]);
                var error = Assert.Single(result.CellErrors);
                Assert.Equal("B1", error.CellAddress);
                Assert.Equal("=1/A1", error.Formula);
                Assert.Equal((int)CalculationMode.Manual, RequireSuccess(calculation.GetSettings(_fixture.BatchToken)).ModeValue);
            }
            var formulas = RequireSuccess(_commands.GetFormulas(_fixture.BatchToken, sheetName, "A1:B1"));
            Assert.Equal("=1/A1", formulas.Formulas[0][1]);
            RangeReadAssertions.Number(1d, formulas.Values[0][0]);
            Assert.Equal("#DIV/0!", formulas.Values[0][1]);
            Assert.Equal("B1", Assert.Single(formulas.CellErrors).CellAddress);
            Assert.Equal((int)CalculationMode.Manual, RequireSuccess(calculation.GetSettings(_fixture.BatchToken)).ModeValue);

            RequireSuccess(calculation.Calculate(_fixture.BatchToken, CalculationScope.Range, sheetName, "B1"));
            var refreshed = RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A1:B1"));
            RangeReadAssertions.Number(1d, refreshed.Values[0][0]);
            RangeReadAssertions.Number(1d, refreshed.Values[0][1]);
            Assert.Empty(refreshed.CellErrors);
        });
        var restoration = Record.Exception(() => RequireSuccess(calculation.SetSettings(
            _fixture.BatchToken, (CalculationMode)original.ModeValue)));
        if (failure is not null && restoration is not null)
            throw new AggregateException("Cache verification and calculation-mode restoration failed.", failure, restoration);
        if (failure is not null) ExceptionDispatchInfo.Capture(failure).Throw();
        if (restoration is not null) ExceptionDispatchInfo.Capture(restoration).Throw();
    }

    [Fact]
    public void FormatSummary_MixedConditionalFormattingReportsDisplayedDifferenceOnly()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:B1", [[1d, 2d]]));
        _fixture.Send("rangeformat.format", new
        {
            sheetName,
            rangeAddresses = (string[])["A1:B1"],
            formatOptions = new { bold = false }
        });
        _fixture.Send("conditionalformat.add-rule", new
        {
            sheetName,
            rangeAddress = "A1:B1",
            ruleType = "expression",
            formula1 = "=A1=1",
            fontBold = true
        });

        var summary = RequireSuccess(_commands.GetFormatSummary(_fixture.BatchToken, sheetName, "A1:B1", FormatView.Both));
        Assert.Equal(2L, summary.CellCount);
        Assert.NotNull(summary.Stored);
        Assert.NotNull(summary.Displayed);
        Assert.False(summary.Stored.Font.Bold);
        Assert.DoesNotContain("font.bold", summary.Stored.MixedFields);
        Assert.Null(summary.Displayed.Font.Bold);
        Assert.Contains("font.bold", summary.Displayed.MixedFields);
        Assert.Equal(["font.bold"], summary.Displayed.MixedFields.Except(summary.Stored.MixedFields).Order(StringComparer.Ordinal));
        var first = RequireSuccess(_commands.GetFormatSummary(_fixture.BatchToken, sheetName, "A1", FormatView.Both));
        var second = RequireSuccess(_commands.GetFormatSummary(_fixture.BatchToken, sheetName, "B1", FormatView.Both));
        Assert.False(first.Stored!.Font.Bold);
        Assert.True(first.Displayed!.Font.Bold);
        Assert.False(second.Stored!.Font.Bold);
        Assert.False(second.Displayed!.Font.Bold);
        Assert.DoesNotContain("font.bold", first.Displayed.MixedFields);
        Assert.DoesNotContain("font.bold", second.Displayed.MixedFields);
        var values = RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A1:B1"));
        RangeReadAssertions.Number(1d, values.Values[0][0]);
        RangeReadAssertions.Number(2d, values.Values[0][1]);
    }
}
