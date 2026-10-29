using Xunit;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceRangeFormulaValidationTests
{
    private static readonly string[][] OneFormula = [["=1+1"]];
    private static readonly string[] ExpectedUnverifiedChecks =
    [
        "excel-compilation", "full-formula-grammar", "function-and-udf-availability",
        "external-references", "defined-names", "circular-references", "overwrite-and-write-eligibility"
    ];

    [Theory]
    [InlineData("A1:garbage", 1, 1)]
    [InlineData("XFD1048576", 1, 2)]
    [InlineData("A1048576", 2, 1)]
    [InlineData("A1:B2", 1, 2)]
    [InlineData("A1:B1", 1, 1)]
    public async Task ValidateFormulas_RejectsInvalidTargetGeometryBeforeReportingValid(string scope, int rows, int columns)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheet, "A1", [["guard"]]));
        var formulas = Enumerable.Range(0, rows).Select(_ => Enumerable.Repeat("=1+1", columns).ToList()).ToList();
        var response = await _fixture.SendForFailureAsync("range.validate-formulas", new { sheetName = sheet, rangeAddress = scope, formulas });
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Null(response.Result);
        AssertCellUnchanged(batch, sheet, "A1", "guard", null);
        Assert.Null(RequireSuccess(_commands.GetValues(batch, sheet, "XFD1048576")).Values[0][0]);
    }

    [Fact]
    public async Task ValidateFormulas_MissingTargetSheetCannotReportValid()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheet, "A1", [["guard"]]));
        var response = await _fixture.SendForFailureAsync("range.validate-formulas", new
        {
            sheetName = "Missing_" + Guid.NewGuid().ToString("N")[..8],
            rangeAddress = "A1",
            formulas = OneFormula
        });
        Assert.Equal("NotFound", response.ErrorCategory);
        Assert.Null(response.Result);
        AssertCellUnchanged(batch, sheet, "A1", "guard", null);
    }

    [Fact]
    public void ValidateFormulas_ExistingQuotedSheetAndLiteralParentheses_AreValidAndReadOnly()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        string quotedSheet = _fixture.CreateNamedTestSheet(batch, "O'Brien (" + Guid.NewGuid().ToString("N")[..8] + ")");
        RequireSuccess(_commands.SetValues(batch, sheet, "B2:D2", [["guard-a", "guard-b", "guard-c"]]));
        string reference = "='" + quotedSheet.Replace("'", "''", StringComparison.Ordinal) + "'!$A$1";
        var result = RequireSuccess(_commands.ValidateFormulas(batch, sheet, "$b$2:$d$2",
            [["=\"(\"", reference, "=" + sheet + "!A1"]]));
        var response = _fixture.Send("range.validate-formulas", new
        {
            sheetName = sheet,
            rangeAddress = "B2",
            formulas = OneFormula
        });
        using var json = JsonDocument.Parse(response.Result!);
        Assert.Equal("lexical-and-local-sheet-references", json.RootElement.GetProperty("validationScope").GetString());
        Assert.Equal(ExpectedUnverifiedChecks,
            json.RootElement.GetProperty("unverifiedChecks").EnumerateArray().Select(value => value.GetString()!).ToArray());
        Assert.True(result.IsValid);
        Assert.Equal(3, result.ValidCount);
        Assert.Equal(0, result.ErrorCount);
        Assert.Null(result.Errors);
        var values = RequireSuccess(_commands.GetValues(batch, sheet, "B2:D2"));
        Assert.Equal(["guard-a", "guard-b", "guard-c"], values.Values[0]);
        Assert.All(RequireSuccess(_commands.GetFormulas(batch, sheet, "B2:D2")).Formulas[0], formula => Assert.Equal("", formula));
    }

    [Fact]
    public void ValidateFormulas_MissingQuotedSheet_ReportsCorrectAbsoluteTargetWithoutWriting()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheet, "C4", [["guard"]]));
        var result = RequireSuccess(_commands.ValidateFormulas(batch, sheet, "$C$4", [["='Missing Sheet (guard)'!$A$1"]]));
        Assert.False(result.IsValid);
        Assert.Equal(0, result.ValidCount);
        var error = Assert.Single(result.Errors!);
        Assert.Equal("invalid-reference", error.Category);
        Assert.Equal("C4", error.CellAddress);
        Assert.Equal((4, 3), (error.Row, error.Column));
        AssertCellUnchanged(batch, sheet, "C4", "guard", null);
    }
}
