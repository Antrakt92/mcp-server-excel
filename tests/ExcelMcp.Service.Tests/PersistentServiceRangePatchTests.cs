using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceRangePatchTests(PersistentServiceWorkbookFixture fixture)
    : PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void ConflictPreventsEarlierWriteAndSuccessfulPatchPreservesManualException()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, sheetName, "A1:A3", [["=1"], [42], ["=3"]]));

        var conflict = SendPatch(sheetName, "=99");
        Assert.False(conflict.GetProperty("success").GetBoolean());
        Assert.Equal("A3", conflict.GetProperty("conflictAddresses")[0].GetString());
        Assert.Empty(conflict.GetProperty("confirmedAddresses").EnumerateArray());
        Assert.Equal("=1", RequireSuccess(_commands.GetFormulas(_fixture.BatchToken, sheetName, "A1")).Formulas[0][0]);

        var applied = SendPatch(sheetName, "=3");
        Assert.True(applied.GetProperty("success").GetBoolean(), applied.ToString());
        Assert.Equal(["A1", "A3"], applied.GetProperty("confirmedAddresses").EnumerateArray().Select(cell => cell.GetString()));
        var result = RequireSuccess(_commands.GetFormulas(_fixture.BatchToken, sheetName, "A1:A3"));
        Assert.Equal("=2", result.Formulas[0][0]);
        RangeReadAssertions.Number(42, result.Values[1][0]);
        Assert.Equal("00123", result.Values[2][0]);
        Assert.True(string.IsNullOrEmpty(result.Formulas[2][0]));
    }

    [Fact]
    public void LiteralFormulaTextAndBlankAreDistinctFromFormulaAndZero()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "B2:B4", [[false], [0], ["'=1"]]));

        using var options = JsonDocument.Parse("""
            {"edits":[
              {"address":"B2","expected":{"value":false},"replacement":{"value":true}},
              {"address":"B3","expected":{"value":0},"replacement":{"value":null}},
              {"address":"B4","expected":{"value":"=1"},"replacement":{"value":"=2"}}
            ]}
            """);
        var result = _fixture.Send("range.patch-cells", new
        {
            sheetName,
            patchOptions = options.RootElement
        });

        using var json = JsonDocument.Parse(result.Result!);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean(), result.Result);
        var readback = RequireSuccess(_commands.GetFormulas(_fixture.BatchToken, sheetName, "B2:B4"));
        Assert.Equal(true, readback.Values[0][0]);
        Assert.Null(readback.Values[1][0]);
        Assert.Equal("=2", readback.Values[2][0]);
        Assert.True(string.IsNullOrEmpty(readback.Formulas[2][0]));
    }

    [Fact]
    public void ScalarReplacementConfirmsExcelImplicitIntersectionWithoutChangingAdjacentCells()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        bool supportsFormula2 = _fixture.ExecuteRawVerification((context, _) => context.Capabilities.SupportsFormula2);
        RequireSuccess(_commands.SetValues(batch, sheetName, "A1:A4", [[1], [4], [9], [16]]));
        RequireSuccess(_commands.SetValues(batch, sheetName, "B3", [[91]]));

        using var options = JsonDocument.Parse("""
            {"edits":[
              {"address":"B2","expected":{"value":null},"replacement":{"formula":"=SQRT(A1:A4)"}},
              {"address":"C2","expected":{"value":null},"replacement":{"value":17}}
            ]}
            """);
        var response = _fixture.Send("range.patch-cells", new { sheetName, patchOptions = options.RootElement });
        Assert.NotNull(response.Result);
        using var result = JsonDocument.Parse(response.Result);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean(), response.Result);
        Assert.Equal(["B2", "C2"], result.RootElement.GetProperty("confirmedAddresses").EnumerateArray().Select(cell => cell.GetString()));

        var readback = RequireSuccess(_commands.GetFormulas(batch, sheetName, "B2:C3"));
        Assert.Equal(supportsFormula2 ? "=SQRT(@A1:A4)" : "=SQRT(A1:A4)", readback.Formulas[0][0]);
        RangeReadAssertions.Number(2, readback.Values[0][0]);
        RangeReadAssertions.Number(17, readback.Values[0][1]);
        RangeReadAssertions.Number(91, readback.Values[1][0]);
        Assert.Null(readback.Values[1][1]);
    }

    [Fact]
    public async Task OversizedLiteralIsRejectedBeforeAnyPatchWrite()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var edits = new[]
        {
            new { address = "A1", expected = new { value = (string?)null }, replacement = new { value = "First" } },
            new { address = "B1", expected = new { value = (string?)null }, replacement = new { value = new string('x', 32767) } }
        };
        var result = await _fixture.SendForFailureAsync("range.patch-cells", new { sheetName, patchOptions = new { edits } });
        Assert.Equal("InvalidInput", result.ErrorCategory);
        Assert.All(RequireSuccess(_commands.GetValues(batch, sheetName, "A1:B1")).Values[0], value => Assert.Null(value));
    }

    [Fact]
    public async Task LaterInvalidFormulaReportsEarlierNativeWriteAndLeavesNextCellUnchanged()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "C1:C3", [[1], [2], [3]]));
        using var options = JsonDocument.Parse("""
            {"edits":[
              {"address":"C1","expected":{"value":1},"replacement":{"value":11}},
              {"address":"C2","expected":{"value":2},"replacement":{"formula":"=SUM("}},
              {"address":"C3","expected":{"value":3},"replacement":{"value":33}}
            ]}
            """);
        var response = await _fixture.SendForFailureAsync("range.patch-cells", new { sheetName, patchOptions = options.RootElement });
        using var state = JsonDocument.Parse(response.PartialState!);
        Assert.Equal("C2", state.RootElement.GetProperty("unconfirmedAddress").GetString());
        Assert.Equal("C1", state.RootElement.GetProperty("confirmedAddresses")[0].GetString());
        var values = RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "C1:C3"));
        RangeReadAssertions.Number(11, values.Values[0][0]);
        RangeReadAssertions.Number(3, values.Values[2][0]);
    }

    [Fact]
    public async Task TableTargetStopsPatchBeforeTouchingOrdinaryCell()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:A3", [["Qty"], [2], [3]]));
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "C1", [[1]]));
        _fixture.Send("table.create", new { sheetName, tableName = "PatchGuardTable", rangeAddress = "A1:A3" });
        using var options = JsonDocument.Parse("""
            {"edits":[
              {"address":"C1","expected":{"value":1},"replacement":{"value":11}},
              {"address":"A2","expected":{"value":2},"replacement":{"formula":"=20"}}
            ]}
            """);
        var response = await _fixture.SendForFailureAsync("range.patch-cells", new { sheetName, patchOptions = options.RootElement });
        Assert.Contains("A2", response.ErrorMessage, StringComparison.Ordinal);
        RangeReadAssertions.Number(1, RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "C1")).Values[0][0]);
        RangeReadAssertions.Number(3, RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A3")).Values[0][0]);
    }

    private JsonElement SendPatch(string sheetName, string expected)
    {
        var response = _fixture.Send("range.patch-cells", new
        {
            sheetName,
            patchOptions = new
            {
                edits = new object[]
                {
                    new { address = "A1", expected = new { formula = "=1" }, replacement = new { formula = "=2" } },
                    new { address = "$A$3", expected = new { formula = expected }, replacement = new { value = "00123" } }
                }
            }
        });
        Assert.NotNull(response.Result);
        using var json = JsonDocument.Parse(response.Result);
        return json.RootElement.Clone();
    }
}
