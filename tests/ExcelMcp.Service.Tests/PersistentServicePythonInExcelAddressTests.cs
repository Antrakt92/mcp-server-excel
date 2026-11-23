using Sbroenne.ExcelMcp.Core.Commands.PythonInExcel;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PythonInExcel")]
[Trait("Speed", "Medium")]
public sealed class PersistentServicePythonInExcelAddressTests(PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IPythonInExcelCommands _pythonCommands = fixture.CreateCommands<IPythonInExcelCommands>();

    [Theory]
    [InlineData("A1:B1")]
    [InlineData("A1,B1")]
    public async Task GetResult_OnMultipleCells_RejectsTheAddressWithoutChangingValues(string range)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheet, "A1:B1", [["Left", "Right"]]));

        // Plain values keep this test independent of licensed or cloud Python in Excel.
        var response = await _fixture.SendForFailureAsync("pythoninexcel.get-result", new
        {
            sheetName = sheet,
            rangeAddress = range
        });

        Assert.Equal(new object?[] { "Left", "Right" },
            Assert.Single(RequireSuccess(_commands.GetValues(batch, sheet, "A1:B1")).Values));
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Contains("exactly one cell", response.ErrorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void GetResult_OnOneNonPythonCell_RetainsTheFormulaTypeDiagnosis()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheet, "A1", [["keep"]]));
        var result = _pythonCommands.GetResult(batch, sheet, "$A$1");

        Assert.False(result.Success);
        Assert.Equal("$A$1", result.RangeAddress);
        Assert.Contains("does not contain a Python in Excel", result.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal("keep", Assert.Single(Assert.Single(RequireSuccess(_commands.GetValues(batch, sheet, "A1")).Values)));
    }
}
