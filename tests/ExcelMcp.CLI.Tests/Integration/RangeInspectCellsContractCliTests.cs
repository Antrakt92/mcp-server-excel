using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class RangeInspectCellsContractCliTests
{
    [Fact]
    public async Task InspectCells_ForwardsRangeAndPreservesTheTypedResult()
    {
        ServiceRequest? captured = null;
        const string response =
            """{"success":true,"cells":[{"address":"$A$1","value":0,"displayedText":"00000"},{"address":"$B$1","formula":"=1/0","value":"#DIV/0!","displayedText":"#DIV/0!","excelErrorCode":-2146826281},{"address":"$C$1","value":null,"displayedText":""}]}""";
        var result = await InProcessCliHelper.RunAsync(
            ["range", "inspect-cells", "--session", "session-1", "--sheet-name", "Sheet1", "--range-address", "A1:C1"],
            request =>
            {
                captured = request;
                return new ServiceResponse { Success = true, Result = response };
            });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("range.inspect-cells", captured.Command);
        Assert.Equal("session-1", captured.SessionId);
        using var args = JsonDocument.Parse(captured.Args!);
        using var expectedArgs = JsonDocument.Parse("""{"sheetName":"Sheet1","rangeAddress":"A1:C1"}""");
        Assert.True(JsonElement.DeepEquals(expectedArgs.RootElement, args.RootElement));
        using var output = JsonDocument.Parse(result.Stdout);
        using var expected = JsonDocument.Parse(response);
        Assert.True(JsonElement.DeepEquals(expected.RootElement.GetProperty("cells"),
            output.RootElement.GetProperty("data").GetProperty("cells")));
    }
}
