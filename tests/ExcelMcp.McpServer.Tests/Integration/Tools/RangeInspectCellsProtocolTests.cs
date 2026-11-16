using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class RangeInspectCellsProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task InspectCells_UsesReadToolAndPreservesTypedValuesAndNativeText()
    {
        const string response =
            """{"success":true,"sheetName":"Sheet1","rangeAddress":"$A$1:$C$1","cells":[{"address":"$A$1","value":123,"displayedText":"00123"},{"address":"$B$1","value":false,"formula":"=FALSE()","displayedText":"FALSE"},{"address":"$C$1","value":null,"displayedText":""}]}""";
        var call = await fixture.CallToolAsync("range_read", new()
        {
            ["action"] = "inspect-cells",
            ["workbook_session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1:C1"
        }, RecordingToolTest.Success(response), "range.inspect-cells",
            """{"sheetName":"Sheet1","rangeAddress":"A1:C1"}""");
        Assert.False(call.Result.IsError);
        using var actual = JsonDocument.Parse(call.JsonResult);
        using var expected = JsonDocument.Parse(response);
        Assert.True(JsonElement.DeepEquals(expected.RootElement, actual.RootElement));
    }
}
