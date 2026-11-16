using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class RangeMergeProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Merge_OnlyForwardsDiscardPermissionWhenSupplied(bool allow)
    {
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = "merge-cells",
            ["workbook_session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1:B1"
        };
        if (allow)
            arguments["overwrite_policy"] = "allow";
        var expected = allow
            ? """{"sheetName":"Sheet1","rangeAddress":"A1:B1","overwritePolicy":"allow"}"""
            : """{"sheetName":"Sheet1","rangeAddress":"A1:B1"}""";

        var call = await fixture.CallToolAsync("range_format", arguments,
            RecordingToolTest.Success("""{"success":true}"""), "rangeformat.merge-cells", expected);

        Assert.False(call.Result.IsError);
    }
}
