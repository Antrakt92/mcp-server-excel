using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "false")]
public sealed class RangePatchToolProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task PatchRoutesFormulaSeparatelyFromExplicitBlankAndPreservesConflict()
    {
        using var options = JsonDocument.Parse(
            """{"edits":[{"address":"A1","expected":{"formula":"=1"},"replacement":{"value":null}}]}""");
        var call = await fixture.CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "patch-cells",
            ["workbook_session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["patch_options"] = options.RootElement
        }, RecordingToolTest.Success(
            """{"success":false,"errorMessage":"Cell contents differ at A1.","errorCategory":"Conflict","conflictAddresses":["A1"]}"""),
            "range.patch-cells",
            """{"sheetName":"Sheet1","patchOptions":{"edits":[{"address":"A1","expected":{"formula":"=1"},"replacement":{"value":null}}]}}""");

        using var args = JsonDocument.Parse(call.Request.Args!);
        var edit = args.RootElement.GetProperty("patchOptions").GetProperty("edits")[0];
        Assert.Equal("=1", edit.GetProperty("expected").GetProperty("formula").GetString());
        Assert.False(edit.GetProperty("expected").TryGetProperty("value", out _));
        Assert.Equal(JsonValueKind.Null, edit.GetProperty("replacement").GetProperty("value").ValueKind);
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.False(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("A1", result.RootElement.GetProperty("conflictAddresses")[0].GetString());
    }
}
