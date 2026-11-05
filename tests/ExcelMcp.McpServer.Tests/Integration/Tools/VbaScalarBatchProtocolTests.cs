using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "VBA")]
[Trait("RequiresExcel", "false")]
public sealed class VbaScalarBatchProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task Run_InvokePreservesNativeJsonScalarKindsAndFalseReturn()
    {
        using var commands = JsonDocument.Parse(
            """
            [{
              "command":"vba.invoke",
              "args":{
                "procedureName":"ScalarProbe",
                "arguments":[
                  {"type":"number","value":-12.5},
                  {"type":"boolean","value":false},
                  {"type":"string","value":"=literal"},
                  {"type":"null","value":null},
                  {"type":"missing"}
                ]
              }
            }]
            """);

        var call = await fixture.CallSequenceAsync("batch", new()
        {
            ["action"] = "run",
            ["workbook_session_id"] = "recording-session",
            ["stop_on_error"] = true,
            ["commands"] = commands.RootElement.Clone()
        }, RecordingToolTest.Success("""{"success":true,"returnKind":"boolean","value":false}"""));

        RecordingToolTest.AssertRequest(
            Assert.Single(call.Requests),
            "vba.invoke",
            "recording-session",
            """
            {"procedureName":"ScalarProbe","arguments":[{"type":"number","value":-12.5},{"type":"boolean","value":false},{"type":"string","value":"=literal"},{"type":"null","value":null},{"type":"missing"}]}
            """);

        Assert.False(call.Result.IsError == true);
        using var output = JsonDocument.Parse(
            call.Result.StructuredContent!.Value.GetRawText());
        var root = output.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean());
        Assert.Equal(1, root.GetProperty("totalCount").GetInt32());
        Assert.Equal(1, root.GetProperty("executedCount").GetInt32());
        Assert.False(root.GetProperty("stoppedEarly").GetBoolean());
        var result = Assert.Single(root.GetProperty("results").EnumerateArray());
        Assert.Equal(0, result.GetProperty("index").GetInt32());
        Assert.Equal("vba.invoke", result.GetProperty("command").GetString());
        Assert.True(result.GetProperty("success").GetBoolean());
        Assert.Equal("boolean", result.GetProperty("returnKind").GetString());
        Assert.False(result.GetProperty("value").GetBoolean());
    }
}
