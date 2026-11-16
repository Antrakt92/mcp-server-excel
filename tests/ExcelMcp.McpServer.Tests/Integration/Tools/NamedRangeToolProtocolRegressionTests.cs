using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Parameters")]
[Trait("RequiresExcel", "false")]
public sealed class NamedRangeToolProtocolRegressionTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task NamedRangeList_ReturnsStructuredResult_AndSessionRemainsUsable()
    {
        const string sessionId = "recording-session";
        var listCall = await _fixture.CallToolAsync(
            "namedrange_read",
            new Dictionary<string, object?>
            {
                ["action"] = "list",
                ["workbook_session_id"] = sessionId
            },
            RecordingToolTest.Success(
                """{"success":true,"namedRanges":[{"name":"CsvFolder","refersTo":"=Sheet1!$B$4","value":"C:\\Data"}]}"""),
            "namedrange.list",
            null);

        Assert.Equal("namedrange.list", listCall.Request.Command);
        Assert.Equal(sessionId, listCall.Request.SessionId);
        using (var result = JsonDocument.Parse(listCall.JsonResult))
        {
            var range = Assert.Single(
                result.RootElement.GetProperty("namedRanges").EnumerateArray());
            Assert.Equal("CsvFolder", range.GetProperty("name").GetString());
            Assert.Contains(
                "$B$4",
                range.GetProperty("refersTo").GetString(),
                StringComparison.OrdinalIgnoreCase);
            Assert.Equal("C:\\Data", range.GetProperty("value").GetString());
        }

        var worksheetCall = await _fixture.CallToolAsync(
            "worksheet_read",
            new Dictionary<string, object?>
            {
                ["action"] = "list",
                ["workbook_session_id"] = sessionId
            },
            RecordingToolTest.Success(
                """{"success":true,"worksheets":[{"name":"Sheet1"}]}"""),
            "sheet.list",
            "{}");

        Assert.Equal("sheet.list", worksheetCall.Request.Command);
        using var worksheetResult = JsonDocument.Parse(worksheetCall.JsonResult);
        Assert.Equal(
            "Sheet1",
            worksheetResult.RootElement.GetProperty("worksheets")[0]
                .GetProperty("name").GetString());
    }

    [Fact]
    public async Task NamedRangeCreate_ForwardsExplicitHiddenVisibility()
    {
        var call = await _fixture.CallToolAsync(
            "namedrange",
            new Dictionary<string, object?>
            {
                ["action"] = "create",
                ["workbook_session_id"] = "recording-session",
                ["name"] = "Marker",
                ["reference"] = "Sheet1!$A$1",
                ["visible"] = false
            },
            RecordingToolTest.Success("""{"success":true}"""),
            "namedrange.create",
            """{"name":"Marker","reference":"Sheet1!$A$1","visible":false}""");

        Assert.Equal("recording-session", call.Request.SessionId);
    }

    [Fact]
    public async Task NamedRangeRead_ForwardsMetadataFlagAndPreservesUnavailableResolution()
    {
        const string response = """
            {"name":"Marker","refersTo":"=#REF!","valueType":"NotRead","visible":false,"resolutionStatus":"Unavailable","resolutionHResult":-2146827284}
            """;
        var (call, requests) = await _fixture.CallSequenceAsync(
            "namedrange_read",
            new Dictionary<string, object?>
            {
                ["action"] = "read",
                ["workbook_session_id"] = "recording-session",
                ["name"] = "Marker",
                ["metadata_only"] = true
            },
            RecordingToolTest.Success(response));

        RecordingToolTest.AssertRequest(Assert.Single(requests), "namedrange.read",
            "recording-session", """{"name":"Marker","metadataOnly":true}""");
        using var result = JsonDocument.Parse(McpResponseAssertions.ReadText(call, expectedError: false));
        Assert.Equal("=#REF!", result.RootElement.GetProperty("refersTo").GetString());
        Assert.False(result.RootElement.GetProperty("visible").GetBoolean());
        Assert.Equal("Unavailable", result.RootElement.GetProperty("resolutionStatus").GetString());
        Assert.Equal(unchecked((int)0x800A03EC), result.RootElement.GetProperty("resolutionHResult").GetInt32());
        Assert.False(result.RootElement.TryGetProperty("value", out _));
    }
}
