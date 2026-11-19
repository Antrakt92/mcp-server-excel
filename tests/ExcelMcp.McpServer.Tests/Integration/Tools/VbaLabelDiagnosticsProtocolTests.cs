using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "VBA")]
[Trait("RequiresExcel", "false")]
public sealed class VbaLabelDiagnosticsProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task List_OptInAdvertisesAndForwardsLabelDiagnostics()
    {
        var tools = await fixture.ListToolsAsync();
        var tool = Assert.Single(tools, candidate => candidate.Name == "vba_read");
        var property = tool.JsonSchema.GetProperty("properties").GetProperty("include_label_warnings");
        Assert.Equal(["boolean", "null"], property.GetProperty("type").EnumerateArray()
            .Select(value => value.GetString()).Order(StringComparer.Ordinal));
        Assert.Contains("list only", property.GetProperty("description").GetString(), StringComparison.OrdinalIgnoreCase);

        var call = await fixture.CallToolAsync("vba_read", new()
        {
            ["action"] = "list",
            ["workbook_session_id"] = "session-1",
            ["include_label_warnings"] = true
        }, RecordingToolTest.Success("""
            {"success":true,"scripts":[],"labelDiagnostics":{"warnings":[{"moduleName":"Module1","procedureName":"Probe","line":4,"label":"Reset"}]}}
            """), "vba.list", """{"includeLabelWarnings":true}""");

        Assert.False(call.Result.IsError);
        using var result = JsonDocument.Parse(call.JsonResult);
        var diagnostic = result.RootElement.GetProperty("labelDiagnostics");
        var warning = Assert.Single(diagnostic.GetProperty("warnings").EnumerateArray());
        Assert.Equal("Module1", warning.GetProperty("moduleName").GetString());
        Assert.Equal("Reset", warning.GetProperty("label").GetString());
        Assert.Equal(4, warning.GetProperty("line").GetInt32());
    }
}
