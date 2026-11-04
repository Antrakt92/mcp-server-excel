using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "VBA")]
[Trait("RequiresExcel", "false")]
public sealed class VbaScalarArgumentProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task Invoke_ForwardsScalarKindsNullAndMissingWithoutCoercion()
    {
        var tools = await fixture.ListToolsAsync();
        var tool = Assert.Single(tools, candidate => candidate.Name == "vba");
        var items = tool.JsonSchema.GetProperty("properties")
            .GetProperty("arguments").GetProperty("items");
        // The SDK's unrestricted value schema must not narrow numbers or booleans to strings.
        Assert.Equal(JsonValueKind.True, items.GetProperty("properties").GetProperty("value").ValueKind);

        var call = await fixture.CallToolAsync("vba", new()
        {
            ["action"] = "invoke",
            ["workbook_session_id"] = "session-1",
            ["procedure_name"] = "ScalarProbe",
            ["arguments"] = new object[]
            {
                new { type = "number", value = -12.5 },
                new { type = "boolean", value = false },
                new { type = "string", value = "=literal" },
                new { type = "null" },
                new { type = "missing" }
            }
        }, RecordingToolTest.Success("""{"success":true,"returnKind":"number","value":-12.5}"""),
            "vba.invoke",
            """
            {"procedureName":"ScalarProbe","arguments":[{"type":"number","value":-12.5},{"type":"boolean","value":false},{"type":"string","value":"=literal"},{"type":"null"},{"type":"missing"}]}
            """);

        Assert.False(call.Result.IsError);
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.Equal("number", result.RootElement.GetProperty("returnKind").GetString());
        Assert.Equal(-12.5, result.RootElement.GetProperty("value").GetDouble());
    }

    [Theory]
    [InlineData("number", "12")]
    [InlineData("boolean", "true")]
    public async Task Invoke_DoesNotCoerceQuotedScalarsBeforeCoreValidation(string type, string value)
    {
        var expectedArgs = JsonSerializer.Serialize(new
        {
            procedureName = "ScalarProbe",
            arguments = new[] { new { type, value } }
        });
        var call = await fixture.CallToolAsync("vba", new()
        {
            ["action"] = "invoke",
            ["workbook_session_id"] = "session-1",
            ["procedure_name"] = "ScalarProbe",
            ["arguments"] = new[] { new { type, value } }
        }, new ServiceResponse
        {
            Success = false,
            ErrorCategory = "InvalidInput",
            ErrorMessage = "The scalar value does not match its declared VBA type."
        }, "vba.invoke", expectedArgs);

        Assert.True(call.Result.IsError);
        using var sent = JsonDocument.Parse(call.Request.Args!);
        var scalar = sent.RootElement.GetProperty("arguments")[0].GetProperty("value");
        Assert.Equal(JsonValueKind.String, scalar.ValueKind);
        Assert.Equal(value, scalar.GetString());
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.False(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("InvalidInput", result.RootElement.GetProperty("errorCategory").GetString());
    }
}
