using System.Text.Json;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "VBA")]
[Trait("RequiresExcel", "false")]
public sealed class VbaProjectProtectionProtocolTests(RecordingProgramTransportFixture fixture)
{
    private const string AppliedJson = """{"success":true,"projectName":"SyntheticProject","currentProtection":"None","configurationApplied":true,"saveRequired":true,"reopenRequired":true}""";

    [Theory]
    [InlineData("  Synthetic leading and trailing !7  ")]
    [InlineData("Synthetic\"quote\\backslash!7")]
    [InlineData("Synthetic-Gr\u00fc\u00dfe-\u0416\u6f22\u5b57!7")]
    public async Task ProtectProject_TransportsExactPasswordAndPendingSaveResult(string password)
    {
        var call = await fixture.CallToolAsync("vba", new()
        {
            ["action"] = "protect-project",
            ["workbook_session_id"] = "session-1",
            ["password"] = password
        }, RecordingToolTest.Success(AppliedJson), "vba.protect-project",
            JsonSerializer.Serialize(new { password }));

        Assert.False(call.Result.IsError);
        using var args = JsonDocument.Parse(call.Request.Args!);
        Assert.Equal(password, args.RootElement.GetProperty("password").GetString());
        using var result = JsonDocument.Parse(call.JsonResult);
        var root = result.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean());
        Assert.Equal("SyntheticProject", root.GetProperty("projectName").GetString());
        Assert.Equal("None", root.GetProperty("currentProtection").GetString());
        Assert.True(root.GetProperty("configurationApplied").GetBoolean());
        Assert.True(root.GetProperty("saveRequired").GetBoolean());
        Assert.True(root.GetProperty("reopenRequired").GetBoolean());
        AssertSecretAbsent(password, call.JsonResult);
        Assert.NotNull(call.Result.StructuredContent);
        Assert.True(JsonElement.DeepEquals(root, call.Result.StructuredContent!.Value));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task MissingPassword_IsRefusedBeforeDispatch(bool explicitNull)
    {
        Dictionary<string, object?> arguments = new()
        {
            ["action"] = "protect-project",
            ["workbook_session_id"] = "session-1"
        };
        if (explicitNull) arguments["password"] = null;

        var result = await fixture.CallResultWithoutDispatchAsync("vba", arguments);
        Assert.True(result.IsError);
        var text = Assert.IsType<TextContentBlock>(Assert.Single(result.Content)).Text;
        Assert.Contains("password", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task UnconfirmedProtection_PreservesErrorAndPartialStateWithoutEchoingPassword()
    {
        const string password = "Synthetic\"Unconfirmed-Gr\u00fc\u00dfe!7";
        var call = await fixture.CallToolAsync("vba", new()
        {
            ["action"] = "protect-project",
            ["workbook_session_id"] = "session-1",
            ["password"] = password
        }, new ServiceResponse
        {
            Success = false,
            ErrorCategory = "Timeout",
            ErrorMessage = "The protection outcome is unconfirmed.",
            PartialState = """{"phase":"acceptance-requested","acceptanceRequested":true,"dialogCleanupConfirmed":false,"rollbackAttempted":false}"""
        }, "vba.protect-project", JsonSerializer.Serialize(new { password }));

        Assert.True(call.Result.IsError);
        using var result = JsonDocument.Parse(call.JsonResult);
        var root = result.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal("Timeout", root.GetProperty("errorCategory").GetString());
        Assert.Equal("The protection outcome is unconfirmed.", root.GetProperty("errorMessage").GetString());
        var state = root.GetProperty("partialState");
        Assert.Equal("acceptance-requested", state.GetProperty("phase").GetString());
        Assert.True(state.GetProperty("acceptanceRequested").GetBoolean());
        Assert.False(state.GetProperty("dialogCleanupConfirmed").GetBoolean());
        Assert.False(state.GetProperty("rollbackAttempted").GetBoolean());
        Assert.False(state.TryGetProperty("configurationApplied", out _));
        AssertSecretAbsent(password, call.JsonResult);
    }

    private static void AssertSecretAbsent(string password, string output)
    {
        Assert.DoesNotContain(password, output, StringComparison.Ordinal);
        Assert.DoesNotContain(JsonEncodedText.Encode(password).ToString(), output, StringComparison.Ordinal);
    }
}
