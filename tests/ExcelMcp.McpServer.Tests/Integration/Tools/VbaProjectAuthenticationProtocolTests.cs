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
public sealed class VbaProjectAuthenticationProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData("unlock-project", true)]
    [InlineData("verify-password", true)]
    [InlineData("verify-password", false)]
    [InlineData("unlock-project", null)]
    public async Task Authentication_TransportsExactSecretAndDistinctOutcome(string action, bool? verified)
    {
        const string password = "  Synthetic\"Exact\\!7  ";
        var response = JsonSerializer.Serialize(new
        {
            success = true,
            projectName = "SyntheticProject",
            protectionBefore = verified.HasValue ? "Locked" : "None",
            protectionAfter = verified is false ? "Locked" : "None",
            authenticationPerformed = verified.HasValue,
            alreadyAccessible = !verified.HasValue,
            passwordVerified = verified,
            accessUnlocked = verified is true,
            protectionSettingsChanged = false,
            saved = false,
            reopened = false
        });
        var call = await fixture.CallToolAsync("vba", new()
        {
            ["action"] = action,
            ["workbook_session_id"] = "session-1",
            ["password"] = password
        }, RecordingToolTest.Success(response), "vba." + action, JsonSerializer.Serialize(new { password }));
        Assert.False(call.Result.IsError);
        using var args = JsonDocument.Parse(call.Request.Args!);
        Assert.Equal(password, args.RootElement.GetProperty("password").GetString());
        using var result = JsonDocument.Parse(call.JsonResult);
        var root = result.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean());
        var observed = root.GetProperty("passwordVerified");
        if (verified.HasValue) Assert.Equal(verified.Value, observed.GetBoolean());
        else Assert.Equal(JsonValueKind.Null, observed.ValueKind);
        Assert.Equal(verified.HasValue, root.GetProperty("authenticationPerformed").GetBoolean());
        Assert.False(root.GetProperty("saved").GetBoolean());
        Assert.False(root.GetProperty("reopened").GetBoolean());
        AssertSecretAbsent(password, call.JsonResult);
        Assert.NotNull(call.Result.StructuredContent);
        Assert.True(JsonElement.DeepEquals(root, call.Result.StructuredContent!.Value));
    }

    [Theory]
    [InlineData("unlock-project")]
    [InlineData("verify-password")]
    public async Task Authentication_MissingPasswordIsRejectedBeforeDispatch(string action)
    {
        var result = await fixture.CallResultWithoutDispatchAsync("vba", new()
        { ["action"] = action, ["workbook_session_id"] = "session-1" });
        Assert.True(result.IsError);
        Assert.Contains("password", Assert.IsType<TextContentBlock>(Assert.Single(result.Content)).Text, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("protect-project")]
    [InlineData("unlock-project")]
    [InlineData("verify-password")]
    public async Task PasswordActions_PreserveWhitespaceOnlyPasswordThroughPublicTool(string action)
    {
        foreach (var password in new[] { " ", "   " })
        {
            var call = await fixture.CallToolAsync("vba", new()
            {
                ["action"] = action,
                ["workbook_session_id"] = "session-1",
                ["password"] = password
            }, RecordingToolTest.Success("{\"success\":true}"), "vba." + action,
                JsonSerializer.Serialize(new { password }));
            Assert.False(call.Result.IsError);
            using var args = JsonDocument.Parse(call.Request.Args!);
            Assert.Equal(password, args.RootElement.GetProperty("password").GetString());
            Assert.DoesNotContain("\"password\":", call.JsonResult, StringComparison.Ordinal);
        }
    }

    [Theory]
    [InlineData("protect-project")]
    [InlineData("unlock-project")]
    [InlineData("verify-password")]
    public async Task PasswordActions_RejectNullAndEmptyBeforeDispatch(string action)
    {
        foreach (var password in new string?[] { null, "" })
        {
            var result = await fixture.CallResultWithoutDispatchAsync("vba", new()
            {
                ["action"] = action,
                ["workbook_session_id"] = "session-1",
                ["password"] = password
            });
            Assert.True(result.IsError);
            Assert.Contains("password", Assert.IsType<TextContentBlock>(Assert.Single(result.Content)).Text, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public async Task UnknownAuthentication_DoesNotBecomeFalseOrEchoThePassword()
    {
        const string password = " Synthetic\"Unknown\\!7 ";
        var call = await fixture.CallToolAsync("vba", new()
        {
            ["action"] = "unlock-project",
            ["workbook_session_id"] = "session-1",
            ["password"] = password
        }, new ServiceResponse
        {
            Success = false,
            ErrorCategory = "Timeout",
            ErrorMessage = "Authentication completion is unknown.",
            PartialState = """{"phase":"requesting-authentication","passwordVerified":null,"accessUnlocked":null,"dialogCleanupConfirmed":false,"saved":false,"reopened":false}"""
        }, "vba.unlock-project", JsonSerializer.Serialize(new { password }));
        Assert.True(call.Result.IsError);
        using var result = JsonDocument.Parse(call.JsonResult);
        var state = result.RootElement.GetProperty("partialState");
        Assert.Equal(JsonValueKind.Null, state.GetProperty("passwordVerified").ValueKind);
        Assert.Equal(JsonValueKind.Null, state.GetProperty("accessUnlocked").ValueKind);
        Assert.False(state.GetProperty("dialogCleanupConfirmed").GetBoolean());
        AssertSecretAbsent(password, call.JsonResult);
    }

    private static void AssertSecretAbsent(string password, string output)
    {
        Assert.DoesNotContain(password, output, StringComparison.Ordinal);
        Assert.DoesNotContain(JsonEncodedText.Encode(password).ToString(), output, StringComparison.Ordinal);
    }
}
