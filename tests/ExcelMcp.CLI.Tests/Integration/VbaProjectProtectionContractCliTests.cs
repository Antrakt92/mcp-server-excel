using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "VBA")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class VbaProjectProtectionContractCliTests
{
    private static readonly string[] BatchArguments = ["batch", "--input", "-", "--stop-on-error"];
    private const string AppliedJson = """{"success":true,"projectName":"SyntheticProject","currentProtection":"None","configurationApplied":true,"saveRequired":true,"reopenRequired":true}""";

    [Theory]
    [InlineData("  Synthetic leading and trailing !7  ")]
    [InlineData("Synthetic\"quote\\backslash!7")]
    [InlineData("Synthetic-Gr\u00fc\u00dfe-\u0416\u6f22\u5b57!7")]
    public async Task StdinBatch_TransportsExactPasswordAndPendingSaveResult(string password)
    {
        ServiceRequest? captured = null;
        var dispatches = 0;
        var result = await InProcessCliHelper.RunAsync(BatchArguments,
            request =>
            {
                captured = request;
                dispatches++;
                return new ServiceResponse { Success = true, Result = AppliedJson };
            }, Input(password));

        Assert.Equal(0, result.ExitCode);
        Assert.Equal(1, dispatches);
        Assert.NotNull(captured);
        Assert.Equal("vba.protect-project", captured.Command);
        Assert.Equal("session-1", captured.SessionId);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal(password, args.RootElement.GetProperty("password").GetString());
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.True(output.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("vba.protect-project", output.RootElement.GetProperty("command").GetString());
        var data = output.RootElement.GetProperty("data");
        Assert.Equal("SyntheticProject", data.GetProperty("projectName").GetString());
        Assert.Equal("None", data.GetProperty("currentProtection").GetString());
        Assert.True(data.GetProperty("configurationApplied").GetBoolean());
        Assert.True(data.GetProperty("saveRequired").GetBoolean());
        Assert.True(data.GetProperty("reopenRequired").GetBoolean());
        AssertSecretAbsent(password, result.Stdout);
        AssertSecretAbsent(password, result.Stderr);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task StdinBatch_MissingPasswordIsRefusedBeforeDispatch(bool explicitNull)
    {
        Dictionary<string, object?> args = new();
        if (explicitNull) args["password"] = null;
        var input = JsonSerializer.Serialize(new { command = "vba.protect-project", sessionId = "session-1", args });
        var dispatches = 0;
        var result = await InProcessCliHelper.RunAsync(BatchArguments,
            _ =>
            {
                dispatches++;
                return new ServiceResponse { Success = true, Result = AppliedJson };
            }, input);

        Assert.Equal(1, result.ExitCode);
        Assert.Equal(0, dispatches);
        using var output = JsonDocument.Parse(result.Stdout);
        var root = output.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        var error = root.GetProperty("error");
        Assert.Equal("InvalidInput", error.GetProperty("category").GetString());
        Assert.Contains("password", error.GetProperty("message").GetString(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task StdinBatch_UnconfirmedProtectionPreservesStateWithoutEchoingPassword()
    {
        const string password = "Synthetic\"Unconfirmed-Gr\u00fc\u00dfe!7";
        var dispatches = 0;
        var result = await InProcessCliHelper.RunAsync(BatchArguments,
            request =>
            {
                dispatches++;
                Assert.Equal("vba.protect-project", request.Command);
                using var args = JsonDocument.Parse(request.Args!);
                Assert.Equal(password, args.RootElement.GetProperty("password").GetString());
                return new ServiceResponse
                {
                    Success = false,
                    ErrorCategory = "Timeout",
                    ErrorMessage = "The protection outcome is unconfirmed.",
                    PartialState = """{"phase":"acceptance-requested","acceptanceRequested":true,"dialogCleanupConfirmed":false,"rollbackAttempted":false}"""
                };
            }, Input(password));

        Assert.Equal(1, result.ExitCode);
        Assert.Equal(1, dispatches);
        using var output = JsonDocument.Parse(result.Stdout);
        var root = output.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal("Timeout", root.GetProperty("error").GetProperty("category").GetString());
        Assert.Equal("The protection outcome is unconfirmed.", root.GetProperty("error").GetProperty("message").GetString());
        var state = root.GetProperty("partialState");
        Assert.Equal("acceptance-requested", state.GetProperty("phase").GetString());
        Assert.True(state.GetProperty("acceptanceRequested").GetBoolean());
        Assert.False(state.GetProperty("dialogCleanupConfirmed").GetBoolean());
        Assert.False(state.GetProperty("rollbackAttempted").GetBoolean());
        Assert.False(state.TryGetProperty("configurationApplied", out _));
        AssertSecretAbsent(password, result.Stdout);
        AssertSecretAbsent(password, result.Stderr);
    }

    private static string Input(string password) => JsonSerializer.Serialize(new
    {
        command = "vba.protect-project",
        sessionId = "session-1",
        args = new { password }
    });

    private static void AssertSecretAbsent(string password, string output)
    {
        Assert.DoesNotContain(password, output, StringComparison.Ordinal);
        Assert.DoesNotContain(JsonEncodedText.Encode(password).ToString(), output, StringComparison.Ordinal);
    }
}
