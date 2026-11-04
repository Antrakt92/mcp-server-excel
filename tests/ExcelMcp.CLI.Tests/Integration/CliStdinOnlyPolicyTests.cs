using System.Reflection;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "VBA")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class CliStdinOnlyPolicyTests
{
    private const string Secret = "  Synthetic\"quote\\slash !7  ";

    [Theory]
    [InlineData("protect-project")]
    [InlineData("unlock-project")]
    [InlineData("verify-password")]
    public async Task DirectInvocation_IsRefusedWithoutPrintingArgumentsOrConnecting(string action)
    {
        var result = await Run(["VBA", "--session", "owned", action.ToUpperInvariant(),
            "--password", Secret, "--timeout", Secret]);
        Assert.Equal(1, result.ExitCode);
        Assert.Equal(0, result.Client.Connections);
        Assert.Empty(result.Client.Requests);
        Assert.Contains("batch --input -", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain(Secret, result.Output + result.Error, StringComparison.Ordinal);
        Assert.DoesNotContain(JsonEncodedText.Encode(Secret).ToString(), result.Output + result.Error, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("--session:owned")]
    [InlineData("-s:owned")]
    [InlineData("--session=owned")]
    [InlineData("-s=owned")]
    public async Task InlineSessionOptions_CannotBypassRefusalBeforeTypedConversion(string sessionOption)
    {
        foreach (var action in new[] { "protect-project", "unlock-project", "verify-password" })
        {
            var result = await Run(["vba", sessionOption, action, "--start-line", Secret]);
            Assert.Equal(1, result.ExitCode);
            Assert.Equal(0, result.Client.Connections);
            Assert.Empty(result.Client.Requests);
            using var output = JsonDocument.Parse(result.Output);
            Assert.Equal(_CliCategoryMetadata.BatchStdinOnlyMessage,
                output.RootElement.GetProperty("error").GetProperty("message").GetString());
            Assert.DoesNotContain(Secret, result.Output + result.Error, StringComparison.Ordinal);
            Assert.DoesNotContain(JsonEncodedText.Encode(Secret).ToString(), result.Output + result.Error, StringComparison.Ordinal);
        }
    }

    [Theory]
    [InlineData("protect-project")]
    [InlineData("unlock-project")]
    [InlineData("verify-password")]
    public async Task SensitiveActionHelp_AdvertisesStdinWithoutPasswordFlags(string action)
    {
        var result = await Run(["vba", action, "--help"]);
        Assert.Equal(0, result.ExitCode);
        Assert.Equal(0, result.Client.Connections);
        Assert.Contains("batch --input -", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("--password", result.Output, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null, Secret)]
    [InlineData("-", Secret)]
    [InlineData(null, " ")]
    [InlineData("-", " ")]
    [InlineData(null, "   ")]
    [InlineData("-", "   ")]
    public async Task StdinBatch_TransportsExactCredentials(string? inputFile, string password)
    {
        foreach (var action in new[] { "protect-project", "unlock-project", "verify-password" })
        {
            var arguments = inputFile is null ? new[] { "batch" } : new[] { "batch", "--input", inputFile };
            var result = await Run(arguments, Entry("vba." + action, password));
            Assert.Equal(0, result.ExitCode);
            Assert.Equal(1, result.Client.Connections);
            var request = Assert.Single(result.Client.Requests);
            using var payload = JsonDocument.Parse(request.Args!);
            Assert.Equal(password, payload.RootElement.GetProperty("password").GetString());
            Assert.DoesNotContain("\"password\":", result.Output + result.Error, StringComparison.Ordinal);
            Assert.DoesNotContain(Secret, result.Output + result.Error, StringComparison.Ordinal);
            Assert.DoesNotContain(JsonEncodedText.Encode(Secret).ToString(), result.Output + result.Error, StringComparison.Ordinal);
        }
    }

    [Theory]
    [InlineData("{}")]
    [InlineData("{\"password\":null}")]
    [InlineData("{\"password\":\"\"}")]
    [InlineData("{\"password\":42}")]
    [InlineData("{\"password\":true}")]
    public async Task StdinBatch_InvalidRequiredPasswordIsRejectedBeforeConnecting(string argsJson)
    {
        foreach (var action in new[] { "protect-project", "unlock-project", "verify-password" })
        {
            var input = "{\"command\":\"vba." + action + "\",\"sessionId\":\"owned\",\"args\":" + argsJson + "}";
            var result = await Run(["batch"], input);
            Assert.Equal(1, result.ExitCode);
            Assert.Equal(0, result.Client.Connections);
            Assert.Empty(result.Client.Requests);
            using var output = JsonDocument.Parse(result.Output);
            Assert.Equal("InvalidInput", output.RootElement.GetProperty("error").GetProperty("category").GetString());
        }
    }

    [Theory]
    [InlineData("vba.protect-project")]
    [InlineData("VBA.UNLOCK-PROJECT")]
    [InlineData("vba.verify-password")]
    public async Task MixedFileBatch_IsRejectedBeforeAnyConnectionOrDispatch(string command)
    {
        string path = Path.Combine(Path.GetTempPath(), $"stdin-policy-{Guid.NewGuid():N}.json");
        try
        {
            await File.WriteAllTextAsync(path, "[{\"command\":\"diag.ping\"}," + Entry(command) + "]");
            var result = await Run(["batch", "--input", path]);
            Assert.Equal(1, result.ExitCode);
            Assert.Equal(0, result.Client.Connections);
            Assert.Empty(result.Client.Requests);
            Assert.Contains("batch --input -", result.Output, StringComparison.Ordinal);
            Assert.DoesNotContain(Secret, result.Output + result.Error, StringComparison.Ordinal);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public async Task NonSensitiveFileBatch_RemainsAvailable()
    {
        string path = Path.Combine(Path.GetTempPath(), $"stdin-policy-{Guid.NewGuid():N}.json");
        try
        {
            await File.WriteAllTextAsync(path, "[{\"command\":\"diag.ping\"}]");
            var result = await Run(["batch", "--input", path]);
            Assert.Equal(0, result.ExitCode);
            Assert.Equal("diag.ping", Assert.Single(result.Client.Requests).Command);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public async Task MalformedBatch_DoesNotEchoInputTokens()
    {
        var result = await Run(["batch"], "{\"command\":\"vba.protect-project\",\"" + Secret + "\":0}");
        Assert.Equal(1, result.ExitCode);
        Assert.Equal(0, result.Client.Connections);
        Assert.DoesNotContain("Synthetic", result.Output + result.Error, StringComparison.Ordinal);
        using var output = JsonDocument.Parse(result.Output);
        var failure = output.RootElement.GetProperty("error");
        Assert.Equal(nameof(JsonException), failure.GetProperty("exceptionType").GetString());
        Assert.Contains("LineNumber:", failure.GetProperty("message").GetString(), StringComparison.Ordinal);
        Assert.Contains("BytePositionInLine:", failure.GetProperty("message").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public void GeneratedPolicyAndSettings_MatchSourceAttributesWithoutPasswordOption()
    {
        var methods = typeof(IVbaCommands).GetMethods().Select(method => method.GetCustomAttribute<ServiceActionAttribute>())
            .Where(attribute => attribute?.CliBatchStdinOnly == true).ToArray();
        Assert.Equal(3, methods.Length);
        foreach (var method in methods)
        {
            Assert.True(_CliCategoryMetadata.RequiresBatchStdin("VBA." + method!.Action.ToUpperInvariant()));
            var error = Assert.Throws<ArgumentException>(() => ServiceRegistry.Vba.RouteFromSettings(method.Action,
                new ServiceRegistry.Vba.CliSettings { SessionId = "owned", Action = method.Action }));
            Assert.Equal(_CliCategoryMetadata.BatchStdinOnlyMessage, error.Message);
        }
        Assert.Null(typeof(ServiceRegistry.Vba.CliSettings).GetProperty("Password"));
        Assert.False(_CliCategoryMetadata.RequiresBatchStdin("vba.status"));
    }

    private static string Entry(string command, string password = Secret) => JsonSerializer.Serialize(new
    {
        command,
        sessionId = "owned",
        args = new { password }
    });

    private static async Task<(int ExitCode, string Output, string Error, RecordingClient Client)> Run(
        string[] arguments, string input = "")
    {
        using var output = new StringWriter();
        using var error = new StringWriter();
        using var reader = new StringReader(input);
        var client = new RecordingClient();
        var runtime = new CliCommandRuntime(client, reader, output, error, isOutputRedirected: true);
        int code = await Program.RunAsync(["--quiet", .. arguments], runtime);
        return (code, output.ToString(), error.ToString(), client);
    }

    private sealed class RecordingClient : ICliRequestClientFactory, ICliRequestClient
    {
        internal int Connections { get; private set; }
        internal List<ServiceRequest> Requests { get; } = [];
        public Task<ICliRequestClient> ConnectAsync(CancellationToken cancellationToken)
        {
            Connections++;
            return Task.FromResult<ICliRequestClient>(this);
        }
        public Task<ServiceResponse> SendAsync(ServiceRequest request, CancellationToken cancellationToken)
        {
            Requests.Add(request);
            return Task.FromResult(new ServiceResponse { Success = true, Result = "{\"success\":true}" });
        }
        public void Dispose() { }
    }
}
