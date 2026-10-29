using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "CommandRuntime")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
[Collection("Sequential")]
public sealed class AuditSessionAdapterTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Open_ForwardsExplicitAuditModeWithoutChangingVisibility(bool readOnly)
    {
        var arguments = new List<string> { "session", "open", "C:/fixture.xlsx" };
        if (readOnly) arguments.Add("--read-only");
        var calls = 0;
        var result = await InProcessCliHelper.RunAsync(arguments, request =>
        {
            calls++;
            Assert.Equal("session.open", request.Command);
            using var args = JsonDocument.Parse(request.Args!);
            Assert.Equal(readOnly, args.RootElement.TryGetProperty("readOnly", out var mode) && mode.GetBoolean());
            Assert.False(args.RootElement.GetProperty("show").GetBoolean());
            Assert.Equal("C:/fixture.xlsx", args.RootElement.GetProperty("filePath").GetString());
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { success = true, sessionId = "audit-fixture", auditOnly = readOnly })
            };
        });
        Assert.Equal(1, calls);
        Assert.Equal(0, result.ExitCode);
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.Equal(readOnly, output.RootElement.GetProperty("data").GetProperty("auditOnly").GetBoolean());
    }
}
