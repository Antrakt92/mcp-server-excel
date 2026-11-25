using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "ChartDepth")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ChartOptionValidationCliTests
{
    [Theory]
    [InlineData("set-error-bars", "--error-bar-options", """{"enable":false,"kind":"StandardError"}""", "enable")]
    [InlineData("set-point-format", "--point-options", """{"fillColour":"#112233"}""", "fillColour")]
    public async Task TypedChartPayloads_RejectUnknownOptionsBeforeDispatch(
        string action, string option, string json, string unknownOption)
    {
        var calls = 0;
        var args = new List<string>
        {
            "-q", "chartconfig", action, "--session", "session-1", "--chart-name", "Sales", "--series-index", "2"
        };
        if (action == "set-point-format")
            args.AddRange(["--point-index", "3"]);
        args.AddRange([option, json]);
        var result = await InProcessCliHelper.RunAsync([.. args], _ =>
        {
            calls++;
            return new ServiceResponse { Success = false, ErrorMessage = "The test service must not receive this command." };
        });

        Assert.Equal(1, result.ExitCode);
        Assert.Equal(0, calls);
        using var output = JsonDocument.Parse(result.Stdout);
        var error = output.RootElement.GetProperty("error");
        Assert.Equal("InvalidInput", error.GetProperty("category").GetString());
        Assert.Contains($"'{unknownOption}'", error.GetProperty("innerError").GetString(), StringComparison.Ordinal);
    }
}
