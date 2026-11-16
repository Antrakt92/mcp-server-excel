using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Parameters")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class NamedRangeContractCliTests
{
    [Theory]
    [InlineData("namedrange create --name Marker --reference Sheet1!$A$1 --visible false",
        "namedrange.create", "visible", false)]
    [InlineData("namedrange read --name Marker --metadata-only true",
        "namedrange.read", "metadataOnly", true)]
    public async Task NamedRangeOptions_ForwardExplicitValues(
        string command, string expectedCommand, string option, bool expectedValue)
    {
        var (result, requests) = await InProcessCliHelper.RunRecordingAsync(command + " --session target",
            new ServiceResponse { Success = true, Result = """{"success":true}""" });

        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        var request = Assert.Single(requests);
        Assert.Equal(expectedCommand, request.Command);
        Assert.Equal("target", request.SessionId);
        using var args = JsonDocument.Parse(request.Args!);
        Assert.Equal("Marker", args.RootElement.GetProperty("name").GetString());
        Assert.Equal(expectedValue, args.RootElement.GetProperty(option).GetBoolean());
    }
}
