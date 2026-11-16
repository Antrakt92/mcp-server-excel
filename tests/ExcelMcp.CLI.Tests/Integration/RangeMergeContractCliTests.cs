using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class RangeMergeContractCliTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Merge_OnlyForwardsDiscardPermissionWhenSupplied(bool allow)
    {
        var arguments = new List<string>
        {
            "rangeformat", "merge-cells", "--session", "session-1", "--sheet-name", "Sheet1", "--range-address", "A1:B1"
        };
        if (allow)
            arguments.AddRange(["--overwrite-policy", "allow"]);
        ServiceRequest? captured = null;

        var result = await InProcessCliHelper.RunAsync(arguments, request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });

        Assert.Equal(0, result.ExitCode);
        Assert.NotNull(captured);
        Assert.Equal("rangeformat.merge-cells", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal(allow, args.RootElement.TryGetProperty("overwritePolicy", out var policy));
        if (allow)
            Assert.Equal("allow", policy.GetString(), ignoreCase: true);
    }
}
