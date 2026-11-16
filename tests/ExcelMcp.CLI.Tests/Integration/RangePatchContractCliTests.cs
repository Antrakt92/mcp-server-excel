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
public sealed class RangePatchContractCliTests
{
    [Fact]
    public async Task PatchRoutesFormulaSeparatelyFromExplicitBlankAndPreservesConflict()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "range", "patch-cells", "--session", "session-1", "--sheet-name", "Sheet1",
            "--patch-options", """{"edits":[{"address":"A1","expected":{"formula":"=1"},"replacement":{"value":null}}]}"""
        ], request =>
        {
            captured = request;
            return new ServiceResponse
            {
                Success = true,
                Result = """{"success":false,"errorMessage":"Cell contents differ at A1.","errorCategory":"Conflict","conflictAddresses":["A1"]}"""
            };
        });
        Assert.NotNull(captured);
        Assert.Equal("range.patch-cells", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        var edit = args.RootElement.GetProperty("patchOptions").GetProperty("edits")[0];
        Assert.Equal("=1", edit.GetProperty("expected").GetProperty("formula").GetString());
        Assert.False(edit.GetProperty("expected").TryGetProperty("value", out _));
        Assert.Equal(JsonValueKind.Null, edit.GetProperty("replacement").GetProperty("value").ValueKind);
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("A1", result.Stdout + result.Stderr, StringComparison.Ordinal);
    }
}
