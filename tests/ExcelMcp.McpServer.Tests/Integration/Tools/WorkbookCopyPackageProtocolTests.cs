using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Parameters")]
[Trait("RequiresExcel", "false")]
public sealed class WorkbookCopyPackageProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task SaveCopyAs_ForwardsPackageCheckAndRetainsRefusalDetails()
    {
        var response = new ServiceResponse
        {
            Success = true,
            Result = """
                {"success":false,"errorCategory":"Conflict","errorMessage":"Package change refused.","filePublished":false,"packageComparison":{"matchesAllowlist":false,"changes":[{"partName":"xl/styles.xml"}]}}
                """
        };
        using var policy = JsonDocument.Parse("""{"baselinePath":"C:/source.xlsm","allowedParts":["xl/vbaProject.bin"]}""");
        var (call, requests) = await fixture.CallSequenceAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "save-copy-as",
            ["workbook_session_id"] = "recording-session",
            ["target_path"] = "C:/result.xlsm",
            ["package_check"] = policy.RootElement
        }, response);

        var request = Assert.Single(requests);
        Assert.Equal("workbook.save-copy-as", request.Command);
        Assert.Equal("recording-session", request.SessionId);
        using var args = JsonDocument.Parse(request.Args!);
        Assert.Equal("C:/result.xlsm", args.RootElement.GetProperty("targetPath").GetString());
        var check = args.RootElement.GetProperty("packageCheck");
        Assert.Equal("C:/source.xlsm", check.GetProperty("baselinePath").GetString());
        Assert.Equal("xl/vbaProject.bin", Assert.Single(check.GetProperty("allowedParts").EnumerateArray()).GetString());
        using var result = JsonDocument.Parse(McpResponseAssertions.ReadText(call, expectedError: true));
        Assert.False(result.RootElement.GetProperty("filePublished").GetBoolean());
        Assert.False(result.RootElement.GetProperty("packageComparison").GetProperty("matchesAllowlist").GetBoolean());
    }
}
