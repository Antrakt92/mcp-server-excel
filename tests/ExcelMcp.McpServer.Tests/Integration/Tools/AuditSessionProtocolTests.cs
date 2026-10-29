using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "File")]
[Trait("RequiresExcel", "false")]
public sealed class AuditSessionProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Open_MapsAuditModeAndReportsItsSession(bool readOnly)
    {
        var call = await fixture.CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "open",
            ["file_path"] = "C:/fixture.xlsx",
            ["read_only"] = readOnly
        }, new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new { success = true, sessionId = "audit-fixture", auditOnly = readOnly })
        }, "session.open", expectedSessionId: null,
            expectedArgsJson: JsonSerializer.Serialize(new { filePath = "C:/fixture.xlsx", show = false, timeoutSeconds = 120, readOnly = readOnly ? (bool?)true : null }, ServiceProtocol.JsonOptions));
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.Equal("audit-fixture", result.RootElement.GetProperty("workbook_session_id").GetString());
        Assert.Equal(readOnly, result.RootElement.GetProperty("auditOnly").GetBoolean());
    }
}
