using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "ServiceDaemon")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class CompletionRecordingFailureTests
{
    [Theory]
    [InlineData(false, "Timeout", null, null, true)]
    [InlineData(false, "Cancelled", null, null, true)]
    [InlineData(false, "ExcelDisconnected", null, "{\"outcome\":\"completed\"}", true)]
    [InlineData(false, "ServiceUnavailable", null, "{\"outcome\":\"unknown\"}", true)]
    [InlineData(true, null, "{\"success\":false,\"errorCategory\":\"Timeout\"}", null, true)]
    [InlineData(false, "ComInterop", null, "{\"outcome\":\"completed\"}", false)]
    [InlineData(true, null, "{\"success\":true}", null, false)]
    [InlineData(true, null, "{", null, true)]
    [InlineData(true, null, "{\"success\":true,\"success\":false}", null, true)]
    public void CompletionRecordingFailure_PreservesTheBatchStopDecision(
        bool delivered, string? category, string? result, string? partialState, bool mustStop)
    {
        var request = new ServiceRequest { Command = "vba.run", SessionId = "target" };
        var original = new ServiceResponse
        {
            Success = delivered,
            Command = request.Command,
            SessionId = request.SessionId,
            ErrorCategory = category,
            Result = result,
            PartialState = partialState
        };

        var wrapped = original.WithCompletionFailure(request, new IOException("Could not record workbook ownership."));
        var batchOutcome = ResultJson.Read(wrapped.Result, request.Command, wrapped.Success,
            wrapped.ErrorMessage, wrapped.ErrorCategory, wrapped.SessionId, wrapped.ExceptionType,
            wrapped.HResult, wrapped.InnerError, wrapped.PartialState);

        Assert.False(batchOutcome.Success);
        Assert.Equal(mustStop, batchOutcome.MayStillBeRunning);
        Assert.Equal("ServiceUnavailable", batchOutcome.Error!["category"]!.GetValue<string>());
        using var state = JsonDocument.Parse(wrapped.PartialState!);
        Assert.Equal(mustStop ? "unknown" : "completed", state.RootElement.GetProperty("outcome").GetString());
        Assert.True(JsonElement.DeepEquals(JsonSerializer.SerializeToElement(original, ServiceProtocol.JsonOptions),
            state.RootElement.GetProperty("response")));
    }
}
