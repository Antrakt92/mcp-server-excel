using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "ServiceDaemon")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class OperationReceiptsTests
{
    [Fact]
    public async Task OversizedSessionIdIsRejectedBeforeTracking()
    {
        using var controller = new ExcelController(launcher: new("unused.exe", []));
        var operation = Guid.NewGuid().ToString();
        var result = await controller.ProcessAsync(new ServiceRequest
        {
            Command = "session.save",
            SessionId = new string('s', 129),
            OperationId = operation
        });
        Assert.Equal("InvalidInput", result.ErrorCategory);
        using var lookup = JsonDocument.Parse((await controller.ProcessAsync(new ServiceRequest
        {
            Command = OperationReceipts.LookupCommand,
            SessionId = "target",
            Args = JsonSerializer.Serialize(new { operationId = operation })
        })).Result!);
        Assert.Equal("unknown", lookup.RootElement.GetProperty("state").GetString());
    }

    [Fact]
    public void RetentionDoesNotEvictPendingWorkAndKeepsExpiredResultsUnknown()
    {
        var clock = new ReceiptClock();
        var receipts = new OperationReceipts(clock, capacity: 1);
        var id = Guid.NewGuid();
        const string session = "target";
        Assert.Equal(OperationReceipts.Admission.Accepted,
            receipts.TryBegin(id, "table.append", session, out var operation));

        clock.Now += TimeSpan.FromHours(1);
        Assert.Equal(OperationReceipts.Admission.Full,
            receipts.TryBegin(Guid.NewGuid(), "session.save", session, out _));
        Assert.Equal("pending", Lookup(receipts, id, session).GetProperty("state").GetString());
        Assert.Equal("unknown", Lookup(receipts, id, "another-session").GetProperty("state").GetString());
        Assert.Equal(OperationReceipts.Admission.Duplicate,
            receipts.TryBegin(id, "table.append", session, out _));

        receipts.Complete(operation, new ServiceResponse
        {
            Success = false,
            ErrorCategory = "ComInterop",
            ErrorMessage = "Append failed after expansion.",
            PartialState = """{"rowsAdded":1}"""
        });
        var result = Lookup(receipts, id, session);
        Assert.Equal("failed", result.GetProperty("state").GetString());
        Assert.Equal("ComInterop", result.GetProperty("result").GetProperty("errorCategory").GetString());
        Assert.Equal(1, result.GetProperty("result").GetProperty("partialState").GetProperty("rowsAdded").GetInt32());
        clock.Now += TimeSpan.FromMinutes(15);
        Assert.Equal("unknown", Lookup(receipts, id, session).GetProperty("state").GetString());
        Assert.Equal(OperationReceipts.Admission.Accepted,
            receipts.TryBegin(Guid.NewGuid(), "session.save", session, out _));
    }

    [Theory]
    [InlineData("ServiceUnavailable", "{\"outcome\":\"unknown\"}")]
    [InlineData("ExcelProcessDied", null)]
    public void TransportDisconnectKeepsTheOperationOutcomeUnknown(string category, string? partialState)
    {
        var receipts = new OperationReceipts();
        var id = Guid.NewGuid();
        Assert.Equal(OperationReceipts.Admission.Accepted, receipts.TryBegin(id, "session.save", "target", out var operation));
        receipts.Complete(operation, new ServiceResponse
        {
            Success = false,
            ErrorCategory = category,
            ErrorMessage = "Worker exited before replying.",
            PartialState = partialState
        });
        var result = Lookup(receipts, id, "target");
        Assert.Equal("unknown", result.GetProperty("state").GetString());
        Assert.False(result.GetProperty("result").GetProperty("success").GetBoolean());
    }

    [Fact]
    public void RecordingFailureDoesNotTurnConfirmedSaveIntoFailedWorkbookOperation()
    {
        const string session = "target";
        var receipts = new OperationReceipts();
        var id = Guid.NewGuid();
        var request = new ServiceRequest { Command = "session.save", SessionId = session };
        Assert.Equal(OperationReceipts.Admission.Accepted, receipts.TryBegin(id, request.Command, session, out var operation));
        var completed = new ServiceResponse { Success = true, Result = """{"success":true,"saved":true}""" };
        receipts.Complete(operation, completed.WithCompletionFailure(request, new IOException("Ownership record unavailable.")));

        var result = Lookup(receipts, id, session);
        Assert.Equal("completed", result.GetProperty("state").GetString());
        Assert.True(result.GetProperty("result").GetProperty("saved").GetBoolean());
        Assert.Contains("Ownership record unavailable.", result.GetProperty("recordingError").GetString(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("session.save", "{\"success\":true}", "unknown")]
    [InlineData("session.save", "{\"saved\":false}", "unknown")]
    [InlineData("table.append", "{\"success\":true}", "completed")]
    public void SuccessfulTransportAloneDoesNotProveWorkbookChanges(string command, string payload, string state)
    {
        var receipts = new OperationReceipts();
        var id = Guid.NewGuid();
        Assert.Equal(OperationReceipts.Admission.Accepted, receipts.TryBegin(id, command, "target", out var operation));
        receipts.Complete(operation, new ServiceResponse { Success = true, Result = payload });
        Assert.Equal(state, Lookup(receipts, id, "target").GetProperty("state").GetString());
    }

    [Theory]
    [InlineData(false, "unknown")]
    [InlineData(true, "partially-applied")]
    public async Task LateAppendFailureRetainsVerifiedExpansionAndWritePhase(bool expansionVerified, string expectedState)
    {
        const string session = "target";
        var receipts = new OperationReceipts();
        var id = Guid.NewGuid();
        using var service = new ExcelMcpService();
        var request = new ServiceRequest
        {
            Command = "table.append",
            SessionId = session,
            LateNativeCompletion = Task.FromException<OperationResult>(new TableAppendException("Append stopped.",
                new("Stock", 5, 7, expansionVerified, false), new InvalidOperationException("Native failure")))
        };
        Assert.Equal(OperationReceipts.Admission.Accepted, receipts.TryBegin(id, request.Command, session, out var operation));
        var response = await service.ResolveLateNativeResponseAsync(request);
        receipts.Complete(operation, Assert.IsType<ServiceResponse>(response));
        var result = Lookup(receipts, id, session);
        Assert.Equal(expectedState, result.GetProperty("state").GetString());
        var append = result.GetProperty("result").GetProperty("partialState").GetProperty("append");
        Assert.Equal("Stock", append.GetProperty("tableName").GetString());
        Assert.Equal(5, append.GetProperty("originalRowCount").GetInt32());
        Assert.Equal(7, append.GetProperty("expectedRowCount").GetInt32());
        Assert.Equal(expansionVerified, append.GetProperty("expansionVerified").GetBoolean());
        Assert.False(append.GetProperty("valueWriteAttempted").GetBoolean());
    }

    private static JsonElement Lookup(OperationReceipts receipts, Guid id, string sessionId)
    {
        var response = receipts.Lookup(JsonSerializer.Serialize(new { operationId = id }), sessionId);
        using var json = JsonDocument.Parse(response.Result!);
        return json.RootElement.Clone();
    }

    private sealed class ReceiptClock : TimeProvider
    {
        internal DateTimeOffset Now { get; set; } = DateTimeOffset.UnixEpoch;
        public override DateTimeOffset GetUtcNow() => Now;
    }
}
