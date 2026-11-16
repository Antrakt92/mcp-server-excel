using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ExcelMcpServiceErrorTests
{
    [Fact]
    public async Task LateAppendDisconnectKeepsObservedExpansionWithUnknownOutcome()
    {
        using var service = new ExcelMcpService();
        var cause = CreateDiagnosticSnapshotFailure(unchecked((int)0x80010108));
        var failure = new TableAppendException("Append interrupted.", new TableAppendPartialState("Target", 3, 4, true, true), cause);
        var response = Assert.IsType<ServiceResponse>(await service.ResolveLateNativeResponseAsync(new ServiceRequest
        {
            Command = "table.append",
            SessionId = "late-append",
            LateNativeCompletion = Task.FromException<OperationResult>(failure)
        }));

        Assert.False(response.Success);
        Assert.Equal("ExcelDisconnected", response.ErrorCategory);
        using var json = JsonDocument.Parse(Assert.IsType<string>(response.PartialState));
        Assert.Equal("unknown", json.RootElement.GetProperty("outcome").GetString());
        Assert.True(json.RootElement.GetProperty("append").GetProperty("expansionVerified").GetBoolean());
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(false, true)]
    public async Task DisconnectReportsObservedPatchOrAppendProgress(bool excelAlive, bool append)
    {
        using var service = new ExcelMcpService();
        var cause = CreateDiagnosticSnapshotFailure(unchecked((int)0x80010108));
        var patch = new RangePatchResult { SheetName = "Sheet1", UnconfirmedAddress = "A3" };
        patch.WriteReturnedAddresses.Add("A1");
        patch.ConfirmedAddresses.Add("A1");
        Exception failure = append
            ? new TableAppendException("Append interrupted.", new TableAppendPartialState("Target", 3, 4, true, true), cause)
            : new RangePatchException(patch, cause);
        var batch = new FakeBatch
        {
            ExecuteException = new AggregateException(new InvalidOperationException("snapshot"), failure),
            IsAliveAfterExecuteException = excelAlive
        };
        const string sessionId = "agent-workflow-disconnect";
        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = append ? "table.append" : "range.patch-cells",
            SessionId = sessionId,
            Args = append
                ? """{"tableName":"Target","rows":[[1]]}"""
                : """{"sheetName":"Sheet1","patchOptions":{"edits":[{"address":"A1","expected":{"value":null},"replacement":{"value":1}}]}}"""
        });

        Assert.False(response.Success);
        Assert.Equal(excelAlive ? "ExcelDisconnected" : "ExcelProcessDied", response.ErrorCategory);
        using var json = JsonDocument.Parse(Assert.IsType<string>(response.PartialState));
        var state = json.RootElement;
        Assert.Equal("unknown", state.GetProperty("outcome").GetString());
        if (append)
        {
            var observed = state.GetProperty("append");
            Assert.Equal("Target", observed.GetProperty("tableName").GetString());
            Assert.True(observed.GetProperty("expansionVerified").GetBoolean());
        }
        else
        {
            Assert.Equal("A1", state.GetProperty("writeReturnedAddresses")[0].GetString());
            Assert.Equal("A1", state.GetProperty("confirmedAddresses")[0].GetString());
            Assert.Equal("A3", state.GetProperty("unconfirmedAddress").GetString());
        }
    }
}
