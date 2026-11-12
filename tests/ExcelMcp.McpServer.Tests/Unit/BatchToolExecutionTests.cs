using System.Text.Json;
using Sbroenne.ExcelMcp.McpServer.ServiceBridge;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Sbroenne.ExcelMcp.McpServer.Telemetry;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Bridge = Sbroenne.ExcelMcp.McpServer.ServiceBridge.ServiceBridge;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "McpProtocol")]
[Trait("RequiresExcel", "false")]
public sealed class BatchToolExecutionTests
{
    [Fact]
    public async Task CommandBudget_RejectsOversizedBatchBeforeDispatch()
    {
        var backend = new Backend(_ => Task.FromResult(new ServiceResponse
        {
            Success = true,
            Result = """{"sheets":[]}"""
        }));
        using var bridge = new Bridge(() => backend);
        var commands = Enumerable.Repeat(new BatchCommandEntry { Command = "sheet.list" }, 129).ToArray();

        var result = await ExcelBatchTool.ExcelBatch(BatchAction.Run, bridge, commands, "target");

        Assert.Empty(backend.Requests);
        Assert.True(result.IsError);
        var output = result.StructuredContent!.Value;
        Assert.Equal("InvalidInput", output.GetProperty("errorCategory").GetString());
        Assert.Contains("128", output.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task ResultBudget_StopsBeforeNextCommandAndKeepsPriorResultsAndCrossingFailure()
    {
        const int itemSize = 4 * 1024 * 1024;
        var backend = new Backend(request => Task.FromResult(request.SessionId == "failure"
            ? new ServiceResponse
            {
                Success = false,
                ErrorCategory = "Permissions",
                ErrorMessage = "Original workbook error.",
                PartialState = JsonSerializer.Serialize(new { payload = new string('b', itemSize), marker = "retained" })
            }
            : new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { payload = new string('a', itemSize), marker = "first" })
            }));
        using var bridge = new Bridge(() => backend);

        var result = await ExcelBatchTool.ExcelBatch(BatchAction.Run, bridge,
            [
                new() { Command = "sheet.list", SessionId = "first" },
                new() { Command = "sheet.list", SessionId = "failure" },
                new() { Command = "sheet.list", SessionId = "must-not-dispatch" }
            ]);

        Assert.Equal(2, backend.Requests.Count);
        Assert.True(result.IsError);
        var output = result.StructuredContent!.Value;
        Assert.False(output.GetProperty("success").GetBoolean());
        Assert.Equal(3, output.GetProperty("totalCount").GetInt32());
        Assert.Equal(2, output.GetProperty("executedCount").GetInt32());
        Assert.True(output.GetProperty("stoppedEarly").GetBoolean());
        Assert.Equal("result-budget", output.GetProperty("stopReason").GetString());
        var items = output.GetProperty("results");
        Assert.Equal(2, items.GetArrayLength());
        Assert.Equal(0, items[0].GetProperty("index").GetInt32());
        Assert.Equal("first", items[0].GetProperty("marker").GetString());
        Assert.Equal(itemSize, items[0].GetProperty("payload").GetString()!.Length);
        Assert.Equal(1, items[1].GetProperty("index").GetInt32());
        Assert.Equal("Permissions", items[1].GetProperty("errorCategory").GetString());
        Assert.Equal("Original workbook error.", items[1].GetProperty("errorMessage").GetString());
        var state = items[1].GetProperty("partialState");
        Assert.Equal("retained", state.GetProperty("marker").GetString());
        Assert.Equal(itemSize, state.GetProperty("payload").GetString()!.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ResultBudget_ReportsFailureOnlyWhenAnotherCommandRemains(bool anotherCommand)
    {
        var backend = new Backend(_ => Task.FromResult(new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new { payload = new string('a', 8 * 1024 * 1024) })
        }));
        using var bridge = new Bridge(() => backend);
        BatchCommandEntry first = new() { Command = "sheet.list" };
        BatchCommandEntry[] commands = anotherCommand ? [first, first] : [first];

        var result = await ExcelBatchTool.ExcelBatch(BatchAction.Run, bridge, commands, "target");

        Assert.Single(backend.Requests);
        var output = result.StructuredContent!.Value;
        Assert.Equal(anotherCommand, result.IsError == true);
        Assert.Equal(!anotherCommand, output.GetProperty("success").GetBoolean());
        Assert.Equal(1, output.GetProperty("executedCount").GetInt32());
        Assert.Equal(anotherCommand, output.GetProperty("stoppedEarly").GetBoolean());
        Assert.Equal(anotherCommand, output.TryGetProperty("stopReason", out var reason));
        if (anotherCommand)
        {
            Assert.Equal("result-budget", reason.GetString());
            Assert.Equal("ResourceLimit", output.GetProperty("errorCategory").GetString());
        }
    }

    [Fact]
    public async Task CancelledOpen_DoesNotRevealTheDefaultWorkbook()
    {
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var response = new TaskCompletionSource<ServiceResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
        var revealed = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var backend = new Backend(_ =>
        {
            started.SetResult();
            return response.Task;
        }, () => revealed.TrySetResult());
        using var bridge = new Bridge(() => backend);
        using var cancellation = new CancellationTokenSource();
        var pending = ExcelBatchTool.ExcelBatch(BatchAction.Run, bridge,
            [new() { Command = "session.open", Args = JsonSerializer.SerializeToElement(new { filePath = @"C:\fixture.xlsx" }) }],
            "existing-hidden-session", cancellationToken: cancellation.Token);
        try
        {
            await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => pending);
            Assert.Null(Assert.Single(backend.Requests).SessionId);
            Assert.Empty(backend.Revealed);
        }
        finally
        {
            response.SetResult(new ServiceResponse { Success = true, Result = """{"success":true,"sessionId":"new-session"}""" });
            await revealed.Task.WaitAsync(TimeSpan.FromSeconds(5));
        }
        Assert.Equal(["new-session"], backend.Revealed);
    }

    [Fact]
    public async Task Cancellation_StopsWaitingForValidationWithoutStartingExcel()
    {
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var validationCompleted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        using var cancellation = new CancellationTokenSource();
        var backend = new Backend(_ => throw new InvalidOperationException("Validation must finish before dispatch."));
        using var bridge = new Bridge(() => backend);
        var pending = Task.Run(() => ExcelBatchTool.ExecuteBatchAsync(BatchAction.Run, bridge,
            [new() { Command = "range.get-values" }], "existing-session", false, cancellation.Token,
            (_, _) =>
            {
                entered.Set();
                release.Wait();
                validationCompleted.SetResult();
            },
            (_, _, _, _) => { }));
        try
        {
            Assert.True(entered.Wait(TimeSpan.FromSeconds(5)));
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => pending.WaitAsync(TimeSpan.FromSeconds(2)));
            Assert.False(validationCompleted.Task.IsCompleted);
            Assert.Empty(backend.Requests);
        }
        finally
        {
            release.Set();
            await validationCompleted.Task.WaitAsync(TimeSpan.FromSeconds(5));
            try { await pending; }
            catch (OperationCanceledException) when (cancellation.IsCancellationRequested) { }
        }
    }

    [Fact]
    public async Task Usage_RecordsEachEntryWithSafeIdentityAndNoExtraBatchOperation()
    {
        var backend = new Backend(_ => Task.FromResult(new ServiceResponse
        {
            Success = true,
            Result = """{"success":true,"sheets":[]}"""
        }));
        using var bridge = new Bridge(() => backend);
        var recorded = new List<(string Tool, string Action, ToolInvocationResult Result)>();
        await ExcelBatchTool.ExecuteBatchAsync(BatchAction.Run, bridge,
            [
                new() { Command = "sheet.list" },
                new() { Command = "range.get-values", Args = JsonSerializer.SerializeToElement(new { sheet = "private-sheet", rangeAddress = "A1" }) },
                new() { Command = @"private-workbook.xlsx.secret" }
            ], "private-session", false, CancellationToken.None,
            Sbroenne.ExcelMcp.Generated.ServiceRegistry.ValidateCommandArguments,
            (tool, action, _, result) => recorded.Add((tool, action, result)));

        Assert.Equal(3, recorded.Count);
        Assert.Equal(("worksheet_read", "list", new ToolInvocationResult(ToolInvocationOutcome.Succeeded, null)), recorded[0]);
        Assert.Equal(("range_read", "get-values", new ToolInvocationResult(ToolInvocationOutcome.Failed, ToolFailureClass.InputState)), recorded[1]);
        Assert.Equal(("other", "other", new ToolInvocationResult(ToolInvocationOutcome.Failed, ToolFailureClass.InputState)), recorded[2]);
        Assert.Single(backend.Requests);
    }

    [Fact]
    public async Task Usage_FileTestNegativeWithCanOpen_RecordsExpectedNegative()
    {
        var backend = new Backend(_ => Task.FromResult(new ServiceResponse
        {
            Success = true,
            Result = """{"canOpen":false,"exists":false,"message":"Missing fixture"}"""
        }));
        using var bridge = new Bridge(() => backend);
        ToolInvocationResult? recorded = null;
        await ExcelBatchTool.ExecuteBatchAsync(BatchAction.Run, bridge,
            [new() { Command = "session.test", Args = JsonSerializer.SerializeToElement(new { filePath = @"C:\fixture.xlsx" }) }],
            null, false, CancellationToken.None,
            Sbroenne.ExcelMcp.Generated.ServiceRegistry.ValidateCommandArguments,
            (_, _, _, result) => recorded = result);
        Assert.Equal(new ToolInvocationResult(ToolInvocationOutcome.ExpectedNegative, null), recorded);
    }

    [Theory]
    [InlineData("Timeout", "TimeoutCancellation", "Timeout")]
    [InlineData("InvalidInput", "InputState", null)]
    [InlineData("ServiceUnavailable", "InternalProductFault", null)]
    public async Task Usage_FileTestFailureRetainsItsClassAndCause(
        string category, string failureClass, string? cause)
    {
        var backend = new Backend(_ => Task.FromResult(new ServiceResponse
        {
            Success = false,
            ErrorCategory = category,
            ErrorMessage = "File test failed"
        }));
        using var bridge = new Bridge(() => backend);
        ToolInvocationResult? recorded = null;
        await ExcelBatchTool.ExecuteBatchAsync(BatchAction.Run, bridge,
            [new() { Command = "session.test", Args = JsonSerializer.SerializeToElement(new { filePath = @"C:\fixture.xlsx" }) }],
            null, false, CancellationToken.None,
            Sbroenne.ExcelMcp.Generated.ServiceRegistry.ValidateCommandArguments,
            (_, _, _, result) => recorded = result);
        Assert.Equal(new ToolInvocationResult(ToolInvocationOutcome.Failed,
            Enum.Parse<ToolFailureClass>(failureClass), cause is null ? null : Enum.Parse<ToolFailureCause>(cause)), recorded);
    }

    [Fact]
    public async Task Cancellation_StopsDispatchAndRetainsPendingWork()
    {
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var response = new TaskCompletionSource<ServiceResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
        var backend = new Backend(_ =>
        {
            started.SetResult();
            return response.Task;
        });
        using var bridge = new Bridge(() => backend);
        using var cancellation = new CancellationTokenSource();
        var pending = ExcelBatchTool.ExcelBatch(BatchAction.Run, bridge,
            [new() { Command = "sheet.list" }, new() { Command = "sheet.list" }],
            "session-1", cancellationToken: cancellation.Token);
        await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => pending);
        Assert.Single(backend.Requests);
        Assert.True(backend.Requests[0].CallerCancellation.IsCancellationRequested);
        Assert.Equal(["session-1"], backend.Revealed);
        Assert.False(backend.Disposed);
        Assert.False(response.Task.IsCompleted);
        response.SetResult(new ServiceResponse { Success = true, Result = """{"success":true,"sheets":[]}""" });
    }

    [Fact]
    public async Task UnexpectedExecutionFailure_PropagatesInsteadOfBecomingInvalidInput()
    {
        var backend = new Backend(_ => throw new JsonException("Invalid Service response"));
        using var bridge = new Bridge(() => backend);
        var exception = await Assert.ThrowsAsync<JsonException>(() => ExcelBatchTool.ExcelBatch(
            BatchAction.Run, bridge, [new() { Command = "sheet.list" }, new() { Command = "sheet.list" }],
            "session-1"));
        Assert.Equal("Invalid Service response", exception.Message);
        Assert.Single(backend.Requests);
        Assert.False(backend.Disposed);
    }

    [Fact]
    public async Task AlreadyCancelled_DoesNotStartBackend()
    {
        var started = false;
        using var bridge = new Bridge(() =>
        {
            started = true;
            throw new InvalidOperationException("Should not start");
        });
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ExcelBatchTool.ExcelBatch(
            BatchAction.Run, bridge, [new() { Command = "sheet.list" }],
            cancellationToken: cancellation.Token));
        Assert.False(started);
    }

    private sealed class Backend(Func<ServiceRequest, Task<ServiceResponse>> process, Action? onReveal = null) : IServiceBridgeBackend
    {
        public List<ServiceRequest> Requests { get; } = [];
        public List<string> Revealed { get; } = [];
        public bool Disposed { get; private set; }

        public Task<ServiceResponse> ProcessAsync(ServiceRequest request)
        {
            Requests.Add(request);
            return process(request);
        }

        public void RevealSession(string sessionId)
        {
            Revealed.Add(sessionId);
            onReveal?.Invoke();
        }
        public void Dispose() => Disposed = true;
    }
}
