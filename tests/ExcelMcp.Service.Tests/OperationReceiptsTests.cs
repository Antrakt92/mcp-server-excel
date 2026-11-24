using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Service.Workers;
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
    public void RecordedWorkerPhaseAndLateExitAreBoundToTheOriginalOperationAndProcess()
    {
        var clock = new ReceiptClock();
        var receipts = new OperationReceipts(clock);
        var id = Guid.NewGuid();
        var session = Guid.NewGuid().ToString("N");
        const string command = "vba.invoke";
        Assert.Equal(OperationReceipts.Admission.Accepted, receipts.TryBegin(id, command, session, out var operation));
        receipts.RecordRequestedProcedure(operation, """{"procedureName":"TimeLogs.CommitThenRefresh","arguments":[{"value":"private fixture value"}]}""");
        Assert.Equal("controllerAccepted", Lookup(receipts, id, session).GetProperty("operationEvidence").GetProperty("phase").GetString());

        var self = WorkerProcess.Current;
        var excel = new ExcelProcessIdentity(123, 134_381_234_567_890_123);
        var worker = new WorkerRegistration("test-pipe", self, self, "generation", WorkerRegistry.Build, excel);
        var admitted = new WorkerOperationEvidence(id.ToString("N"), operation.DispatchId.ToString("N"),
            command, session, excel, WorkerOperationEvidence.AdmittedPhase);
        receipts.BindWorker(operation, worker, null);
        receipts.ObserveWorkerEvidence(worker with { Generation = "earlier", OperationEvidence = admitted });
        receipts.ObserveWorkerEvidence(worker with { OperationEvidence = admitted with { SessionId = Guid.NewGuid().ToString("N") } });
        receipts.ObserveWorkerEvidence(worker with { OperationEvidence = admitted with { OperationId = Guid.NewGuid().ToString("N") } });
        Assert.Equal("controllerDispatchStarted", Lookup(receipts, id, session).GetProperty("operationEvidence").GetProperty("phase").GetString());

        receipts.ObserveWorkerEvidence(worker with { OperationEvidence = admitted });
        Assert.Equal("workerAdmitted", Lookup(receipts, id, session).GetProperty("operationEvidence").GetProperty("phase").GetString());
        var original = new ServiceResponse
        {
            ErrorCategory = "ExcelDisconnected",
            ExceptionType = "COMException",
            HResult = "0x800706BE",
            PartialState = """{"outcome":"unknown"}"""
        };
        var request = new ServiceRequest { Command = command, SessionId = session };
        receipts.Complete(operation, original.WithCompletionFailure(request, new IOException("worker bookkeeping"))
            .WithCompletionFailure(request, new InvalidOperationException("controller bookkeeping")));

        var response = admitted.WithResponse(original);
        receipts.ObserveWorkerEvidence(worker with { OperationEvidence = response });
        receipts.ObserveWorkerEvidence(worker with { OperationEvidence = admitted });
        receipts.ObserveNativeExit(new(excel with { StartedAtUtcFileTime = 456 }, ExcelNativeProcessState.Exited, 0, null));
        var before = Lookup(receipts, id, session);
        Assert.False(before.GetProperty("operationEvidence").TryGetProperty("nativeExit", out _));
        receipts.ObserveNativeExit(new(excel, ExcelNativeProcessState.Exited, null, null));
        receipts.ObserveNativeExit(new(excel, ExcelNativeProcessState.Exited, null, 5));
        receipts.ObserveNativeExit(new(excel, ExcelNativeProcessState.Exited, null, 87));
        receipts.ObserveNativeExit(new(excel, ExcelNativeProcessState.Exited, null, null));
        var exitWithoutCode = Lookup(receipts, id, session).GetProperty("operationEvidence").GetProperty("nativeExit");
        Assert.Equal(JsonValueKind.Null, exitWithoutCode.GetProperty("exitCode").ValueKind);
        Assert.Equal(5, exitWithoutCode.GetProperty("nativeError").GetInt32());
        receipts.ObserveNativeExit(new(excel, ExcelNativeProcessState.Exited, 0xC0000005u, null));
        receipts.ObserveNativeExit(new(excel, ExcelNativeProcessState.Exited, 0, null));

        var after = Lookup(receipts, id, session);
        Assert.Equal("unknown", after.GetProperty("state").GetString());
        Assert.Equal(before.GetProperty("finishedAt").GetString(), after.GetProperty("finishedAt").GetString());
        Assert.Equal("ExcelDisconnected", after.GetProperty("result").GetProperty("errorCategory").GetString());
        Assert.Contains("controller bookkeeping", after.GetProperty("recordingError").GetString(), StringComparison.Ordinal);
        var evidence = after.GetProperty("operationEvidence");
        Assert.Equal("TimeLogs.CommitThenRefresh", evidence.GetProperty("requestedProcedure").GetString());
        Assert.DoesNotContain("private fixture value", after.GetRawText(), StringComparison.Ordinal);
        Assert.Equal("serviceResponseRecorded", evidence.GetProperty("phase").GetString());
        Assert.Equal("134381234567890123", evidence.GetProperty("startedAtUtcFileTime").GetString());
        Assert.Equal("0x800706BE", evidence.GetProperty("workerFailure").GetProperty("hresult").GetString());
        Assert.Equal("0x800706BE", evidence.GetProperty("initialFailure").GetProperty("hresult").GetString());
        Assert.Equal(0xC0000005u, evidence.GetProperty("nativeExit").GetProperty("exitCode").GetUInt32());
        Assert.Equal(OperationReceipts.Admission.Duplicate, receipts.TryBegin(id, command, session, out _));

        clock.Now += TimeSpan.FromMinutes(15);
        Assert.Equal(OperationReceipts.Admission.Accepted, receipts.TryBegin(id, command, session, out var reused));
        receipts.BindWorker(reused, worker, null);
        receipts.ObserveWorkerEvidence(worker with { OperationEvidence = response });
        Assert.Equal("controllerDispatchStarted", Lookup(receipts, id, session).GetProperty("operationEvidence").GetProperty("phase").GetString());
        receipts.ObserveWorkerEvidence(worker with { OperationEvidence = admitted with { DispatchId = reused.DispatchId.ToString("N") } });
        Assert.Equal("workerAdmitted", Lookup(receipts, id, session).GetProperty("operationEvidence").GetProperty("phase").GetString());
    }

    [Theory]
    [InlineData("vba.invoke", """{"procedureName":"Book.xlsm!Module.Run"}""")]
    [InlineData("vba.invoke", """{"procedureName":"Module.Run","ProcedureName":"Other.Run"}""")]
    [InlineData("vba.invoke", """{"ProcedureName":"Module.Run"}""")]
    [InlineData("vba.invoke", """{"procedureName":"Module..Run"}""")]
    [InlineData("vba.invoke", """{"procedureName":null}""")]
    [InlineData("vba.invoke", """{"procedureName":"Module.Run\n"}""")]
    [InlineData("vba.invoke", """{"procedureName":""")]
    [InlineData("session.save", """{"procedureName":"Module.Run"}""")]
    public void ProcedureMetadataOmitsUnrecognizedSelectorsWithoutReplacingTheFailure(string command, string args)
    {
        var receipts = new OperationReceipts();
        var id = Guid.NewGuid();
        Assert.Equal(OperationReceipts.Admission.Accepted, receipts.TryBegin(id, command, "session", out var operation));
        receipts.RecordRequestedProcedure(operation, args);
        var failure = new ServiceResponse { ErrorCategory = "InvalidInput", ErrorMessage = "The original failure." };
        var response = receipts.IncludeEvidenceInFailure(operation, failure);
        Assert.Equal(failure.ErrorMessage, response.ErrorMessage);
        using var partial = JsonDocument.Parse(response.PartialState!);
        Assert.False(partial.RootElement.GetProperty("diagnostics").GetProperty("operationEvidence")
            .TryGetProperty("requestedProcedure", out _));
        Assert.False(Lookup(receipts, id, "session").GetProperty("operationEvidence")
            .TryGetProperty("requestedProcedure", out _));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task MissingLateReplyPreservesInitialFailureAndUnknownState(bool transportFault)
    {
        using var controller = new ExcelController(launcher: new("unused.exe", []));
        const System.Reflection.BindingFlags hidden = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
        var receipts = Assert.IsType<OperationReceipts>(typeof(ExcelController).GetField("_operationReceipts", hidden)!.GetValue(controller));
        var id = Guid.NewGuid();
        var session = Guid.NewGuid().ToString("N");
        Assert.Equal(OperationReceipts.Admission.Accepted, receipts.TryBegin(id, "vba.invoke", session, out var operation));
        var original = new ServiceResponse
        {
            ErrorCategory = "Timeout",
            ExceptionType = "TimeoutException",
            ErrorMessage = "Waiting on Excel.",
            PartialState = """{"outcome":"unknown"}"""
        };
        var wrapped = original.WithCompletionFailure(new ServiceRequest { Command = "vba.invoke", SessionId = session },
            new IOException("Could not record the timeout"));
        receipts.ObserveInitialFailure(operation, wrapped);
        var late = transportFault
            ? Task.FromException<ServiceResponse?>(new IOException("Lost late worker response"))
            : Task.FromResult<ServiceResponse?>(null);
        var capture = typeof(ExcelController).GetMethod("RecordLateOutcomeAsync", hidden)!;
        await (Task)capture.Invoke(controller, [operation, late, wrapped])!;
        var receipt = Lookup(receipts, id, session);
        Assert.Equal("unknown", receipt.GetProperty("state").GetString());
        Assert.Equal("Timeout", receipt.GetProperty("result").GetProperty("errorCategory").GetString());
        Assert.Equal("TimeoutException", receipt.GetProperty("operationEvidence").GetProperty("initialFailure")
            .GetProperty("exceptionType").GetString());
    }

    [Fact]
    public void Controller_RecoversRecordedEvidenceOnlyFromTheKnownWorkerGeneration()
    {
        const System.Reflection.BindingFlags hidden = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
        var directory = Directory.CreateTempSubdirectory("excelmcp-receipt-recovery-");
        var registry = new WorkerRegistry(directory.FullName);
        using var controller = new ExcelController(launcher: new("unused.exe", []), registry: registry);
        var receipts = Assert.IsType<OperationReceipts>(typeof(ExcelController).GetField("_operationReceipts", hidden)!.GetValue(controller));
        var generation = Assert.IsType<string>(typeof(ExcelController).GetField("_generation", hidden)!.GetValue(controller));
        var capture = typeof(ExcelController).GetMethod("CaptureRetainedOperationEvidence", hidden)!;
        var id = Guid.NewGuid();
        var session = Guid.NewGuid().ToString("N");
        const string command = "vba.invoke";
        var process = WorkerProcess.Current;
        var excel = new ExcelProcessIdentity(123, 134_381_234_567_890_123);
        var owner = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", process, process, generation, WorkerRegistry.Build, excel);
        Assert.Equal(OperationReceipts.Admission.Accepted, receipts.TryBegin(id, command, session, out var operation));
        receipts.BindWorker(operation, owner, null);
        var saved = owner with
        {
            OperationEvidence = new(id.ToString("N"), operation.DispatchId.ToString("N"), command, session, excel,
                WorkerOperationEvidence.ResponsePhase,
                "ExcelDisconnected", "COMException", "0x800706BE"),
            LastExcelObservation = new(excel, ExcelNativeProcessState.Exited, 0xC0000005u, null)
        };
        try
        {
            registry.Write(saved with { Generation = "previous-controller" });
            capture.Invoke(controller, [owner]);
            var before = Lookup(receipts, id, session).GetProperty("operationEvidence");
            Assert.Equal("controllerDispatchStarted", before.GetProperty("phase").GetString());
            Assert.False(before.TryGetProperty("nativeExit", out _));

            registry.Write(saved);
            capture.Invoke(controller, [owner]);
            var after = Lookup(receipts, id, session);
            Assert.Equal("pending", after.GetProperty("state").GetString());
            var evidence = after.GetProperty("operationEvidence");
            Assert.Equal("serviceResponseRecorded", evidence.GetProperty("phase").GetString());
            Assert.Equal("0x800706BE", evidence.GetProperty("workerFailure").GetProperty("hresult").GetString());
            Assert.Equal(0xC0000005u, evidence.GetProperty("nativeExit").GetProperty("exitCode").GetUInt32());
        }
        finally
        {
            registry.Remove(owner.PipeName);
            directory.Delete();
        }
    }

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
    [InlineData("ComInterop", "{\"diagnostics\":{\"execution\":{\"complete\":true,\"callbackReturned\":true}}}")]
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

    [Fact]
    public async Task LateVbaInvokeRetainsTheOriginalTypedResult()
    {
        const string session = "target";
        using var service = new ExcelMcpService();
        var completion = new TaskCompletionSource<VbaInvokeResult>(TaskCreationOptions.RunContinuationsAsynchronously);
        var request = new ServiceRequest
        {
            Command = "vba.invoke",
            SessionId = session,
            LateNativeCompletion = completion.Task
        };
        var receipts = new OperationReceipts();
        var id = Guid.NewGuid();
        Assert.Equal(OperationReceipts.Admission.Accepted,
            receipts.TryBegin(id, request.Command, session, out var operation));
        var response = service.ResolveLateNativeResponseAsync(request);
        Assert.False(response.IsCompleted);
        Assert.Equal("pending", Lookup(receipts, id, session).GetProperty("state").GetString());

        completion.SetResult(new VbaInvokeResult
        {
            Success = true,
            Action = "invoke",
            ReturnKind = "array",
            Values = [[false, 0], ["sample", null]]
        });
        receipts.Complete(operation, Assert.IsType<ServiceResponse>(await response));

        var receipt = Lookup(receipts, id, session);
        Assert.Equal("completed", receipt.GetProperty("state").GetString());
        Assert.Equal("vba.invoke", receipt.GetProperty("command").GetString());
        var result = receipt.GetProperty("result");
        Assert.Equal("array", result.GetProperty("returnKind").GetString());
        var values = result.GetProperty("values");
        Assert.False(values[0][0].GetBoolean());
        Assert.Equal(0, values[0][1].GetInt32());
        Assert.Equal("sample", values[1][0].GetString());
        Assert.Equal(JsonValueKind.Null, values[1][1].ValueKind);
        Assert.Equal("unknown", Lookup(receipts, id, "different-session").GetProperty("state").GetString());
        Assert.Equal(OperationReceipts.Admission.Duplicate,
            receipts.TryBegin(id, request.Command, session, out _));
    }

    [Fact]
    public async Task LateVbaDisconnectLeavesTheReceiptUnknown()
    {
        const string session = "target";
        using var service = new ExcelMcpService();
        var request = new ServiceRequest
        {
            Command = "vba.invoke",
            SessionId = session,
            LateNativeCompletion = Task.FromException<VbaInvokeResult>(
                Marshal.GetExceptionForHR(ResiliencePipelines.RPC_E_CALL_FAILED, new IntPtr(-1))!)
        };
        var receipts = new OperationReceipts();
        var id = Guid.NewGuid();
        Assert.Equal(OperationReceipts.Admission.Accepted,
            receipts.TryBegin(id, request.Command, session, out var operation));

        var response = Assert.IsType<ServiceResponse>(await service.ResolveLateNativeResponseAsync(request));
        receipts.Complete(operation, response);

        var receipt = Lookup(receipts, id, session);
        Assert.Equal("unknown", receipt.GetProperty("state").GetString());
        var result = receipt.GetProperty("result");
        Assert.False(result.GetProperty("success").GetBoolean());
        Assert.Equal("ExcelDisconnected", result.GetProperty("errorCategory").GetString());
        Assert.Equal("0x800706BE", result.GetProperty("hresult").GetString());
    }

    [Fact]
    public async Task LateVbaInvokeKeepsCompletionWhenTheReturnExceedsTheReceiptBudget()
    {
        const string session = "target";
        using var service = new ExcelMcpService();
        var value = new string('\u00E9', 5000);
        var request = new ServiceRequest
        {
            Command = "vba.invoke",
            SessionId = session,
            LateNativeCompletion = Task.FromResult(new VbaInvokeResult
            {
                Success = true,
                Action = "invoke",
                ReturnKind = "string",
                Value = value
            })
        };
        var receipts = new OperationReceipts();
        var id = Guid.NewGuid();
        Assert.Equal(OperationReceipts.Admission.Accepted,
            receipts.TryBegin(id, request.Command, session, out var operation));
        var response = Assert.IsType<ServiceResponse>(await service.ResolveLateNativeResponseAsync(request));
        using var payload = JsonDocument.Parse(response.Result!);
        Assert.Equal(value, payload.RootElement.GetProperty("value").GetString());

        receipts.Complete(operation, response);

        var receipt = Lookup(receipts, id, session);
        Assert.Equal("completed", receipt.GetProperty("state").GetString());
        Assert.False(receipt.TryGetProperty("result", out _));
        Assert.DoesNotContain(value, receipt.GetRawText(), StringComparison.Ordinal);
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
