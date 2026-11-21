using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

// Shares the error-transport fixture without starting Excel.
public sealed partial class ExcelMcpServiceErrorTests
{
    [Theory]
    [InlineData(true, true, null)]
    [InlineData(false, true, null)]
    [InlineData(null, true, "Could not observe the Excel save event.")]
    [InlineData(true, null, null)]
    public async Task ProcessAsync_SaveSeparatesReportedWriteFromUnsavedChanges(
        bool? excelReportedSaveSucceeded, bool? hasUnsavedChanges, string? observationError)
    {
        using var service = new ExcelMcpService();
        var nativeFailure = hasUnsavedChanges is null
            ? Assert.IsType<COMException>(Marshal.GetExceptionForHR(unchecked((int)0x800AC472))) : null;
        var failure = new WorkbookSaveIncompleteException("fixture.xlsx",
            excelReportedSaveSucceeded, observationError, hasUnsavedChanges, nativeFailure);
        var batch = new FakeBatch { SaveException = failure };
        const string sessionId = "save-with-unsaved-changes";
        const string command = "session.save";
        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest { Command = command, SessionId = sessionId });
        var result = ResultJson.Read(response.Result, command, response.Success, response.ErrorMessage,
            response.ErrorCategory, response.SessionId, response.ExceptionType, response.HResult,
            response.InnerError, response.PartialState);

        Assert.False(result.Success);
        Assert.Equal(1, batch.SaveCalls);
        Assert.Same(batch, service.SessionManager.GetSession(sessionId));
        Assert.Contains(failure.Message, response.ErrorMessage, StringComparison.Ordinal);
        using var partial = JsonDocument.Parse(Assert.IsType<string>(response.PartialState));
        var save = partial.RootElement.GetProperty("save");
        var dirty = save.GetProperty("hasUnsavedChanges");
        if (hasUnsavedChanges is { } remainingChanges)
            Assert.Equal(remainingChanges, dirty.GetBoolean());
        else
        {
            Assert.Equal(JsonValueKind.Null, dirty.ValueKind);
            Assert.Equal("Busy", response.ErrorCategory);
            Assert.Equal("0x800AC472", response.HResult);
            Assert.Contains(nativeFailure!.Message, response.InnerError, StringComparison.Ordinal);
        }
        var reported = save.GetProperty("excelReportedSaveSucceeded");
        if (excelReportedSaveSucceeded is { } succeeded)
            Assert.Equal(succeeded, reported.GetBoolean());
        else
            Assert.Equal(JsonValueKind.Null, reported.ValueKind);
        Assert.Equal(observationError, save.TryGetProperty("observationError", out var error) ? error.GetString() : null);
        foreach (var output in new[] { result.ToToolJson(), result.ToCliJson(command, 0) })
        {
            using var envelope = JsonDocument.Parse(output);
            var serializedSave = envelope.RootElement.GetProperty("partialState").GetProperty("save");
            Assert.Equal(save.GetRawText(), serializedSave.GetRawText());
        }
    }

    [Theory]
    [InlineData(false, unchecked((int)0x800A03EC), "ComInterop", false)]
    [InlineData(true, unchecked((int)0x800A03EC), "ComInterop", true)]
    [InlineData(true, unchecked((int)0x80010108), "ExcelDisconnected", true)]
    public async Task ProcessAsync_WorkbookSecurityUncertaintyControlsBatchContinuation(
        bool postDispatch, int hresult, string expectedCategory, bool mustStop)
    {
        using var service = new ExcelMcpService();
        var cause = CreateDiagnosticSnapshotFailure(hresult);
        const string message = "Synthetic workbook security transition failure.";
        var batch = new FakeBatch
        {
            ExecuteException = postDispatch
                ? new WorkbookSecurityCompletionUnconfirmedException(message, cause)
                : new InvalidOperationException(message, cause)
        };
        const string sessionId = "security-outcome-diagnostics";
        RegisterSession(service, sessionId, batch);

        // Feed a typed exception into the shared service error boundary without a live workbook.
        const string command = "sheet.list";
        var response = await service.ProcessAsync(new ServiceRequest { Command = command, SessionId = sessionId });
        var outcome = ResultJson.Read(response.Result, command, response.Success,
            response.ErrorMessage, response.ErrorCategory, response.SessionId, response.ExceptionType,
            response.HResult, response.InnerError, response.PartialState);

        Assert.False(response.Success);
        Assert.Equal(expectedCategory, response.ErrorCategory);
        Assert.Equal($"0x{hresult:X8}", response.HResult);
        Assert.Contains(cause.Message, response.InnerError);
        Assert.Equal(mustStop ? "unknown" : null, outcome.PartialState?["outcome"]?.GetValue<string>());
        Assert.Equal(mustStop, outcome.CompletionUnconfirmed);
    }

    [Fact]
    public async Task ProcessAsync_WorkbookOpeningSettingsFailureStopsBatchWithAllCauses()
    {
        using var service = new ExcelMcpService();
        var security = CreateDiagnosticSnapshotFailure(unchecked((int)0x800A03EC), "Synthetic AutomationSecurity restoration failure.");
        var alerts = CreateDiagnosticSnapshotFailure(unchecked((int)0x80070005), "Synthetic DisplayAlerts restoration failure.");
        var failure = new WorkbookSecurityCompletionUnconfirmedException("Excel opening settings are unconfirmed.",
            new AggregateException(security, alerts));
        const string sessionId = "opening-settings-diagnostics";
        const string command = "sheet.list";
        RegisterSession(service, sessionId, new FakeBatch { ExecuteException = failure });

        var response = await service.ProcessAsync(new ServiceRequest { Command = command, SessionId = sessionId });

        Assert.False(response.Success);
        Assert.Equal("ComInterop", response.ErrorCategory);
        Assert.Contains(security.Message, response.InnerError);
        Assert.Contains(alerts.Message, response.InnerError);
        var outcome = ResultJson.Read(response.Result, command, response.Success,
            response.ErrorMessage, response.ErrorCategory, response.SessionId, response.ExceptionType,
            response.HResult, response.InnerError, response.PartialState);
        Assert.True(outcome.CompletionUnconfirmed);
        Assert.Equal("unknown", outcome.PartialState?["outcome"]?.GetValue<string>());
    }

    [Theory]
    [InlineData(unchecked((int)0x800A03EC), "ComInterop")]
    [InlineData(unchecked((int)0x80010108), "ExcelDisconnected")]
    public async Task ProcessAsync_PostMutationFailureRetainsUnknownOutcomeAndNativeCause(int hresult, string expectedCategory)
    {
        using var service = new ExcelMcpService();
        var cause = CreateDiagnosticSnapshotFailure(hresult);
        var batch = new FakeBatch
        {
            ExecuteException = new OperationCompletionUnconfirmedException("Could not confirm worksheet deletion.", cause)
        };
        const string sessionId = "delete-follow-up-failure";
        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "sheet.delete",
            SessionId = sessionId,
            Args = """{"sheetName":"Target"}"""
        });

        Assert.False(response.Success);
        Assert.Equal(expectedCategory, response.ErrorCategory);
        Assert.Equal($"0x{hresult:X8}", response.HResult);
        Assert.Contains(cause.Message, response.InnerError);
        using var state = JsonDocument.Parse(Assert.IsType<string>(response.PartialState));
        Assert.Equal("unknown", state.RootElement.GetProperty("outcome").GetString());
    }

    [Theory]
    [InlineData(unchecked((int)0x800A03EC), false)]
    [InlineData(unchecked((int)0x800706BE), true)]
    public async Task ProcessAsync_MarkedAggregateAddsExecutionFactsWithoutChangingOriginalFailure(int hresult, bool disconnected)
    {
        using var service = new ExcelMcpService();
        const string secret = "synthetic-sensitive-operation-content";
        var operation = new InvalidOperationException(secret);
        var snapshot = CreateDiagnosticSnapshotFailure(hresult);
        var wire = new AggregateException(operation, snapshot);
        var batch = new FakeBatch { ExecuteException = wire };
        const string sessionId = "execution-diagnostics";
        RegisterSession(service, sessionId, batch);
        var request = new ServiceRequest { Command = "sheet.list", SessionId = sessionId };
        var baseline = await service.ProcessAsync(request);
        ExcelExecutionFailureContext.Attach(wire, operation, snapshot, callbackReturned: false);

        var response = await service.ProcessAsync(request);

        AssertOriginalDiagnosticErrorPreserved(baseline, response);
        // The ordinary classifier may choose the primary operation's category; baseline equality owns that contract.
        if (disconnected)
            Assert.Equal("ExcelDisconnected", response.ErrorCategory);
        Assert.Equal("sheet.list", response.Command);
        Assert.Equal(sessionId, response.SessionId);
        AssertExistingDiagnosticStatePreserved(baseline.PartialState, response.PartialState);
        var execution = ReadExecutionDiagnostics(response);
        Assert.True(execution.GetProperty("complete").GetBoolean());
        Assert.False(execution.GetProperty("callbackReturned").GetBoolean());
        Assert.Equal(nameof(InvalidOperationException), execution.GetProperty("operationExceptionType").GetString());
        Assert.Equal($"0x{operation.HResult:X8}", execution.GetProperty("operationHResult").GetString());
        Assert.Equal(nameof(COMException), execution.GetProperty("snapshotExceptionType").GetString());
        Assert.Equal($"0x{hresult:X8}", execution.GetProperty("snapshotHResult").GetString());
        Assert.DoesNotContain(secret, execution.GetRawText(), StringComparison.Ordinal);
        Assert.Equal(2, batch.ExecuteCalls);
        Assert.Equal(0, batch.SaveCalls);
        Assert.Equal(0, batch.DisposeCalls);
    }

    [Fact]
    public async Task ProcessAsync_SnapshotOnlyDoesNotInventAnOperationFailureOrSuccessfulSave()
    {
        using var service = new ExcelMcpService();
        var snapshot = CreateDiagnosticSnapshotFailure(unchecked((int)0x800A03EC));
        ExcelExecutionFailureContext.Attach(snapshot, null, snapshot, callbackReturned: true);
        var batch = new FakeBatch { ExecuteException = snapshot };
        RegisterSession(service, "snapshot-only-diagnostics", batch);

        var response = await service.ProcessAsync(new ServiceRequest { Command = "sheet.list", SessionId = "snapshot-only-diagnostics" });

        Assert.False(response.Success);
        Assert.Equal("sheet.list", response.Command);
        Assert.Equal("0x800A03EC", response.HResult);
        var execution = ReadExecutionDiagnostics(response);
        Assert.True(execution.GetProperty("complete").GetBoolean());
        Assert.True(execution.GetProperty("callbackReturned").GetBoolean());
        Assert.Equal(nameof(COMException), execution.GetProperty("snapshotExceptionType").GetString());
        Assert.Equal("0x800A03EC", execution.GetProperty("snapshotHResult").GetString());
        Assert.False(execution.TryGetProperty("operationExceptionType", out _));
        Assert.False(execution.TryGetProperty("operationHResult", out _));
        Assert.False(execution.TryGetProperty("saved", out _));
        Assert.Equal(0, batch.SaveCalls);
    }

    [Fact]
    public async Task ProcessAsync_UnmarkedAggregateDoesNotInferExecutionPhases()
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch
        {
            ExecuteException = new AggregateException(
                new InvalidOperationException("operation"), CreateDiagnosticSnapshotFailure(unchecked((int)0x800A03EC)))
        };
        RegisterSession(service, "unmarked-diagnostics", batch);

        var response = await service.ProcessAsync(new ServiceRequest { Command = "sheet.list", SessionId = "unmarked-diagnostics" });

        Assert.False(response.Success);
        Assert.Equal("sheet.list", response.Command);
        if (response.PartialState is { } state)
        {
            var partial = JsonSerializer.Deserialize<JsonElement>(state);
            Assert.False(partial.TryGetProperty("diagnostics", out var diagnostics) && diagnostics.TryGetProperty("execution", out _));
        }
        Assert.Equal(1, batch.ExecuteCalls);
    }

    [Theory]
    [InlineData("edit", false)]
    [InlineData("protection", false)]
    [InlineData("authentication", false)]
    [InlineData("edit", true)]
    [InlineData("protection", true)]
    [InlineData("authentication", true)]
    public async Task ProcessAsync_ExecutionDiagnosticsPreserveObservedVbaState(string kind, bool disconnected)
    {
        using var service = new ExcelMcpService();
        var cause = new InvalidOperationException("operation");
        object observedState = kind switch
        {
            "edit" => new VbaSourceEditState { ModuleName = "Module1", Phase = "verifying", LineCount = 7, SourceHash = "synthetic-fingerprint" },
            "protection" => new VbaProjectProtectionState { Phase = "restoring-visibility", ConfigurationApplied = true, SaveRequired = true, ReopenRequired = true },
            _ => new VbaProjectAuthenticationState { Phase = "restoring-visibility", PasswordVerified = true, AccessUnlocked = true, AuthenticationPerformed = true, Saved = false, Reopened = false }
        };
        Exception operation = observedState switch
        {
            VbaSourceEditState state => new VbaSourceEditException("edit", cause, state),
            VbaProjectProtectionState state => new VbaProjectProtectionException("protection", cause, state),
            VbaProjectAuthenticationState state => new VbaProjectAuthenticationException("authentication", cause, state),
            _ => throw new InvalidOperationException("Unknown fixture state.")
        };
        var snapshot = CreateDiagnosticSnapshotFailure(disconnected ? unchecked((int)0x800706BE) : unchecked((int)0x800A03EC));
        var wire = new AggregateException(operation, snapshot);
        ExcelExecutionFailureContext.Attach(wire, operation, snapshot, callbackReturned: false);
        var batch = new FakeBatch { ExecuteException = wire };
        RegisterSession(service, "vba-execution-diagnostics", batch);

        var response = await service.ProcessAsync(new ServiceRequest { Command = "sheet.list", SessionId = "vba-execution-diagnostics" });

        Assert.False(response.Success);
        var partial = JsonSerializer.Deserialize<JsonElement>(Assert.IsType<string>(response.PartialState));
        foreach (var property in JsonSerializer.SerializeToElement(observedState, ServiceProtocol.JsonOptions).EnumerateObject())
            Assert.True(JsonElement.DeepEquals(property.Value, partial.GetProperty(property.Name)));
        if (disconnected)
        {
            Assert.Equal("ExcelDisconnected", response.ErrorCategory);
            Assert.Equal("unknown", partial.GetProperty("outcome").GetString());
        }
        var execution = ReadExecutionDiagnostics(response);
        Assert.True(execution.GetProperty("complete").GetBoolean());
        Assert.False(execution.GetProperty("callbackReturned").GetBoolean());
        Assert.Equal(operation.GetType().Name, execution.GetProperty("operationExceptionType").GetString());
        Assert.Equal(0, batch.SaveCalls);
        Assert.Equal(0, batch.DisposeCalls);
    }

    [Fact]
    public async Task ProcessAsync_DeadDisconnectRetainsExecutionDiagnosticsAfterSessionCleanup()
    {
        using var service = new ExcelMcpService();
        var operation = new InvalidOperationException("save operation");
        var snapshot = CreateDiagnosticSnapshotFailure(unchecked((int)0x800706BE));
        var wire = new AggregateException(operation, snapshot);
        ExcelExecutionFailureContext.Attach(wire, operation, snapshot, callbackReturned: false);
        var batch = new FakeBatch { SaveException = wire, IsAliveAfterSaveException = false };
        RegisterSession(service, "dead-execution-diagnostics", batch);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = "dead-execution-diagnostics",
            Args = "{\"save\":true}"
        });

        Assert.False(response.Success);
        Assert.Equal("ExcelProcessDied", response.ErrorCategory);
        Assert.Equal("session.close", response.Command);
        Assert.Equal("dead-execution-diagnostics", response.SessionId);
        Assert.Equal("0x800706BE", response.HResult);
        var execution = ReadExecutionDiagnostics(response);
        Assert.True(execution.GetProperty("complete").GetBoolean());
        Assert.False(execution.GetProperty("callbackReturned").GetBoolean());
        Assert.Equal(nameof(InvalidOperationException), execution.GetProperty("operationExceptionType").GetString());
        Assert.Equal("0x800706BE", execution.GetProperty("snapshotHResult").GetString());
        Assert.Equal(1, batch.SaveCalls);
        Assert.Equal(1, batch.DisposeCalls);
    }

    [Theory]
    [InlineData("ambiguous")]
    [InlineData("deep-unfound")]
    [InlineData("deep-found")]
    public async Task ProcessAsync_IncompleteContextDoesNotChooseAuthoritativePhases(string kind)
    {
        using var service = new ExcelMcpService();
        var marked = new InvalidOperationException("snapshot");
        ExcelExecutionFailureContext.Attach(marked, null, marked, callbackReturned: true);
        Exception wire;
        if (kind == "ambiguous")
        {
            var other = new InvalidOperationException("different snapshot");
            ExcelExecutionFailureContext.Attach(other, null, other, callbackReturned: false);
            wire = new AggregateException(marked, other);
        }
        else
        {
            wire = marked;
            for (var index = 0; index < 40; index++)
                wire = new InvalidOperationException("wrapper", wire);
            if (kind == "deep-found")
                ExcelExecutionFailureContext.Attach(wire, null, wire, callbackReturned: true);
        }
        var batch = new FakeBatch { ExecuteException = wire };
        RegisterSession(service, "incomplete-execution-diagnostics", batch);

        var response = await service.ProcessAsync(new ServiceRequest { Command = "sheet.list", SessionId = "incomplete-execution-diagnostics" });

        Assert.False(response.Success);
        var execution = ReadExecutionDiagnostics(response);
        Assert.False(execution.GetProperty("complete").GetBoolean());
        Assert.Single(execution.EnumerateObject());
    }

    private static COMException CreateDiagnosticSnapshotFailure(int hresult, string message = "Synthetic snapshot failure.")
    {
#pragma warning disable CA2201 // Synthetic HRESULT exercises Service transport; no native Excel is used.
        return new COMException(message, hresult);
#pragma warning restore CA2201
    }

    private static JsonElement ReadExecutionDiagnostics(ServiceResponse response) =>
        JsonSerializer.Deserialize<JsonElement>(Assert.IsType<string>(response.PartialState))
            .GetProperty("diagnostics").GetProperty("execution");

    private static void AssertExistingDiagnosticStatePreserved(string? baseline, string? actual)
    {
        if (baseline is null)
            return;
        var original = JsonSerializer.Deserialize<JsonElement>(baseline);
        var retained = JsonSerializer.Deserialize<JsonElement>(Assert.IsType<string>(actual));
        foreach (var property in original.EnumerateObject())
            Assert.True(JsonElement.DeepEquals(property.Value, retained.GetProperty(property.Name)));
    }

    private static void AssertOriginalDiagnosticErrorPreserved(ServiceResponse baseline, ServiceResponse actual)
    {
        Assert.False(baseline.Success);
        Assert.False(actual.Success);
        Assert.Equal(baseline.Command, actual.Command);
        Assert.Equal(baseline.SessionId, actual.SessionId);
        Assert.Equal(baseline.ErrorCategory, actual.ErrorCategory);
        Assert.Equal(baseline.ErrorMessage, actual.ErrorMessage);
        Assert.Equal(baseline.ExceptionType, actual.ExceptionType);
        Assert.Equal(baseline.HResult, actual.HResult);
        Assert.Equal(baseline.InnerError, actual.InnerError);
        Assert.Equal(baseline.Result, actual.Result);
    }
}
