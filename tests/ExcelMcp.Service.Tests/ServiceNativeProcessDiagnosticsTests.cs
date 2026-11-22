using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service.Workers;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ExcelMcpServiceErrorTests
{
    [Theory]
    [InlineData("Exited", 17u, false)]
    [InlineData("Exited", 17u, true)]
    [InlineData("Exited", null, false)]
    [InlineData("Indeterminate", null, false)]
    [InlineData("Alive", null, false)]
    public void WorkerSnapshot_RecordsConfirmedExcelExitBeforeOmittingSession(
        string stateName, uint? exitCode, bool registryUnavailable)
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-status-exit-");
        try
        {
            using var service = new ExcelMcpService();
            var batch = new FakeBatch
            {
                ExecuteException = new InvalidOperationException("Synthetic Excel process stopped responding."),
                IsAliveAfterExecuteException = false
            };
            const string sessionId = "exit-observed-during-status";
            RegisterSession(service, sessionId, batch);
            var identity = new ExcelProcessIdentity(batch.ExcelProcessId!.Value, 456);
            var state = Enum.Parse<ExcelNativeProcessState>(stateName);
            var observation = new ExcelNativeProcessObservation(identity, state, exitCode,
                state != ExcelNativeProcessState.Alive && exitCode is null ? 5 : null);
            var process = WorkerProcess.Current;
            var generation = Guid.NewGuid().ToString("N");
            var registration = new WorkerRegistration("excelmcp-worker-" + generation, process, process,
                generation, WorkerRegistry.Build, identity, OwnsExcel: true);
            var registry = new WorkerRegistry(directory.FullName);
            registry.Write(registration);
            using var host = new ExcelWorkerHost(registration, service, registry);
            var refresh = typeof(ExcelWorkerHost).GetMethod("TryRefreshSnapshot", BindingFlags.Instance | BindingFlags.NonPublic)!;
            refresh.Invoke(host, null);
            Assert.Null(GetPrivateField<WorkerStatusFailure?>(host, "_statusFailure"));
            Assert.Single(GetPrivateField<WorkerSnapshot>(host, "_snapshot").Sessions);

            Assert.Same(batch.ExecuteException, Assert.Throws<InvalidOperationException>(() => batch.Execute(static (_, _) => { })));
            batch.ProcessObservation = observation;
            var temporary = Path.Join(directory.FullName, registration.PipeName + ".json.tmp");
            if (registryUnavailable)
            {
                Directory.CreateDirectory(temporary);
                refresh.Invoke(host, null);
                var failure = GetPrivateField<WorkerStatusFailure>(host, "_statusFailure");
                Assert.Contains("Excel status could not be refreshed", failure.Message);
                Assert.Equal("snapshotRefresh", failure.Diagnostic!.Stage);
                Assert.Single(GetPrivateField<WorkerSnapshot>(host, "_snapshot").Sessions);
                Assert.Equal(registration, registry.Read(registration.PipeName));
                Directory.Delete(temporary);
            }
            refresh.Invoke(host, null);
            Assert.Null(GetPrivateField<WorkerStatusFailure?>(host, "_statusFailure"));

            var snapshot = GetPrivateField<WorkerSnapshot>(host, "_snapshot");
            var expectedExit = state == ExcelNativeProcessState.Exited ? observation : null;
            Assert.Equal(expectedExit, snapshot.Registration.LastExcelObservation);
            Assert.Equal(snapshot.Registration, registry.Read(registration.PipeName));
            Assert.Equal(identity, snapshot.Registration.Excel);
            Assert.True(snapshot.Registration.OwnsExcel);
            Assert.False(snapshot.Registration.ReleaseConfirmed);
            Assert.Empty(snapshot.Sessions);
            using var list = JsonDocument.Parse(snapshot.SessionList.Result!);
            Assert.Empty(list.RootElement.GetProperty("sessions").EnumerateArray());
            Assert.Contains(sessionId, service.SessionManager.ActiveSessionIds);
            Assert.Equal(1, batch.ExecuteCalls);
            Assert.Equal(0, batch.SaveCalls);
            Assert.Equal(0, batch.DisposeCalls);
            if (registryUnavailable)
            {
                // The same exit is already durable, so later polls do not write again.
                Directory.CreateDirectory(temporary);
                refresh.Invoke(host, null);
                Assert.Null(GetPrivateField<WorkerStatusFailure?>(host, "_statusFailure"));
                Assert.Equal(expectedExit, GetPrivateField<WorkerSnapshot>(host, "_snapshot").Registration.LastExcelObservation);
                Assert.Equal(snapshot.Registration, registry.Read(registration.PipeName));
            }
        }
        finally
        {
            directory.Delete(recursive: true);
        }
    }

    [Fact]
    public void NativeDiagnostics_CreationIdentityPreservesDigitsBeyondJavascriptIntegerPrecision()
    {
        const long creationTime = 134_381_234_567_890_123;
        var observation = new ExcelNativeProcessObservation(new(123, creationTime), ExcelNativeProcessState.Exited, 0, null);
        using var document = JsonDocument.Parse(MergeNativeDiagnosticsForTest(null, observation)!);

        var encoded = document.RootElement.GetProperty("diagnostics").GetProperty("nativeProcess")
            .GetProperty("startedAtUtcFileTime");
        Assert.Equal(JsonValueKind.String, encoded.ValueKind);
        Assert.Equal("134381234567890123", encoded.GetString());
        Assert.Equal(creationTime, long.Parse(encoded.GetString()!, System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public async Task ProcessAsync_FatalDisconnectKeepsNativeExitAfterSessionRemoval()
    {
        using var service = new ExcelMcpService();
#pragma warning disable CA2201 // Models a fatal callback and native exit without launching Excel.
        var error = new COMException("Synthetic fatal callback disconnect.", unchecked((int)0x800706BE));
#pragma warning restore CA2201
        var observation = new ExcelNativeProcessObservation(new(123, 456), ExcelNativeProcessState.Exited, 17, null);
        var batch = new FakeBatch
        {
            ExecuteException = error,
            IsAliveAfterExecuteException = false,
            ProcessObservation = observation
        };
        const string sessionId = "fatal-native-exit-diagnostics";
        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest { Command = "sheet.list", SessionId = sessionId });

        Assert.False(response.Success);
        Assert.Equal("ExcelProcessDied", response.ErrorCategory);
        Assert.Equal("0x800706BE", response.HResult);
        Assert.Equal("sheet.list", response.Command);
        Assert.Equal(sessionId, response.SessionId);
        Assert.DoesNotContain(sessionId, service.SessionManager.ActiveSessionIds);
        using var state = JsonDocument.Parse(Assert.IsType<string>(response.PartialState));
        var native = AssertNativeDiagnosticsIdentity(state, "Exited");
        Assert.Equal(17u, native.GetProperty("exitCode").GetUInt32());
        Assert.Equal(JsonValueKind.Null, native.GetProperty("nativeError").ValueKind);
    }

    [Theory]
    [InlineData("sheet.list", null)]
    [InlineData("vba.invoke", """{"procedureName":"SyntheticProbe.Callback"}""")]
    public async Task ProcessAsync_ExcelExitsDuringDisconnectCheck_ReportsTheObservedExit(string command, string? args)
    {
        using var service = new ExcelMcpService();
#pragma warning disable CA2201 // Simulates a COM disconnect followed by exit during the native liveness check.
        var error = new COMException("Synthetic disconnect.", unchecked((int)0x800706BE));
#pragma warning restore CA2201
        var identity = new ExcelProcessIdentity(123, 456);
        var batch = new FakeBatch
        {
            WorkbookPath = Path.Combine(Path.GetTempPath(), "synthetic-disconnect.xlsm"),
            ExecuteException = error,
            IsAliveAfterExecuteException = false,
            ProcessObservation = new(identity, ExcelNativeProcessState.Alive, null, null),
            ProcessObservationAfterDeathCheck = new(identity, ExcelNativeProcessState.Exited, 17, null)
        };
        const string sessionId = "exit-during-disconnect-check";
        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest { Command = command, SessionId = sessionId, Args = args });

        Assert.False(response.Success);
        Assert.Equal("ExcelProcessDied", response.ErrorCategory);
        Assert.Equal("0x800706BE", response.HResult);
        Assert.Equal(command, response.Command);
        Assert.Equal(sessionId, response.SessionId);
        Assert.DoesNotContain(sessionId, service.SessionManager.ActiveSessionIds);
        using var state = JsonDocument.Parse(Assert.IsType<string>(response.PartialState));
        var native = AssertNativeDiagnosticsIdentity(state, "Exited");
        Assert.Equal(17u, native.GetProperty("exitCode").GetUInt32());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ProcessAsync_ExcelExitedBeforeVbaInvoke_ReportsExitWithoutSubmittingTheMacro(bool previousWaitExpired)
    {
        using var service = new ExcelMcpService();
        var identity = new ExcelProcessIdentity(123, 456);
        var batch = new FakeBatch
        {
            WorkbookPath = Path.Combine(Path.GetTempPath(), "synthetic-dead-preflight.xlsm"),
            ExecuteException = new InvalidOperationException("Synthetic process exit after earlier work."),
            IsAliveAfterExecuteException = false,
            ProcessObservationAfterDeathCheck = new(identity, ExcelNativeProcessState.Exited, 17, null)
        };
        const string sessionId = "native-exit-before-vba-invoke";
        RegisterSession(service, sessionId, batch);
        Assert.Same(batch.ExecuteException, Assert.Throws<InvalidOperationException>(() => batch.Execute(static (_, _) => { })));
        batch.HasTimedOutOperation = previousWaitExpired;

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "vba.invoke",
            SessionId = sessionId,
            Args = """{"procedureName":"SyntheticProbe.Callback"}"""
        });

        Assert.False(response.Success);
        Assert.Equal("ExcelProcessDied", response.ErrorCategory);
        Assert.Equal("vba.invoke", response.Command);
        Assert.Equal(sessionId, response.SessionId);
        Assert.Null(response.HResult);
        Assert.Equal(1, batch.ExecuteCalls);
        Assert.Equal(1, batch.DisposeCalls);
        Assert.DoesNotContain(sessionId, service.SessionManager.ActiveSessionIds);
        using var state = JsonDocument.Parse(Assert.IsType<string>(response.PartialState));
        var native = AssertNativeDiagnosticsIdentity(state, "Exited");
        Assert.Equal(17u, native.GetProperty("exitCode").GetUInt32());
    }

    [Fact]
    public async Task ProcessAsync_ReadinessComErrorWithLiveExcel_PreservesOriginalError()
    {
        using var service = new ExcelMcpService();
        var error = NewReadinessComError();
        var batch = new FakeBatch
        {
            HasTimedOutOperation = true,
            RefreshException = error
        };
        const string sessionId = "live-excel-during-readiness-check";
        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest { Command = "sheet.list", SessionId = sessionId });

        Assert.False(response.Success);
        Assert.Equal("ComInterop", response.ErrorCategory);
        Assert.Equal("0x800706BE", response.HResult);
        Assert.Contains(error.Message, response.ErrorMessage);
        Assert.Equal(0, batch.ExecuteCalls);
        Assert.Equal(0, batch.DisposeCalls);
        Assert.Contains(sessionId, service.SessionManager.ActiveSessionIds);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ProcessAsync_ExcelExitsDuringReadinessCheck_ReportsConfirmedExit(bool scanThrows)
    {
        using var service = new ExcelMcpService();
        var identity = new ExcelProcessIdentity(123, 456);
        var batch = new FakeBatch
        {
            HasTimedOutOperation = true,
            ExitDuringRefresh = true,
            RefreshException = scanThrows ? NewReadinessComError() : null,
            ProcessObservationAfterDeathCheck = new(identity, ExcelNativeProcessState.Exited, 17, null)
        };
        const string sessionId = "excel-exit-during-readiness-check";
        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest { Command = "sheet.list", SessionId = sessionId });

        Assert.False(response.Success);
        Assert.Equal("ExcelProcessDied", response.ErrorCategory);
        Assert.Equal("sheet.list", response.Command);
        Assert.Equal(sessionId, response.SessionId);
        Assert.Equal(0, batch.ExecuteCalls);
        Assert.Equal(1, batch.DisposeCalls);
        Assert.DoesNotContain(sessionId, service.SessionManager.ActiveSessionIds);
        using var state = JsonDocument.Parse(Assert.IsType<string>(response.PartialState));
        var native = AssertNativeDiagnosticsIdentity(state, "Exited");
        Assert.Equal(17u, native.GetProperty("exitCode").GetUInt32());
    }

    [Theory]
    [InlineData(true, true)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public async Task ProcessAsync_FatalCloseDisconnect_ReportsSnapshotAndRecoverySteps(bool alive, bool hasObserver)
    {
        using var service = new ExcelMcpService();
#pragma warning disable CA2201 // Simulates a fatal COM disconnect during workbook close without launching Excel.
        var rpcError = new COMException("The RPC server is unavailable.", unchecked((int)0x800706BA));
#pragma warning restore CA2201
        var error = new InvalidOperationException("Workbook close failed. Excel has not been asked to quit.", rpcError);
        var identity = new ExcelProcessIdentity(123, 456);
        var batch = new FakeBatch
        {
            ExecuteException = error,
            IsAliveAfterExecuteException = alive,
            ProcessObservation = hasObserver ? new(identity, ExcelNativeProcessState.Alive, null, null) : null,
            ProcessObservationAfterDeathCheck = !alive ? new(identity, ExcelNativeProcessState.Exited, 0, null) : null
        };
        const string sessionId = "fatal-close-disconnect";
        RegisterSession(service, sessionId, batch);
        service.SessionManager.BeforeCloseSessionHookForTests = () => batch.Execute(static (_, _) => { });

        var response = await service.ProcessAsync(new ServiceRequest { Command = "session.close", SessionId = sessionId });

        Assert.False(response.Success);
        Assert.Equal(alive ? "ExcelDisconnected" : "ExcelProcessDied", response.ErrorCategory);
        Assert.Equal("0x800706BA", response.HResult);
        Assert.Equal("session.close", response.Command);
        Assert.Equal(sessionId, response.SessionId);
        Assert.Contains(rpcError.Message, response.InnerError);
        Assert.Contains("Closing and any requested save are unconfirmed", response.ErrorMessage);
        Assert.Contains("session list", response.ErrorMessage);
        Assert.Contains("session discover", response.ErrorMessage);
        Assert.Contains("file_read", response.ErrorMessage);
        Assert.Contains("saved file", response.ErrorMessage);
        Assert.Equal(alive, service.SessionManager.ActiveSessionIds.Contains(sessionId));
        Assert.Equal(1, batch.ExecuteCalls);
        Assert.Equal(alive ? 0 : 1, batch.DisposeCalls);
        Assert.Equal(0, batch.SaveCalls);
        if (hasObserver)
        {
            Assert.Contains("same controller", response.ErrorMessage);
            Assert.Contains("recentExcelExits", response.ErrorMessage);
            Assert.Contains("startedAtUtcFileTime", response.ErrorMessage);
            using var state = JsonDocument.Parse(Assert.IsType<string>(response.PartialState));
            AssertNativeDiagnosticsIdentity(state, alive ? "Alive" : "Exited");
        }
        else
        {
            Assert.DoesNotContain("recentExcelExits", response.ErrorMessage);
            Assert.DoesNotContain("process was observed alive", response.ErrorMessage);
        }
        if (alive)
        {
            Assert.Contains("The session was still tracked", response.ErrorMessage);
            Assert.Contains(error.Message, response.ErrorMessage);
        }
        if (alive && hasObserver)
            Assert.Contains("process was observed alive when the close error was handled", response.ErrorMessage);
    }

    [Theory]
    [InlineData("Alive")]
    [InlineData("Exited")]
    [InlineData("Indeterminate")]
    [InlineData(null)]
    public async Task ProcessAsync_CloseWithUnknownReadiness_ReportsUnconfirmedOutcome(string? stateName)
    {
        using var service = new ExcelMcpService();
        var batch = new FakeBatch
        {
            RefreshState = WorkbookRefreshState.Unknown,
            ExitDuringRefresh = stateName == "Exited",
            ProcessObservation = stateName is null ? null : new(
                new(123, 456), Enum.Parse<ExcelNativeProcessState>(stateName),
                stateName == "Exited" ? 17u : null, stateName == "Indeterminate" ? 5 : null)
        };
        const string sessionId = "close-with-unknown-readiness";
        RegisterSession(service, sessionId, batch);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = """{"save":true}"""
        });

        Assert.False(response.Success);
        Assert.Equal("Busy", response.ErrorCategory);
        Assert.Equal("ExcelBusyException", response.ExceptionType);
        Assert.Equal("session.close", response.Command);
        Assert.Equal(sessionId, response.SessionId);
        Assert.Contains("refresh state could not be confirmed", response.ErrorMessage);
        Assert.DoesNotContain("workbook remains open", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("session list", response.ErrorMessage);
        Assert.Contains("file_read", response.ErrorMessage);
        Assert.Equal(0, batch.SaveCalls);
        Assert.Equal(0, batch.DisposeCalls);
        Assert.Contains(sessionId, service.SessionManager.ActiveSessionIds);

        using var state = JsonDocument.Parse(Assert.IsType<string>(response.PartialState));
        Assert.Equal("unknown", state.RootElement.GetProperty("outcome").GetString());
        if (stateName is null)
        {
            Assert.False(state.RootElement.TryGetProperty("diagnostics", out _));
            return;
        }
        var native = AssertNativeDiagnosticsIdentity(state, stateName);
        Assert.Equal(stateName == "Exited" ? 17u : (uint?)null, native.GetProperty("exitCode").Deserialize<uint?>());
        Assert.Equal(stateName == "Indeterminate" ? 5 : (int?)null, native.GetProperty("nativeError").Deserialize<int?>());
    }

#pragma warning disable CA2201 // Simulates a COM error from the readiness scan without starting Excel.
    private static COMException NewReadinessComError() =>
        new("Synthetic readiness scan COM error.", unchecked((int)0x800706BE));
#pragma warning restore CA2201

    [Fact]
    public void NativeDiagnostics_NoObservationPreservesExistingBytesAndNull()
    {
        const string original = " { \"phase\": \"verifying\", \"diagnostics\": { \"execution\": { \"complete\": false } } } ";
        Assert.Equal(original, MergeNativeDiagnosticsForTest(original, null));
        Assert.Null(MergeNativeDiagnosticsForTest(null, null));
    }

    [Theory]
    [InlineData(0u)]
    [InlineData(259u)]
    [InlineData(uint.MaxValue)]
    public void NativeDiagnostics_ExitedPreservesUnsignedCodeAndExplicitUnknownError(uint exitCode)
    {
        var observation = new ExcelNativeProcessObservation(new(123, 456), ExcelNativeProcessState.Exited,
            exitCode, null);

        using var document = JsonDocument.Parse(MergeNativeDiagnosticsForTest(null, observation)!);

        var native = AssertNativeDiagnosticsIdentity(document, "Exited");
        Assert.Equal(exitCode, native.GetProperty("exitCode").GetUInt32());
        Assert.Equal(JsonValueKind.Null, native.GetProperty("nativeError").ValueKind);
    }

    [Fact]
    public void NativeDiagnostics_ExitedQueryFailureDoesNotBecomeZero()
    {
        var observation = new ExcelNativeProcessObservation(new(123, 456), ExcelNativeProcessState.Exited,
            null, 5);

        using var document = JsonDocument.Parse(MergeNativeDiagnosticsForTest(null, observation)!);

        var native = AssertNativeDiagnosticsIdentity(document, "Exited");
        Assert.Equal(JsonValueKind.Null, native.GetProperty("exitCode").ValueKind);
        Assert.Equal(5, native.GetProperty("nativeError").GetInt32());
    }

    [Theory]
    [InlineData(0, "Alive")]
    [InlineData(2, "Indeterminate")]
    public void NativeDiagnostics_NonterminalStateDoesNotInventSuccessfulExit(int stateValue, string expectedState)
    {
        var state = (ExcelNativeProcessState)stateValue;
        var observation = new ExcelNativeProcessObservation(new(123, 456), state,
            null, state == ExcelNativeProcessState.Indeterminate ? 5 : null);

        using var document = JsonDocument.Parse(MergeNativeDiagnosticsForTest(null, observation)!);

        var native = AssertNativeDiagnosticsIdentity(document, expectedState);
        Assert.Equal(JsonValueKind.Null, native.GetProperty("exitCode").ValueKind);
        if (state == ExcelNativeProcessState.Indeterminate)
            Assert.Equal(5, native.GetProperty("nativeError").GetInt32());
        else
            Assert.Equal(JsonValueKind.Null, native.GetProperty("nativeError").ValueKind);
    }

    [Fact]
    public void NativeDiagnostics_MergePreservesVbaOutcomeAndExecutionDiagnostics()
    {
        const string existing = """
            {"moduleName":"SyntheticModule","phase":"verifying","sourceHash":"ABC","lineCount":2,
             "outcome":"unknown","diagnostics":{"execution":{"complete":true,"callbackReturned":true,
             "snapshotExceptionType":"COMException","snapshotHResult":"0x80010108"},"other":{"complete":false}}}
            """;
        var observation = new ExcelNativeProcessObservation(new(123, 456), ExcelNativeProcessState.Exited, 17, null);

        using var document = JsonDocument.Parse(MergeNativeDiagnosticsForTest(existing, observation)!);
        using var previous = JsonDocument.Parse(existing);

        var root = document.RootElement;
        foreach (var field in new[] { "moduleName", "phase", "sourceHash", "lineCount", "outcome" })
            Assert.True(JsonElement.DeepEquals(previous.RootElement.GetProperty(field), root.GetProperty(field)));
        var diagnostics = root.GetProperty("diagnostics");
        foreach (var field in new[] { "execution", "other" })
            Assert.True(JsonElement.DeepEquals(previous.RootElement.GetProperty("diagnostics").GetProperty(field),
                diagnostics.GetProperty(field)));
        var native = AssertNativeDiagnosticsIdentity(document, "Exited");
        Assert.Equal(17u, native.GetProperty("exitCode").GetUInt32());
        Assert.DoesNotContain("SyntheticModule", native.GetRawText(), StringComparison.Ordinal);
    }

    private static string? MergeNativeDiagnosticsForTest(string? state, ExcelNativeProcessObservation? observation) =>
        (string?)typeof(ExcelMcpService).GetMethod("MergeNativeProcessDiagnostics", BindingFlags.Static | BindingFlags.NonPublic)!
            .Invoke(null, [state, observation]);

    private static JsonElement AssertNativeDiagnosticsIdentity(JsonDocument document, string expectedState)
    {
        var native = document.RootElement.GetProperty("diagnostics").GetProperty("nativeProcess");
        Assert.Equal(5, native.EnumerateObject().Count());
        Assert.Equal(123, native.GetProperty("processId").GetInt32());
        Assert.Equal("456", native.GetProperty("startedAtUtcFileTime").GetString());
        Assert.Equal(expectedState, native.GetProperty("state").GetString());
        return native;
    }
}
