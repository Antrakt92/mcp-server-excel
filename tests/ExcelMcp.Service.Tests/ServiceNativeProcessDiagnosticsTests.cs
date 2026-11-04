using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ExcelMcpServiceErrorTests
{
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
