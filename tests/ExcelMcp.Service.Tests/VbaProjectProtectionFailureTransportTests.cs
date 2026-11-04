using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "VBA")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class VbaProjectProtectionFailureTransportTests
{
    private static readonly string[] RetainedFiles = ["synthetic.xlsm"];

    [Theory]
    [InlineData("direct")]
    [InlineData("wrapped")]
    [InlineData("aggregate")]
    public void FailureAfterAcceptance_PreservesProtectionStateAndIndependentFailure(string wrapping)
    {
        using var service = new ExcelMcpService();
        var failure = new VbaProjectProtectionException("Protection configuration needs inspection.",
            new InvalidOperationException("Original restoration failure."),
            new VbaProjectProtectionState
            {
                Phase = "restoring-visibility",
                AcceptanceRequested = true,
                ConfigurationApplied = true,
                LockOptionReadback = true,
                DialogCleanupConfirmed = true,
                ActiveProjectRestored = true,
                VisibilityRestored = false,
                SaveRequired = true,
                ReopenRequired = true
            });
        var response = CreateError(service, Wrap(failure, wrapping));
        var shared = JsonSerializer.Deserialize<Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceResponse>(
            JsonSerializer.Serialize(response, ServiceProtocol.JsonOptions), ServiceProtocol.JsonOptions)!;
        var roundTrip = ServiceResponse.FromShared(shared);

        Assert.False(roundTrip.Success);
        Assert.Equal("vba.protect-project", roundTrip.Command);
        Assert.Contains("Original restoration failure.", roundTrip.InnerError);
        if (wrapping == "aggregate") Assert.Contains("Independent cleanup failure.", roundTrip.InnerError);
        using var state = JsonDocument.Parse(roundTrip.PartialState!);
        var root = state.RootElement;
        Assert.Equal("restoring-visibility", root.GetProperty("phase").GetString());
        Assert.True(root.GetProperty("acceptanceRequested").GetBoolean());
        Assert.True(root.GetProperty("configurationApplied").GetBoolean());
        Assert.True(root.GetProperty("lockOptionReadback").GetBoolean());
        Assert.True(root.GetProperty("dialogCleanupConfirmed").GetBoolean());
        Assert.True(root.GetProperty("activeProjectRestored").GetBoolean());
        Assert.False(root.GetProperty("visibilityRestored").GetBoolean());
        Assert.True(root.GetProperty("saveRequired").GetBoolean());
        Assert.True(root.GetProperty("reopenRequired").GetBoolean());
        Assert.False(root.GetProperty("rollbackAttempted").GetBoolean());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(null)]
    public void UnconfirmedAcceptance_DoesNotBecomeAppliedOrPersisted(bool? applied)
    {
        using var service = new ExcelMcpService();
        const string secret = "SyntheticPasswordNeverReturned!7";
        var failure = new VbaProjectProtectionException("The native outcome is unconfirmed.",
            new TimeoutException("The expected dialog did not settle."),
            new VbaProjectProtectionState
            {
                Phase = "acceptance-requested",
                AcceptanceRequested = true,
                ConfigurationApplied = applied,
                DialogCleanupConfirmed = false,
                ActiveProjectRestored = false,
                VisibilityRestored = false
            });
        var response = CreateError(service, failure);
        response = ExcelMcpService.AttachRequestContext(new ServiceRequest
        {
            Command = "vba.protect-project",
            SessionId = "owned-session",
            Args = JsonSerializer.Serialize(new { password = secret })
        }, response);
        var serialized = JsonSerializer.Serialize(response, ServiceProtocol.JsonOptions);
        Assert.DoesNotContain(secret, serialized, StringComparison.Ordinal);
        Assert.Equal("owned-session", response.SessionId);
        Assert.Equal("Timeout", response.ErrorCategory);

        using var state = JsonDocument.Parse(response.PartialState!);
        var root = state.RootElement;
        if (applied.HasValue) Assert.False(root.GetProperty("configurationApplied").GetBoolean());
        else Assert.False(root.TryGetProperty("configurationApplied", out _));
        Assert.False(root.TryGetProperty("lockOptionReadback", out _));
        Assert.False(root.TryGetProperty("saveRequired", out _));
        Assert.False(root.TryGetProperty("reopenRequired", out _));
        Assert.False(root.TryGetProperty("password", out _));
        Assert.False(root.GetProperty("rollbackAttempted").GetBoolean());
    }

    [Theory]
    [InlineData("cancelled")]
    [InlineData("timeout")]
    public void InterruptedOperation_PreservesCauseClassificationAndPartialEntry(string cause)
    {
        using var service = new ExcelMcpService();
        Exception interruption = cause == "cancelled"
            ? new OperationCanceledException("Native interaction was cancelled.")
            : new TimeoutException("Native interaction deadline expired.");
        var failure = new VbaProjectProtectionException("Password entry did not complete.", interruption,
            new VbaProjectProtectionState
            {
                Phase = "entering-password",
                AcceptanceRequested = false,
                ConfigurationApplied = false,
                DialogCleanupConfirmed = true,
                ActiveProjectRestored = true,
                VisibilityRestored = true
            });
        var response = CreateError(service, failure);
        Assert.False(response.Success);
        Assert.Equal(cause == "cancelled" ? "Cancelled" : "Timeout", response.ErrorCategory);
        Assert.Equal(interruption.Message, response.InnerError);
        using var state = JsonDocument.Parse(response.PartialState!);
        Assert.Equal("entering-password", state.RootElement.GetProperty("phase").GetString());
        Assert.False(state.RootElement.GetProperty("acceptanceRequested").GetBoolean());
        Assert.False(state.RootElement.GetProperty("configurationApplied").GetBoolean());
        Assert.True(state.RootElement.GetProperty("dialogCleanupConfirmed").GetBoolean());
    }

    [Theory]
    [InlineData("direct")]
    [InlineData("wrapped")]
    [InlineData("aggregate")]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2201",
        Justification = "Synthetic COM disconnection exercises transport without starting Excel.")]
    public void FatalDisconnect_PreservesProtectionStateAndOriginalNativeCause(string wrapping)
    {
        var failure = new VbaProjectProtectionException("Configuration outcome is unknown.",
            new COMException("Original native disconnect.", unchecked((int)0x80010108)),
            new VbaProjectProtectionState
            {
                Phase = "acceptance-requested",
                AcceptanceRequested = true,
                ConfigurationApplied = null,
                DialogCleanupConfirmed = false,
                ActiveProjectRestored = false,
                VisibilityRestored = false
            });
        var create = typeof(ExcelMcpService).GetMethod("CreateExcelDisconnectedResponse",
            BindingFlags.Static | BindingFlags.NonPublic)!;
        var response = Assert.IsType<ServiceResponse>(create.Invoke(null,
            ["lost-session", Wrap(failure, wrapping), "Excel process died.", null]));
        Assert.False(response.Success);
        Assert.Equal("ExcelProcessDied", response.ErrorCategory);
        Assert.Equal("0x80010108", response.HResult);
        Assert.Equal("lost-session", response.SessionId);
        Assert.Contains("Original native disconnect.", response.InnerError);
        if (wrapping == "aggregate") Assert.Contains("Independent cleanup failure.", response.InnerError);
        using var state = JsonDocument.Parse(response.PartialState!);
        Assert.Equal("acceptance-requested", state.RootElement.GetProperty("phase").GetString());
        Assert.True(state.RootElement.GetProperty("acceptanceRequested").GetBoolean());
        Assert.False(state.RootElement.TryGetProperty("configurationApplied", out _));
        Assert.False(state.RootElement.GetProperty("dialogCleanupConfirmed").GetBoolean());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RetainedSession_CombinesProtectionAndExecutionState(bool waitExpired)
    {
        var protection = new VbaProjectProtectionState
        {
            Phase = "verifying-lock-option",
            AcceptanceRequested = true,
            ConfigurationApplied = null,
            SaveRequired = true,
            ReopenRequired = true
        };
        var operation = new ExcelOperationStatus("running", 1.5, waitExpired, true, null, "shown")
        {
            Command = "vba.protect-project"
        };
        var serialize = typeof(ExcelMcpService).GetMethod("SerializeVbaProtectionState",
            BindingFlags.Static | BindingFlags.NonPublic)!;
        var serialized = Assert.IsType<string>(serialize.Invoke(null,
            [protection, operation, "retained-session", RetainedFiles]));
        using var state = JsonDocument.Parse(serialized);
        var root = state.RootElement;
        Assert.Equal("verifying-lock-option", root.GetProperty("phase").GetString());
        Assert.True(root.GetProperty("acceptanceRequested").GetBoolean());
        Assert.False(root.TryGetProperty("configurationApplied", out _));
        Assert.Equal("retained-session", root.GetProperty("sessionId").GetString());
        Assert.True(root.GetProperty("retained").GetBoolean());
        Assert.Equal("synthetic.xlsm", root.GetProperty("filePaths")[0].GetString());
        var execution = root.GetProperty("operation");
        Assert.Equal("vba.protect-project", execution.GetProperty("command").GetString());
        Assert.True(execution.GetProperty("executionPending").GetBoolean());
        Assert.Equal(waitExpired, execution.GetProperty("waitExpired").GetBoolean());
    }

    [Fact]
    public void RecoveryRevealFailure_PreservesOriginalAndIndependentCleanupCauses()
    {
        using var service = new ExcelMcpService();
        var failure = new VbaProjectProtectionException("Protection configuration needs inspection.",
            new InvalidOperationException("Original protection failure."),
            new VbaProjectProtectionState
            {
                Phase = "restoring-visibility",
                AcceptanceRequested = true,
                ConfigurationApplied = true,
                SaveRequired = true,
                ReopenRequired = true
            });
        var create = typeof(ExcelMcpService).GetMethod("CreateErrorResponse",
            BindingFlags.Instance | BindingFlags.NonPublic)!;
        var response = Assert.IsType<ServiceResponse>(create.Invoke(service,
            [Wrap(failure, "aggregate"), "vba.protect-project", null, null,
                new InvalidOperationException("Recovery reveal failure.")]));
        Assert.False(response.Success);
        Assert.Contains("Original protection failure.", response.InnerError);
        Assert.Contains("Independent cleanup failure.", response.InnerError);
        Assert.Contains("Recovery reveal failure.", response.InnerError);
        using var state = JsonDocument.Parse(response.PartialState!);
        Assert.True(state.RootElement.GetProperty("configurationApplied").GetBoolean());
        Assert.Equal("restoring-visibility", state.RootElement.GetProperty("phase").GetString());
    }

    private static Exception Wrap(VbaProjectProtectionException failure, string wrapping) => wrapping switch
    {
        "wrapped" => new InvalidOperationException("Operation wrapper.", failure),
        "aggregate" => new AggregateException(new InvalidOperationException("Independent cleanup failure."), failure),
        _ => failure
    };

    private static ServiceResponse CreateError(ExcelMcpService service, Exception failure)
    {
        var create = typeof(ExcelMcpService).GetMethod("CreateErrorResponse",
            BindingFlags.Instance | BindingFlags.NonPublic)!;
        return Assert.IsType<ServiceResponse>(create.Invoke(service,
            [failure, "vba.protect-project", null, null, null]));
    }
}
