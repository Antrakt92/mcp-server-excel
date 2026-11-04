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
public sealed class VbaProjectAuthenticationFailureTransportTests
{
    private static readonly string[] RetainedFiles = ["synthetic.xlsm"];

    [Theory]
    [InlineData(null)]
    [InlineData(false)]
    [InlineData(true)]
    public void ObservedOutcome_RoundTripsWithoutInventingVerificationOrEchoingSecret(bool? verified)
    {
        using var service = new ExcelMcpService();
        var error = new VbaProjectAuthenticationException("Authentication needs inspection.",
            new TimeoutException("The final restoration did not settle."), new VbaProjectAuthenticationState
            {
                Phase = "restoring-visibility",
                ChallengeObserved = true,
                AuthenticationPerformed = true,
                PasswordVerified = verified,
                AccessUnlocked = verified.HasValue ? verified.Value : null,
                CurrentProtection = verified is true ? "None" : "Locked",
                DialogCleanupConfirmed = true
            });
        var response = CreateError(service, error);
        const string secret = "Synthetic\"NeverEcho\\!7";
        response = ExcelMcpService.AttachRequestContext(new ServiceRequest
        {
            Command = "vba.unlock-project",
            SessionId = "owned-session",
            Args = JsonSerializer.Serialize(new { password = secret })
        }, response);
        var json = JsonSerializer.Serialize(response, ServiceProtocol.JsonOptions);
        Assert.DoesNotContain(secret, json, StringComparison.Ordinal);
        Assert.DoesNotContain(JsonEncodedText.Encode(secret).ToString(), json, StringComparison.Ordinal);
        var roundTrip = ServiceResponse.FromShared(JsonSerializer.Deserialize<Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceResponse>(
            json, ServiceProtocol.JsonOptions)!);
        Assert.False(roundTrip.Success);
        Assert.Equal("Timeout", roundTrip.ErrorCategory);
        Assert.Equal("owned-session", roundTrip.SessionId);
        using var state = JsonDocument.Parse(roundTrip.PartialState!);
        var actual = state.RootElement.GetProperty("passwordVerified");
        if (verified.HasValue) Assert.Equal(verified.Value, actual.GetBoolean());
        else Assert.Equal(JsonValueKind.Null, actual.ValueKind);
        Assert.False(state.RootElement.GetProperty("saved").GetBoolean());
        Assert.False(state.RootElement.GetProperty("reopened").GetBoolean());
        Assert.False(state.RootElement.GetProperty("protectionSettingsChanged").GetBoolean());
        Assert.False(state.RootElement.GetProperty("rollbackAttempted").GetBoolean());
    }

    [Theory]
    [InlineData("direct")]
    [InlineData("wrapped")]
    [InlineData("aggregate")]
    public void NestedAuthenticationFailure_PreservesNativeAndCleanupCauses(string wrapping)
    {
        using var service = new ExcelMcpService();
        var failure = new VbaProjectAuthenticationException("Native outcome is unknown.",
            new InvalidOperationException("Original authentication cause."), new VbaProjectAuthenticationState
            {
                Phase = "requesting-authentication",
                AuthenticationPerformed = true,
                AccessUnlocked = null
            });
        var response = CreateError(service, Wrap(failure, wrapping), new InvalidOperationException("Recovery reveal cause."));
        Assert.Contains("Original authentication cause.", response.InnerError);
        Assert.Contains("Recovery reveal cause.", response.InnerError);
        if (wrapping == "aggregate") Assert.Contains("Independent dialog cleanup cause.", response.InnerError);
        using var state = JsonDocument.Parse(response.PartialState!);
        Assert.Equal(JsonValueKind.Null, state.RootElement.GetProperty("passwordVerified").ValueKind);
        Assert.Equal(JsonValueKind.Null, state.RootElement.GetProperty("accessUnlocked").ValueKind);
        Assert.False(state.RootElement.GetProperty("dialogCleanupConfirmed").GetBoolean());
    }

    [Theory]
    [InlineData("direct")]
    [InlineData("wrapped")]
    [InlineData("aggregate")]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2201", Justification = "Synthetic fatal COM disconnect exercises failure transport without starting Excel.")]
    public void FatalDisconnect_RetainsUnknownAuthenticationState(string wrapping)
    {
        var failure = new VbaProjectAuthenticationException("Authentication completion is unknown.",
            new COMException("Original authentication disconnect.", unchecked((int)0x80010108)),
            new VbaProjectAuthenticationState { Phase = "requesting-authentication", AuthenticationPerformed = true, AccessUnlocked = null });
        var method = typeof(ExcelMcpService).GetMethod("CreateExcelDisconnectedResponse", BindingFlags.Static | BindingFlags.NonPublic)!;
        var response = Assert.IsType<ServiceResponse>(method.Invoke(null,
            ["lost-session", Wrap(failure, wrapping), "Excel process died."]));
        Assert.Equal("ExcelProcessDied", response.ErrorCategory);
        Assert.Equal("0x80010108", response.HResult);
        Assert.Contains("Original authentication disconnect.", response.InnerError);
        using var state = JsonDocument.Parse(response.PartialState!);
        Assert.Equal(JsonValueKind.Null, state.RootElement.GetProperty("passwordVerified").ValueKind);
        Assert.Equal("requesting-authentication", state.RootElement.GetProperty("phase").GetString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RetainedOperation_CombinesExecutionAndAuthenticationWithoutCachingPassword(bool waitExpired)
    {
        var state = new VbaProjectAuthenticationState { Phase = "writing-password", ChallengeObserved = true };
        var operation = new ExcelOperationStatus("running", 1.2, waitExpired, true, null, "shown")
        { Command = "vba.verify-password" };
        var method = typeof(ExcelMcpService).GetMethod("SerializeVbaAuthenticationState", BindingFlags.Static | BindingFlags.NonPublic)!;
        var json = Assert.IsType<string>(method.Invoke(null, [state, operation, "retained-session", RetainedFiles]));
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        Assert.Equal("retained-session", root.GetProperty("sessionId").GetString());
        Assert.True(root.GetProperty("retained").GetBoolean());
        Assert.Equal("synthetic.xlsm", root.GetProperty("filePaths")[0].GetString());
        Assert.Equal(waitExpired, root.GetProperty("operation").GetProperty("waitExpired").GetBoolean());
        Assert.True(root.GetProperty("operation").GetProperty("executionPending").GetBoolean());
        Assert.Equal(JsonValueKind.Null, root.GetProperty("passwordVerified").ValueKind);
        Assert.False(root.TryGetProperty("password", out _));
    }

    [Theory]
    [InlineData(OperationFailureCategory.Permissions)]
    [InlineData(OperationFailureCategory.Prerequisite)]
    public void ExpectedRefusal_PreservesCategoryAndUnverifiedState(OperationFailureCategory category)
    {
        using var service = new ExcelMcpService();
        var failure = new VbaProjectAuthenticationException("Authentication was refused.",
            new OperationFailureException(category, "Known refusal."),
            new VbaProjectAuthenticationState { Phase = "already-accessible", AlreadyAccessible = true, CurrentProtection = "None" });
        var response = CreateError(service, failure);
        Assert.Equal(category.ToString(), response.ErrorCategory);
        using var state = JsonDocument.Parse(response.PartialState!);
        Assert.True(state.RootElement.GetProperty("alreadyAccessible").GetBoolean());
        Assert.False(state.RootElement.GetProperty("authenticationPerformed").GetBoolean());
        Assert.Equal(JsonValueKind.Null, state.RootElement.GetProperty("passwordVerified").ValueKind);
    }

    private static Exception Wrap(Exception error, string wrapping) => wrapping switch
    {
        "wrapped" => new InvalidOperationException("Operation wrapper.", error),
        "aggregate" => new AggregateException(new InvalidOperationException("Independent dialog cleanup cause."), error),
        _ => error
    };

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InterruptedAuthentication_RetainsCancellationOrTimeoutClassification(bool cancelled)
    {
        using var service = new ExcelMcpService();
        Exception cause = cancelled ? new OperationCanceledException("Native authentication was cancelled.")
            : new TimeoutException("Native authentication deadline expired.");
        var response = CreateError(service, new VbaProjectAuthenticationException("Authentication outcome is unknown.", cause,
            new VbaProjectAuthenticationState { Phase = "requesting-authentication", AuthenticationPerformed = true, AccessUnlocked = null }));
        Assert.Equal(cancelled ? "Cancelled" : "Timeout", response.ErrorCategory);
        using var state = JsonDocument.Parse(response.PartialState!);
        Assert.Equal(JsonValueKind.Null, state.RootElement.GetProperty("passwordVerified").ValueKind);
        Assert.Equal(JsonValueKind.Null, state.RootElement.GetProperty("accessUnlocked").ValueKind);
        Assert.False(state.RootElement.GetProperty("rollbackAttempted").GetBoolean());
    }

    private static ServiceResponse CreateError(ExcelMcpService service, Exception error, Exception? recovery = null)
    {
        var method = typeof(ExcelMcpService).GetMethod("CreateErrorResponse", BindingFlags.Instance | BindingFlags.NonPublic)!;
        return Assert.IsType<ServiceResponse>(method.Invoke(service, [error, "vba.unlock-project", null, null, recovery]));
    }
}
