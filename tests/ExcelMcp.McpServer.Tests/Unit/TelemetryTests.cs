// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.ExcelMcp.McpServer.Telemetry;
using Sbroenne.ExcelMcp.Telemetry;
using System.Reflection;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

/// <summary>
/// Tests for telemetry configuration and sensitive data redaction.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Telemetry")]
[Collection("ProgramTransport")]
[Trait("RequiresExcel", "false")]
public class TelemetryTests
{
    #region ExcelMcpTelemetry Tests

    [Fact]
    public void CliAndMcpAssemblies_ReadTheSamePersistedUserIdentity()
    {
        using var directory = new TemporaryDirectory();
        var cliAssembly = Assembly.Load("excelcli");
        var mcpAssembly = typeof(ExcelMcpTelemetry).Assembly;

        var cliIdentity = ReadIdentityFromAssembly(cliAssembly, directory.Path);
        var mcpIdentity = ReadIdentityFromAssembly(mcpAssembly, directory.Path);

        Assert.True(Guid.TryParseExact(cliIdentity.UserId, "N", out _));
        Assert.Equal(cliIdentity.UserId, mcpIdentity.UserId);
        Assert.NotNull(cliIdentity.LegacyUserId);
        Assert.Null(mcpIdentity.LegacyUserId);
    }

    [Fact]
    public void GetConnectionString_ReturnsNullForPlaceholder()
    {
        // The placeholder should not be treated as a valid connection string
        // (In dev builds, it's "__APPINSIGHTS_CONNECTION_STRING__")
        var connectionString = ExcelMcpTelemetry.GetConnectionString();

        // Either null (placeholder) or a real connection string (CI build)
        // We can't assert null directly because CI might inject a real one
        if (connectionString != null)
        {
            Assert.DoesNotContain("__", connectionString);
        }
    }

    [Fact]
    public void TrackUnhandledException_SendsOnlySanitizedClassification()
    {
        InvalidOperationException inner;
        try
        {
            throw new InvalidOperationException(
                @"Failed to open C:\Users\customer\secret.xlsx for customer@example.com");
        }
        catch (InvalidOperationException ex)
        {
            inner = ex;
        }

        var exception = new AggregateException("******", inner);
        var telemetry = ExcelMcpTelemetry.CreateSanitizedExceptionTelemetry(
            exception,
            "TaskScheduler.UnobservedTaskException");

        Assert.Equal("true", telemetry.Properties["Sanitized"]);
        Assert.Equal("AggregateException", telemetry.Properties["ExceptionType"]);
        Assert.Equal("InvalidOperationException", telemetry.Properties["InnerExceptionTypes"]);
        Assert.Equal("TaskScheduler.UnobservedTaskException", telemetry.Properties["Source"]);
        var serialized = string.Join("\n", telemetry.Properties.Values);
        Assert.DoesNotContain(@"C:\Users", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("customer@example.com", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("secret.xlsx", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("password", serialized, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TrackUnhandledException_RejectsUnapprovedSourceText()
    {
        UsageEvent telemetry = ExcelMcpTelemetry.CreateSanitizedExceptionTelemetry(
            new InvalidOperationException("sensitive"),
            @"C:\Users\customer\source");

        Assert.Equal("Unknown", telemetry.Properties["Source"]);
        Assert.DoesNotContain(
            @"C:\Users",
            string.Join("\n", telemetry.Properties.Values),
            StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    private static (string UserId, string? LegacyUserId) ReadIdentityFromAssembly(
        Assembly assembly,
        string localAppDataPath)
    {
        var store = assembly.GetType("Sbroenne.ExcelMcp.Telemetry.TelemetryUserIdentityStore", throwOnError: true)!;
        var identity = store.GetMethod(
                "GetOrCreate",
                BindingFlags.Static | BindingFlags.NonPublic,
                binder: null,
                types: [typeof(string)],
                modifiers: null)!
            .Invoke(null, [localAppDataPath])!;
        var type = identity.GetType();
        return (
            (string)type.GetProperty("UserId")!.GetValue(identity)!,
            (string?)type.GetProperty("LegacyUserId")!.GetValue(identity));
    }

    private sealed class TemporaryDirectory : IDisposable
    {
        public string Path { get; } = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(),
            $"excelmcp-telemetry-entrypoints-{Guid.NewGuid():N}");

        public void Dispose()
        {
            if (Directory.Exists(Path))
            {
                Directory.Delete(Path, recursive: true);
            }
        }
    }
}
