using System.Diagnostics;
using System.IO.Pipes;
using System.Security.Principal;

namespace Sbroenne.ExcelMcp.ComInterop.ServiceClient;

/// <summary>
/// Security utilities for ExcelMCP Service named pipe communication (client-side).
/// Provides pipe name generation and client connection creation.
/// </summary>
public static class ServiceSecurity
{
    private static readonly Lazy<string> UserSid = new(() =>
    {
        using var identity = WindowsIdentity.GetCurrent();
        return identity.User?.Value ?? throw new InvalidOperationException(
            "Cannot determine current user SID. Named pipe security requires a valid SID for user isolation.");
    });

    private static readonly Lazy<uint> SessionId = new(() =>
    {
        using var process = Process.GetCurrentProcess();
        return checked((uint)process.SessionId);
    });

    internal static uint CurrentWindowsSessionId => SessionId.Value;

    /// <summary>
    /// Gets the pipe name for the MCP Server (per-process isolation).
    /// </summary>
    public static string GetMcpPipeName() => $"excelmcp-mcp-{UserSid.Value}-{Environment.ProcessId}";

    /// <summary>
    /// Gets the default CLI daemon pipe for the current user and Windows session.
    /// </summary>
    public static string GetCliPipeName() => $"excelmcp-cli-{UserSid.Value}-session-{CurrentWindowsSessionId}";

    /// <summary>
    /// Creates a client connection to a service pipe.
    /// </summary>
    public static NamedPipeClientStream CreateClient(string pipeName)
    {
        return new NamedPipeClientStream(
            ".",
            pipeName,
            PipeDirection.InOut,
            PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
    }
}
