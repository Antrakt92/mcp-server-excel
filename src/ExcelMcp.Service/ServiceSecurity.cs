using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using Microsoft.Win32.SafeHandles;
using ClientPipeSecurity = Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceSecurity;

namespace Sbroenne.ExcelMcp.Service;

/// <summary>
/// Security utilities for ExcelMCP Service named pipe communication.
/// Isolates service connections by user SID and Windows session.
/// </summary>
/// <remarks>
/// <para><b>Pipe Name Strategy:</b></para>
/// <list type="bullet">
///   <item>MCP Server: excelmcp-mcp-{SID}-{PID} (per-process isolation, each instance independent)</item>
///   <item>CLI daemon: excelmcp-cli-{SID}-session-{SESSION} (shared within the user's Windows session)</item>
/// </list>
/// <para><b>Security Model:</b></para>
/// <list type="bullet">
///   <item>User Isolation: Pipe name includes user SID - users cannot access each other's service instances</item>
///   <item>Windows ACLs: Named pipe restricts access to current user's SID via PipeSecurity</item>
///   <item>Windows session: Service peers must prove they belong to the current Windows session</item>
/// </list>
/// </remarks>
public static class ServiceSecurity
{
    internal static uint CurrentWindowsSessionId => ClientPipeSecurity.CurrentWindowsSessionId;

    /// <summary>
    /// Gets the pipe name for the MCP Server (per-process isolation).
    /// </summary>
    public static string GetMcpPipeName() => ClientPipeSecurity.GetMcpPipeName();

    /// <summary>
    /// Gets the default CLI daemon pipe for the current user and Windows session.
    /// </summary>
    public static string GetCliPipeName() => ClientPipeSecurity.GetCliPipeName();

    internal static uint? ReadClientSessionId(SafePipeHandle pipe) =>
        GetNamedPipeClientSessionId(pipe, out var sessionId) ? sessionId : null;

    internal static bool IsServerInCurrentSession(SafePipeHandle pipe) =>
        GetNamedPipeServerSessionId(pipe, out var sessionId) && sessionId == CurrentWindowsSessionId;

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetNamedPipeClientSessionId(SafePipeHandle pipe, out uint sessionId);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetNamedPipeServerSessionId(SafePipeHandle pipe, out uint sessionId);

    /// <summary>
    /// Creates a secure named pipe server with ACLs restricting access to current user only.
    /// </summary>
    public static NamedPipeServerStream CreateSecureServer(string pipeName)
    {
        var pipeSecurity = new PipeSecurity();
        using var identity = WindowsIdentity.GetCurrent();

        pipeSecurity.AddAccessRule(new PipeAccessRule(
            identity.User!,
            PipeAccessRights.FullControl,
            AccessControlType.Allow));

        return NamedPipeServerStreamAcl.Create(
            pipeName,
            PipeDirection.InOut,
            maxNumberOfServerInstances: NamedPipeServerStream.MaxAllowedServerInstances,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous,
            inBufferSize: 4096,
            outBufferSize: 4096,
            pipeSecurity);
    }

    /// <summary>
    /// Creates a client connection to a service pipe.
    /// </summary>
    public static NamedPipeClientStream CreateClient(string pipeName) => ClientPipeSecurity.CreateClient(pipeName);
}
