using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using Microsoft.Win32.SafeHandles;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Service.Workers;

internal sealed record WorkerProcess(int ProcessId, long StartedAtUtcFileTime)
{
    internal const int ErrorAccessDenied = 5;
    internal static Action<Process>? BeforeIdentityReadForTests { get; set; }

    internal static WorkerProcess Current
    {
        get
        {
            using var process = Process.GetCurrentProcess();
            return From(process);
        }
    }

    internal static WorkerProcess From(Process process) =>
        new(process.Id, process.StartTime.ToUniversalTime().ToFileTimeUtc());

    internal bool IsAlive()
    {
        Process process;
        try
        {
            process = Process.GetProcessById(ProcessId);
        }
        catch (ArgumentException)
        {
            return false;
        }
        using (process)
        {
            try
            {
                _ = process.Handle;
                if (process.HasExited)
                    return false;
                BeforeIdentityReadForTests?.Invoke(process);
                using var current = Process.GetCurrentProcess();
                return process.SessionId == current.SessionId
                    && process.StartTime.ToUniversalTime().ToFileTimeUtc() == StartedAtUtcFileTime;
            }
            catch (InvalidOperationException) when (process.HasExited)
            {
                return false;
            }
            catch (Win32Exception error) when (error.NativeErrorCode == ErrorAccessDenied
                && OwnedProcessGuard.TryConfirmExited(new ExcelProcessIdentity(ProcessId, StartedAtUtcFileTime)))
            {
                // The full process handle or identity read can race exit. Only
                // the exact process's limited-right probe can resolve that case.
                return false;
            }
        }
    }
}

internal sealed record WorkerRegistration(
    string PipeName,
    WorkerProcess Worker,
    WorkerProcess Controller,
    string Generation,
    string Build,
    ExcelProcessIdentity? Excel,
    bool OwnsExcel = false,
    bool ExcelPreserved = false,
    bool ReleaseConfirmed = false,
    ExcelProcessIdentity? ReleasedExcel = null,
    ExcelNativeProcessObservation? LastExcelObservation = null);

internal sealed record WorkerClaim(WorkerProcess Controller, string Generation, string Build);
internal sealed record WorkerCall(string OperationId, string Generation, ServiceRequest Request);
internal sealed record WorkerExecution(ServiceResponse Response, bool Accepted, WorkerSnapshot Snapshot,
    bool HasLateNativeCompletion = false);
internal sealed record WorkerRelease(ServiceResponse Response, WorkerSnapshot Snapshot);
internal sealed record WorkerStatusFailure(string Message, WorkerStatusDiagnostic? Diagnostic = null);
internal sealed record WorkerStatusDiagnostic(string Stage, string ExceptionType,
    [property: JsonPropertyName("hresult")] string? HResult)
{
    internal static WorkerStatusFailure Capture(string stage, string message, Exception error)
    {
        var cause = error;
        for (var inner = error; inner is not null; inner = inner.InnerException)
        {
            if (inner is COMException)
            {
                cause = inner;
                break;
            }
        }
        return new(message, new(stage, cause.GetType().Name, cause is COMException ? $"0x{cause.HResult:X8}" : null));
    }

    internal JsonObject ToJson(WorkerRegistration registration)
    {
        var diagnostic = new JsonObject
        {
            ["stage"] = Stage,
            ["exceptionType"] = ExceptionType,
            ["hresult"] = HResult
        };
        if (registration.LastExcelObservation is { } observation
            && (registration.Excel == observation.Identity || registration.ReleasedExcel == observation.Identity))
        {
            diagnostic["nativeProcess"] = new JsonObject
            {
                ["processId"] = observation.Identity.ProcessId,
                ["startedAtUtcFileTime"] = observation.Identity.StartedAtUtcFileTime.ToString(CultureInfo.InvariantCulture),
                ["state"] = observation.State.ToString(),
                ["exitCode"] = observation.ExitCode is { } exitCode ? JsonValue.Create(exitCode) : null,
                ["nativeError"] = observation.NativeError is { } nativeError ? JsonValue.Create(nativeError) : null
            };
        }
        return diagnostic;
    }
}

internal sealed record WorkerSnapshot(
    WorkerRegistration Registration,
    IReadOnlyList<SessionDescriptor> Sessions,
    ServiceResponse SessionList,
    string? PendingOperation,
    string? PendingCommand,
    IReadOnlyDictionary<string, IReadOnlyList<string>> OwnershipKeys,
    string? StatusError = null,
    bool Releasing = false,
    long Revision = 0,
    WorkerStatusDiagnostic? StatusDiagnostic = null);

[StreamJsonRpc.JsonRpcContract]
[PolyType.GenerateShape(IncludeMethods = PolyType.MethodShapeFlags.PublicInstance)]
internal partial interface IExcelWorkerRpc
{
    Task<WorkerSnapshot> ClaimAsync(WorkerClaim claim, CancellationToken cancellationToken = default);
    Task<WorkerSnapshot> SnapshotAsync(string generation, CancellationToken cancellationToken = default);
    Task<WorkerExecution> ExecuteAsync(WorkerCall call, CancellationToken cancellationToken = default);
    Task<ServiceResponse?> ObserveLateCompletionAsync(string generation, string operationId, CancellationToken cancellationToken = default);
    Task<WorkerRelease> ReleaseAsync(string generation, CancellationToken cancellationToken = default);
}

internal static class WorkerPipeIdentity
{
    internal static void VerifyServer(SafePipeHandle handle, WorkerProcess expected)
    {
        if (!GetNamedPipeServerProcessId(handle, out var processId))
            throw new Win32Exception(Marshal.GetLastWin32Error());
        if (processId != expected.ProcessId || !expected.IsAlive())
            throw new InvalidOperationException("The Excel worker pipe belongs to a different process.");
    }

    internal static int GetClient(SafePipeHandle handle)
    {
        if (!GetNamedPipeClientProcessId(handle, out var processId))
            throw new Win32Exception(Marshal.GetLastWin32Error());
        return checked((int)processId);
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetNamedPipeServerProcessId(SafePipeHandle pipe, out uint processId);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetNamedPipeClientProcessId(SafePipeHandle pipe, out uint processId);
}
