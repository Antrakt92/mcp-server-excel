using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal enum ExcelNativeProcessState
{
    Alive,
    Exited,
    Indeterminate
}

internal sealed record ExcelNativeProcessObservation(
    ExcelProcessIdentity Identity,
    ExcelNativeProcessState State,
    uint? ExitCode,
    int? NativeError);

/// <summary>
/// Observes one process through its retained, identity-checked handle without
/// acquiring termination rights or following a reused process number.
/// </summary>
internal sealed class ExcelNativeProcessObserver : IDisposable
{
    private const uint Synchronize = 0x00100000;
    private const uint ProcessQueryLimitedInformation = 0x1000;
    private const uint WaitObject0 = 0;
    private const uint WaitTimeout = 258;
    private const uint WaitFailed = uint.MaxValue;
    private const int ErrorInvalidHandle = 6;
    private readonly object _sync = new();
    private readonly SafeProcessHandle _handle;
    private readonly IProcessApi _api;
    private ExcelNativeProcessObservation? _lastObservation;
    private bool _disposed;

    private ExcelNativeProcessObserver(ExcelProcessIdentity identity, SafeProcessHandle handle, IProcessApi api)
    {
        Identity = identity;
        _handle = handle;
        _api = api;
    }

    internal ExcelProcessIdentity Identity { get; }

    internal static ExcelNativeProcessObserver Capture(int processId) =>
        Capture(processId, null, WindowsProcessApi.Instance);

    internal static ExcelNativeProcessObserver Capture(ExcelProcessIdentity expected) =>
        Capture(expected.ProcessId, expected.StartedAtUtcFileTime, WindowsProcessApi.Instance);

    internal static ExcelNativeProcessObserver Capture(int processId, long? expectedCreationTime, IProcessApi api)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(processId);
        ArgumentNullException.ThrowIfNull(api);
        var handle = api.Open(processId, ProcessQueryLimitedInformation | Synchronize, out var error);
        try
        {
            if (handle.IsInvalid)
                throw new Win32Exception(error, "Could not retain the process observation handle.");
            if (!api.GetCreationTime(handle, out var creationTime, out error))
                throw new Win32Exception(error, "Could not confirm the process creation time.");
            if (expectedCreationTime is { } expected && expected != creationTime)
                throw new InvalidOperationException("The process creation time does not match the expected identity.");

            return new ExcelNativeProcessObserver(new ExcelProcessIdentity(processId, creationTime), handle, api);
        }
        catch
        {
            handle.Dispose();
            throw;
        }
    }

    internal ExcelNativeProcessObservation Observe()
    {
        lock (_sync)
        {
            return ObserveCore();
        }
    }

    private ExcelNativeProcessObservation ObserveCore()
    {
        if (_disposed || _lastObservation is { State: ExcelNativeProcessState.Exited, ExitCode: not null })
            return _lastObservation!;
        if (_lastObservation is { State: ExcelNativeProcessState.Exited })
            return ObserveExitCode();

        var waitResult = _api.Wait(_handle, out var error);
        if (waitResult == WaitTimeout)
            return _lastObservation = new(Identity, ExcelNativeProcessState.Alive, null, null);
        if (waitResult != WaitObject0)
            return _lastObservation = new(Identity, ExcelNativeProcessState.Indeterminate, null,
                waitResult == WaitFailed ? error : null);

        // WHY: A signaled handle proves exit independently of the exit-code query.
        // A failed BOOL must not turn the native out parameter into a successful zero.
        return ObserveExitCode();
    }

    private ExcelNativeProcessObservation ObserveExitCode()
    {
        var queried = _api.GetExitCode(_handle, out var exitCode, out var error);
        return _lastObservation = new(Identity, ExcelNativeProcessState.Exited,
            queried ? exitCode : null, queried ? null : error);
    }

    public void Dispose()
    {
        lock (_sync)
        {
            if (_disposed)
                return;
            try
            {
                ObserveCore();
            }
            finally
            {
                _handle.Dispose();
                _disposed = true;
                // WARNING: Once observation stops, a formerly live process is unknown.
                if (_lastObservation is not { State: ExcelNativeProcessState.Exited })
                    _lastObservation = new(Identity, ExcelNativeProcessState.Indeterminate, null, ErrorInvalidHandle);
                GC.SuppressFinalize(this);
            }
        }
    }

    internal interface IProcessApi
    {
        SafeProcessHandle Open(int processId, uint desiredAccess, out int error);
        bool GetCreationTime(SafeProcessHandle handle, out long creationTime, out int error);
        uint Wait(SafeProcessHandle handle, out int error);
        bool GetExitCode(SafeProcessHandle handle, out uint exitCode, out int error);
    }

    private sealed class WindowsProcessApi : IProcessApi
    {
        internal static readonly WindowsProcessApi Instance = new();

        public SafeProcessHandle Open(int processId, uint desiredAccess, out int error)
        {
            var handle = OpenProcess(desiredAccess, inheritHandle: false, processId);
            error = handle.IsInvalid ? Marshal.GetLastWin32Error() : 0;
            return handle;
        }

        public bool GetCreationTime(SafeProcessHandle handle, out long creationTime, out int error)
        {
            var result = GetProcessTimes(handle, out creationTime, out _, out _, out _);
            error = result ? 0 : Marshal.GetLastWin32Error();
            return result;
        }

        public uint Wait(SafeProcessHandle handle, out int error)
        {
            var result = WaitForSingleObject(handle, 0);
            error = result == WaitFailed ? Marshal.GetLastWin32Error() : 0;
            return result;
        }

        public bool GetExitCode(SafeProcessHandle handle, out uint exitCode, out int error)
        {
            var result = GetExitCodeProcess(handle, out exitCode);
            error = result ? 0 : Marshal.GetLastWin32Error();
            return result;
        }
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern SafeProcessHandle OpenProcess(
        uint desiredAccess, [MarshalAs(UnmanagedType.Bool)] bool inheritHandle, int processId);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetProcessTimes(
        SafeProcessHandle processHandle, out long creationTime, out long exitTime, out long kernelTime, out long userTime);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint WaitForSingleObject(SafeProcessHandle handle, uint milliseconds);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetExitCodeProcess(SafeProcessHandle handle, out uint exitCode);
}
