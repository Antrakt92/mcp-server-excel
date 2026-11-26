using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text.Json.Serialization;
using Microsoft.Win32.SafeHandles;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal enum ExcelProcessResourceStatus
{
    [JsonStringEnumMemberName("observed")]
    Observed,
    [JsonStringEnumMemberName("partial")]
    Partial,
    [JsonStringEnumMemberName("exited")]
    Exited,
    [JsonStringEnumMemberName("unavailable")]
    Unavailable
}

internal sealed record ExcelProcessResourceSnapshot(
    ExcelProcessResourceStatus Status,
    DateTimeOffset? ObservedAtUtc = null,
    uint? TotalHandles = null,
    uint? GdiObjects = null,
    uint? UserObjects = null,
    ulong? PrivateBytes = null)
{
    internal static ExcelProcessResourceSnapshot FromProcessState(ExcelNativeProcessState state) =>
        new(state == ExcelNativeProcessState.Exited ? ExcelProcessResourceStatus.Exited : ExcelProcessResourceStatus.Unavailable);
}

/// <summary>Samples Windows counters without opening Excel or obtaining termination or memory-read access.</summary>
internal static class ExcelProcessResourceProbe
{
    private const uint GdiObjects = 0;
    private const uint UserObjects = 1;
    private static readonly uint MemoryCountersSize = (uint)Marshal.SizeOf<ProcessMemoryCounters>();

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "An optional resource snapshot reports unavailable when Windows cannot grant or verify the read-only process handle; callers cannot use it as ownership evidence.")]
    internal static ExcelProcessResourceSnapshot Read(ExcelProcessIdentity identity)
    {
        try
        {
            using var observer = ExcelNativeProcessObserver.Capture(identity);
            return observer.ReadResources();
        }
        catch (Win32Exception)
        {
            return new(ExcelProcessResourceStatus.Unavailable);
        }
        catch (InvalidOperationException)
        {
            return new(ExcelProcessResourceStatus.Unavailable);
        }
    }

    internal static ExcelProcessResourceSnapshot ReadCounters(SafeProcessHandle handle)
    {
        var observedAtUtc = DateTimeOffset.UtcNow;
        uint? handles = GetProcessHandleCount(handle, out var count) ? count : null;
        var gdi = ReadGuiObjects(handle, GdiObjects);
        var user = ReadGuiObjects(handle, UserObjects);
        ulong? privateBytes = K32GetProcessMemoryInfo(handle, out var memory, MemoryCountersSize)
            ? memory.PrivateUsage : null;
        if (handles is null && gdi is null && user is null && privateBytes is null)
            return new(ExcelProcessResourceStatus.Unavailable);

        var status = handles is not null && gdi is not null && user is not null && privateBytes is not null
            ? ExcelProcessResourceStatus.Observed : ExcelProcessResourceStatus.Partial;
        return new(status, observedAtUtc, handles, gdi, user, privateBytes);
    }

    private static uint? ReadGuiObjects(SafeProcessHandle handle, uint kind)
    {
        var result = GetGuiResources(handle, kind);
        // Zero is valid for a process without GUI resources. SetLastError
        // clears the native error before invoking the function.
        return result == 0 && Marshal.GetLastPInvokeError() != 0 ? null : result;
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetProcessHandleCount(SafeProcessHandle handle, out uint count);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetGuiResources(SafeProcessHandle handle, uint kind);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool K32GetProcessMemoryInfo(SafeProcessHandle handle, out ProcessMemoryCounters counters, uint size);

    [StructLayout(LayoutKind.Sequential)]
    private struct ProcessMemoryCounters
    {
        internal uint Size;
        internal uint PageFaultCount;
        internal nuint PeakWorkingSetSize;
        internal nuint WorkingSetSize;
        internal nuint QuotaPeakPagedPoolUsage;
        internal nuint QuotaPagedPoolUsage;
        internal nuint QuotaPeakNonPagedPoolUsage;
        internal nuint QuotaNonPagedPoolUsage;
        internal nuint PagefileUsage;
        internal nuint PeakPagefileUsage;
        internal nuint PrivateUsage;
    }
}
