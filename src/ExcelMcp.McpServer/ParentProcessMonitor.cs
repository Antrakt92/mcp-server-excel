using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Hosting;
using Microsoft.Win32.SafeHandles;

namespace Sbroenne.ExcelMcp.McpServer;

/// <summary>Watches the owning client without touching its synchronous stdin pipe.</summary>
internal sealed class ParentProcessMonitor : IDisposable
{
    private const uint Synchronize = 0x00100000;
    private const uint QueryLimitedInformation = 0x1000;
    private const uint SnapshotProcesses = 2;
    private const uint WaitObject0 = 0;
    private readonly SafeProcessHandle _parent;
    private readonly IHostApplicationLifetime _lifetime;
    private readonly Timer _timer;
    private int _checking;
    private int _disposed;
    private int _signaled;

    private ParentProcessMonitor(IHostApplicationLifetime lifetime, SafeProcessHandle parent,
        TimeSpan interval)
    {
        _parent = parent;
        _lifetime = lifetime;
        _timer = new Timer(CheckParent, null, interval, interval);
    }

    internal static ParentProcessMonitor? Start(IHostApplicationLifetime lifetime)
    {
        // Interactive/file-redirection launches retain the SDK's normal lifecycle.
        if (GetFileType(GetStdHandle(-10)) != 3)
            return null;
        using var snapshot = CreateToolhelp32Snapshot(SnapshotProcesses, 0);
        if (snapshot.IsInvalid)
            return null;
        var entry = new ProcessEntry { Size = (uint)Marshal.SizeOf<ProcessEntry>() };
        if (!Process32First(snapshot, ref entry))
            return null;
        do
        {
            if (entry.ProcessId == (uint)Environment.ProcessId)
            {
                using var self = Process.GetCurrentProcess();
                return StartForProcess(lifetime, entry.ParentProcessId,
                    notCreatedAfter: self.StartTime.ToUniversalTime().ToFileTimeUtc());
            }
        } while (Process32Next(snapshot, ref entry));
        return null;
    }

    internal static ParentProcessMonitor? StartForProcess(IHostApplicationLifetime lifetime,
        uint processId, TimeSpan? interval = null, long? notCreatedAfter = null)
    {
        var parent = OpenProcess(Synchronize | QueryLimitedInformation, false, processId);
        if (parent.IsInvalid)
        {
            var error = Marshal.GetLastWin32Error();
            parent.Dispose();
            if (error == 87) // Parent has already exited.
                lifetime.StopApplication();
            return null;
        }
        // A PID can be reused between the process snapshot and OpenProcess.
        if (!GetProcessTimes(parent, out var created, out _, out _, out _))
        {
            parent.Dispose();
            return null;
        }
        if (notCreatedAfter.HasValue && created > notCreatedAfter.Value)
        {
            parent.Dispose();
            lifetime.StopApplication();
            return null;
        }
        return new ParentProcessMonitor(lifetime, parent, interval ?? TimeSpan.FromSeconds(1));
    }

    private void CheckParent(object? state)
    {
        if (Volatile.Read(ref _disposed) != 0 || Interlocked.Exchange(ref _checking, 1) != 0)
            return;
        try
        {
            // Zero-time process waits do not serialize with stdin reads or block workers.
            if (WaitForSingleObject(_parent, 0) == WaitObject0 &&
                Interlocked.Exchange(ref _signaled, 1) == 0)
                _lifetime.StopApplication();
        }
        catch (ObjectDisposedException) when (Volatile.Read(ref _disposed) != 0)
        {
            // Dispose raced the timer after its initial check.
            return;
        }
        finally { Volatile.Write(ref _checking, 0); }
    }

    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0)
            return;
        _timer.Dispose();
        _parent.Dispose();
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct ProcessEntry
    {
        internal uint Size, Usage, ProcessId;
        internal UIntPtr DefaultHeapId;
        internal uint ModuleId, Threads, ParentProcessId;
        internal int BasePriority;
        internal uint Flags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] internal string ExeFile;
    }
    [DllImport("kernel32.dll")] private static extern IntPtr GetStdHandle(int kind);
    [DllImport("kernel32.dll")] private static extern uint GetFileType(IntPtr handle);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern SafeFileHandle CreateToolhelp32Snapshot(uint flags, uint pid);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "Process32FirstW", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool Process32First(SafeFileHandle snapshot, ref ProcessEntry entry);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "Process32NextW", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool Process32Next(SafeFileHandle snapshot, ref ProcessEntry entry);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern SafeProcessHandle OpenProcess(uint access, [MarshalAs(UnmanagedType.Bool)] bool inherit, uint pid);
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool GetProcessTimes(SafeProcessHandle process, out long created, out long exited, out long kernel, out long user);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern uint WaitForSingleObject(SafeProcessHandle handle, uint milliseconds);
}
