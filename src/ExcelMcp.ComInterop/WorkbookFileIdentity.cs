using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>Reads only Windows path/volume/file metadata; never reads workbook contents.</summary>
internal sealed record WorkbookFileIdentity(string Path, string? FileKey)
{
    internal static WorkbookFileIdentity Read(string fullPath)
    {
        var ancestor = fullPath;
        var missing = new Stack<string>();
        while (true)
        {
            // Zero desired access permits metadata inspection of a workbook open in Excel.
            using var handle = CreateFile(ancestor, 0, 7, IntPtr.Zero, 3, 0x02000000, IntPtr.Zero);
            if (!handle.IsInvalid)
            {
                if (!GetFileInformationByHandle(handle, out var info))
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                var resolved = ReadFinalPath(handle, 0);
                var physicalPath = resolved.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase)
                    ? @"\\" + resolved[8..]
                    : resolved.StartsWith(@"\\?\", StringComparison.Ordinal) ? resolved[4..] : resolved;
                if (!System.IO.Path.IsPathFullyQualified(physicalPath))
                    throw new IOException("Windows returned no absolute physical workbook path.");
                physicalPath = System.IO.Path.GetFullPath(physicalPath);
                if (missing.Count > 0)
                {
                    if ((info.Attributes & 0x10) == 0)
                        throw new IOException("A workbook destination ancestor is not a directory.");
                    while (missing.TryPop(out var leaf)) physicalPath = System.IO.Path.Combine(physicalPath, leaf);
                    return new(physicalPath, null);
                }

                // A volume GUID disambiguates local volumes even if serial numbers collide.
                // Remote handles expose a canonical UNC share instead of a local volume GUID.
                string volume;
                if (physicalPath.StartsWith(@"\\", StringComparison.Ordinal))
                    volume = System.IO.Path.GetPathRoot(physicalPath)!;
                else
                {
                    var volumePath = ReadFinalPath(handle, 1);
                    var end = volumePath.IndexOf('}');
                    if (end < 0) throw new IOException("Windows returned no stable workbook volume identity.");
                    volume = volumePath[..(end + 1)];
                }
                if (!GetFileInformationByHandleEx(handle, 18, out var fileId, (uint)Marshal.SizeOf<FileIdInformation>()))
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                if (fileId.Low == 0 && fileId.High == 0)
                    throw new IOException("Windows returned no stable workbook file identity.");
                return new(physicalPath, $"{volume}:{fileId.Volume:X16}:{fileId.High:X16}{fileId.Low:X16}");
            }

            var error = Marshal.GetLastWin32Error();
            if (error is not (2 or 3)) throw new Win32Exception(error);
            // A dangling junction/symlink is not an ordinary missing destination.
            var attributes = GetFileAttributes(ancestor);
            if (attributes != uint.MaxValue && (attributes & 0x400) != 0)
                throw new IOException("A workbook path has an unresolved reparse point.", new Win32Exception(error));
            if (attributes == uint.MaxValue && Marshal.GetLastWin32Error() is not (2 or 3))
                throw new Win32Exception(Marshal.GetLastWin32Error());
            var parent = System.IO.Path.GetDirectoryName(ancestor);
            var name = System.IO.Path.GetFileName(ancestor);
            if (string.IsNullOrEmpty(parent) || string.IsNullOrEmpty(name)) throw new Win32Exception(error);
            missing.Push(name);
            ancestor = parent;
        }
    }

    private static string ReadFinalPath(SafeFileHandle handle, uint flags)
    {
        var buffer = new char[32768];
        var length = GetFinalPathNameByHandle(handle, buffer, (uint)buffer.Length, flags);
        if (length == 0) throw new Win32Exception(Marshal.GetLastWin32Error());
        if (length >= buffer.Length) throw new IOException("The resolved workbook path exceeds the Windows path limit.");
        return new string(buffer, 0, (int)length);
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Information
    {
        internal uint Attributes;
        internal System.Runtime.InteropServices.ComTypes.FILETIME Creation, Access, Write;
        internal uint Volume, SizeHigh, SizeLow, Links, IndexHigh, IndexLow;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct FileIdInformation
    {
        internal ulong Volume, Low, High;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateFile(string path, uint access, uint share, IntPtr security, uint mode, uint flags, IntPtr template);
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetFileInformationByHandle(SafeFileHandle handle, out Information information);
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetFileInformationByHandleEx(SafeFileHandle handle, int informationClass, out FileIdInformation information, uint size);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint GetFinalPathNameByHandle(SafeFileHandle handle, [Out] char[] path, uint size, uint flags);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint GetFileAttributes(string path);
}
