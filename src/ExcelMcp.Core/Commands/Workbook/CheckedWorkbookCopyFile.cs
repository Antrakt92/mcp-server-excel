using System.Buffers.Binary;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>Publishes the exact checked staging file while excluding competing writes and renames.</summary>
internal static class CheckedWorkbookCopyFile
{
    private const uint GenericRead = 0x80000000;
    private const uint DeleteAccess = 0x00010000;
    private const uint OpenExisting = 3;
    private const uint OpenReparsePoint = 0x00200000;
    private const int FileRenameInfo = 3;
    private const int FileDispositionInfo = 4;
    private const int FileAttributeTagInfo = 9;
    private const int ErrorFileExists = 80;
    private const int ErrorAlreadyExists = 183;

    internal static FileStream Open(string path)
    {
        // Keep DELETE access on our handle while denying write/delete access to other handles.
        var handle = CreateFile(path, GenericRead | DeleteAccess, FileShare.Read, IntPtr.Zero,
            OpenExisting, OpenReparsePoint, IntPtr.Zero);
        try
        {
            if (handle.IsInvalid)
                throw LastIoError($"Could not lock the checked copy for comparison. The destination was not changed. The unpublished stage may remain at '{path}'.");
            if (!GetFileInformationByHandleEx(handle, FileAttributeTagInfo, out var info, (uint)Marshal.SizeOf<AttributeTagInformation>()))
                throw LastIoError($"Could not verify the checked copy file. The destination was not changed. The unpublished stage may remain at '{path}'.");
            if ((info.Attributes & (FileAttributes.Directory | FileAttributes.ReparsePoint)) != 0)
                throw new IOException($"The checked copy stage is not a regular file: '{path}'. The destination was not changed and the stage was left for inspection.");
            return new FileStream(handle, FileAccess.Read);
        }
        catch
        {
            handle.Dispose();
            throw;
        }
    }

    internal static void Publish(FileStream staged, string destination, bool overwrite)
    {
        int nameOffset = (int)Marshal.OffsetOf<RenameInformation>(nameof(RenameInformation.FirstFileNameCharacter));
        int lengthOffset = (int)Marshal.OffsetOf<RenameInformation>(nameof(RenameInformation.FileNameLength));
        byte[] name = Encoding.Unicode.GetBytes(destination);
        // FileNameLength excludes the terminating wide null. RootDirectory is null for an absolute path.
        var info = new byte[Math.Max(Marshal.SizeOf<RenameInformation>(), checked(nameOffset + name.Length + sizeof(char)))];
        info[0] = overwrite ? (byte)1 : (byte)0;
        BinaryPrimitives.WriteUInt32LittleEndian(info.AsSpan(lengthOffset), (uint)name.Length);
        name.CopyTo(info.AsSpan(nameOffset));
        if (SetFileInformationByHandle(staged.SafeFileHandle, FileRenameInfo, info, (uint)info.Length))
            return;

        var error = new Win32Exception(Marshal.GetLastWin32Error());
        if (!overwrite && (error.NativeErrorCode is ErrorFileExists or ErrorAlreadyExists || File.Exists(destination)))
            throw new IOException($"Output file already exists: '{destination}'.", error);
        throw new IOException($"Could not publish the checked copy to '{destination}'.", error);
    }

    internal static void MarkForDeletion(FileStream staged, string path)
    {
        byte[] info = [1];
        if (!SetFileInformationByHandle(staged.SafeFileHandle, FileDispositionInfo, info, (uint)info.Length))
            throw LastIoError($"Could not remove the unpublished checked copy stage at '{path}'.");
    }

    private static IOException LastIoError(string message) => new(message, new Win32Exception(Marshal.GetLastWin32Error()));

    [StructLayout(LayoutKind.Sequential)]
    private struct RenameInformation
    {
        internal uint ReplaceIfExistsOrFlags;
        internal IntPtr RootDirectory;
        internal uint FileNameLength;
        internal ushort FirstFileNameCharacter;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct AttributeTagInformation
    {
        internal FileAttributes Attributes;
        internal uint ReparseTag;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateFile(
        string path, uint desiredAccess, FileShare shareMode, IntPtr securityAttributes,
        uint creationDisposition, uint flagsAndAttributes, IntPtr template);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetFileInformationByHandleEx(
        SafeFileHandle handle, int fileInformationClass, out AttributeTagInformation information, uint size);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetFileInformationByHandle(
        SafeFileHandle handle, int fileInformationClass, byte[] information, uint size);
}
