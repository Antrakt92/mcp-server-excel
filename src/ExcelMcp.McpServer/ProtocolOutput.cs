using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace Sbroenne.ExcelMcp.McpServer;

internal static class ProtocolOutput
{
    internal static Stream Open()
    {
        if (!Console.IsOutputRedirected) return Console.OpenStandardOutput();
        // Windows ConsoleStream suppresses broken-pipe errors; protocol output must report them.
        var handle = new SafeFileHandle(GetStdHandle(-11), ownsHandle: false);
        try { return new FileStream(handle, FileAccess.Write, bufferSize: 1, isAsync: false); }
        catch { handle.Dispose(); throw; }
    }

    [DllImport("kernel32.dll")]
    private static extern IntPtr GetStdHandle(int kind);
}
