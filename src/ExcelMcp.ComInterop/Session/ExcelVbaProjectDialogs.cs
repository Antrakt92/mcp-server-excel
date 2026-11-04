using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Principal;
using Microsoft.Win32.SafeHandles;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal static partial class ExcelVbaProjectProtection
{
    // Both operations use the same initiating STA command/selection handshake.
    // Dialog recognition and cleanup remain specific to each requested operation.
    private interface IProjectDialogInteraction : IDisposable
    {
        Task Start();
        void WaitUntilReady();
        void MarkCommandRequested();
        void MarkCommandReturned();
        void StopWaitingForDialog();
        void ReconcileAfterCommand();
    }

    private sealed class AuthenticationDialogOwner : IDisposable
    {
        private readonly Process process;
        private readonly ExcelProcessIdentity identity;
        private readonly nint editorWindow;
        private readonly nint excelWindow;

        internal AuthenticationDialogOwner(ExcelProcessIdentity identity, nint editorWindow, nint excelWindow)
        {
            this.identity = identity;
            this.editorWindow = editorWindow;
            this.excelWindow = excelWindow;
            process = Process.GetProcessById(identity.ProcessId);
            try
            {
                _ = process.Handle;
                AssertProcess();
                using var current = Process.GetCurrentProcess();
                if (process.SessionId != current.SessionId ||
                    !string.Equals(Path.GetFileName(process.MainModule?.FileName), "EXCEL.EXE", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("VBA authentication requires Excel in the current Windows session.");
                if (!OpenProcessToken(process.Handle, 0x0008, out var token))
                    throw new InvalidOperationException("The Excel process user could not be confirmed.");
                using (token)
                using (var targetUser = new WindowsIdentity(token.DangerousGetHandle()))
                using (var currentUser = WindowsIdentity.GetCurrent())
                    if (targetUser.User is null || currentUser.User is null || !targetUser.User.Equals(currentUser.User))
                        throw new InvalidOperationException("The Excel process does not belong to the current user.");
                AssertWindow(editorWindow);
                AssertWindow(excelWindow);
            }
            catch { process.Dispose(); throw; }
        }

        internal void AssertProcess()
        {
            if (process.HasExited || process.StartTime.ToUniversalTime().ToFileTimeUtc() != identity.StartedAtUtcFileTime)
                throw new InvalidOperationException("The confirmed Excel process is no longer available.");
        }

        internal void AssertWindow(nint window)
        {
            AssertProcess();
            if (!IsWindow(window) || GetWindowThreadProcessId(window, out var owner) == 0 || owner != identity.ProcessId)
                throw new InvalidOperationException("VBA authentication window ownership changed.");
        }

        internal bool OwnedByEditor(nint window)
        {
            var parent = GetWindow(window, 4);
            for (int depth = 0; parent != 0 && depth < 32; depth++, parent = GetWindow(parent, 4))
                if (parent == editorWindow || parent == excelWindow) return true;
            return false;
        }

        internal List<nint> Windows()
        {
            AssertProcess();
            var windows = new List<nint>();
            if (!EnumWindows((window, _) =>
                {
                    if (GetWindowThreadProcessId(window, out var pid) != 0 && pid == identity.ProcessId) windows.Add(window);
                    return true;
                }, 0)) throw new InvalidOperationException("Excel authentication window enumeration failed.");
            return windows;
        }

        internal List<nint> Children(nint root)
        {
            AssertWindow(root);
            var children = new List<nint>();
            _ = EnumChildWindows(root, (window, _) => { children.Add(window); return true; }, 0);
            return children.Where(window => IsWindow(window) && IsChild(root, window) && IsWindowVisible(window)).ToList();
        }

        internal bool HasBlockingDialog() => Windows().Any(window => Class(window) == "#32770" && OwnedByEditor(window));

        internal bool RootsEnabled()
        {
            AssertWindow(editorWindow);
            AssertWindow(excelWindow);
            return IsWindowEnabled(editorWindow) && IsWindowEnabled(excelWindow);
        }

        public void Dispose() => process.Dispose();
    }

    [DllImport("advapi32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool OpenProcessToken(nint process, uint access, out SafeAccessTokenHandle token);
}
