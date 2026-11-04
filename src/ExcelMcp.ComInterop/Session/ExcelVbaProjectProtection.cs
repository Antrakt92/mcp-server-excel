using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Vbe = Microsoft.Vbe.Interop;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>Drives only the selected project's standard, process-owned protection dialog.</summary>
internal static partial class ExcelVbaProjectProtection
{
    internal readonly record struct AppliedProtection(string ProjectName, string CurrentProtection);

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Collects command, dialog and restoration failures, then rethrows every original failure after independent cleanup.")]
    internal static AppliedProtection Protect(ExcelContext context, string password)
    {
        if (context.ProcessIdentity is not { } identity)
            throw new InvalidOperationException("VBA protection requires confirmed Excel process ownership.");
        if (context.Book.ReadOnly)
            throw new InvalidOperationException("VBA protection cannot be configured in a read-only workbook.");

        Vbe.VBProject? project = null;
        Vbe.VBProject? previousProject = null;
        Vbe.VBE? editor = null;
        Vbe.Window? editorWindow = null;
        Office.CommandBars? bars = null;
        Office.CommandBarControl? properties = null;
        bool? visibleBefore = null;
        Exception? primary = null;
        AppliedProtection result = default;
        var progress = new ExcelVbaProtectionProgress();
        var dialogStarted = false;
        try
        {
            project = context.Book.VBProject;
            if (project.Protection != Vbe.vbext_ProjectProtection.vbext_pp_none)
                throw new UnauthorizedAccessException("The VBA project is locked. Its existing password must be authenticated before changing protection.");
            if (project.Mode != Vbe.vbext_VBAMode.vbext_vm_Design)
                throw new InvalidOperationException("Stop VBA execution before configuring project protection.");
            var name = project.Name;
            editor = project.VBE;
            editorWindow = editor.MainWindow;
            visibleBefore = editorWindow.Visible;
            previousProject = editor.ActiveVBProject;
            bars = editor.CommandBars;
            properties = bars.FindControl(Id: 2578);
            if (properties is null || !properties.Enabled)
                throw new InvalidOperationException("Excel's VBA Project Properties command is unavailable.");

            using var driver = new ProtectionDialog(identity, (nint)editorWindow.HWnd, (nint)context.App.Hwnd, name, password, progress);
            driver.AssertNoExistingDialog();
            editorWindow.Visible = true;
            ExecuteDialog(driver, properties, editor, project, ref dialogStarted);

            using (var verification = new ProtectionDialog(identity, (nint)editorWindow.HWnd, (nint)context.App.Hwnd, name, password, progress, verifyOnly: true))
            {
                verification.AssertNoExistingDialog();
                ExecuteDialog(verification, properties, editor, project, ref dialogStarted);
            }

            progress.State = progress.State with { Phase = "configuration-applied", ConfigurationApplied = true, SaveRequired = true, ReopenRequired = true };
            context.Book.Saved = false;
            result = new AppliedProtection(name, Convert.ToInt32(project.Protection, CultureInfo.InvariantCulture) switch
            {
                0 => "None",
                1 => "Locked",
                _ => "Unknown"
            });
        }
        catch (Exception error) { primary = error; }
        finally
        {
            try
            {
                if (progress.CanRestoreEditor && (!dialogStarted || progress.State.DialogCleanupConfirmed))
                {
                    if (editor is not null && previousProject is not null) editor.ActiveVBProject = previousProject;
                    progress.State = progress.State with { ActiveProjectRestored = true };
                }
            }
            catch (Exception error) { primary = Combine(primary, error); }
            try
            {
                if (progress.CanRestoreEditor && (!dialogStarted || progress.State.DialogCleanupConfirmed))
                {
                    if (editorWindow is not null && visibleBefore.HasValue && editorWindow.Visible != visibleBefore.Value)
                        editorWindow.Visible = visibleBefore.Value;
                    progress.State = progress.State with { VisibilityRestored = true };
                }
            }
            catch (Exception error) { primary = Combine(primary, error); }
            ComUtilities.Release(ref properties);
            ComUtilities.Release(ref bars);
            ComUtilities.Release(ref editorWindow);
            ComUtilities.Release(ref previousProject);
            ComUtilities.Release(ref editor);
            ComUtilities.Release(ref project);
        }
        if (primary is not null) throw new ExcelVbaProtectionException(primary, progress.State);
        return result;
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Drains the dedicated native helper and rethrows command and helper failures together before releasing ownership.")]
    private static void ExecuteDialog(IProjectDialogInteraction driver, Office.CommandBarControl command,
        Vbe.VBE editor, Vbe.VBProject target, ref bool commandAttempted)
    {
        var interaction = driver.Start();
        Exception? failure = null;
        try
        {
            driver.WaitUntilReady();
            // Showing VBE can select the active Excel workbook's project. Pin and
            // verify the requested project after visibility/readiness, for each
            // command, before allowing the native driver to enter the dialog.
            SelectProject(editor, target);
            commandAttempted = true;
            driver.MarkCommandRequested();
            command.Execute();
        }
        catch (Exception error) { failure = error; driver.StopWaitingForDialog(); }
        finally { driver.MarkCommandReturned(); }
        try { interaction.GetAwaiter().GetResult(); }
        catch (Exception error) { failure = Combine(failure, error); }
        if (failure is not null)
        {
            // Execute can return after the helper's first cleanup attempt. Reconcile
            // the same introduced dialog before releasing the operation owner.
            try { driver.ReconcileAfterCommand(); }
            catch (Exception cleanup) { failure = Combine(failure, cleanup); }
        }
        if (failure is not null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
    }

    private static void SelectProject(Vbe.VBE editor, Vbe.VBProject target)
    {
        var expectedName = target.Name;
        var expectedFile = target.FileName;
        Vbe.VBProject? selected = null;
        try
        {
            editor.ActiveVBProject = target;
            selected = editor.ActiveVBProject;
            if (string.IsNullOrEmpty(expectedFile) || selected.Name != expectedName ||
                !string.Equals(selected.FileName, expectedFile, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Excel did not select the requested VBA project; no project dialog was requested.");
        }
        finally { ComUtilities.Release(ref selected); }
    }

    private static Exception Combine(Exception? primary, Exception cleanup) => primary is null ? cleanup :
        new AggregateException("VBA project dialog operation and restoration failed.", primary, cleanup);

    private sealed class ProtectionDialog : IProjectDialogInteraction
    {
        private readonly Process process;
        private readonly ExcelProcessIdentity identity;
        private readonly nint editorWindow;
        private readonly nint excelWindow;
        private readonly string projectName;
        private readonly string password;
        private readonly bool verifyOnly;
        private readonly ExcelVbaProtectionProgress progress;
        private readonly Dictionary<nint, (int Id, string Class)> controls = [];
        private readonly Stopwatch deadline = Stopwatch.StartNew();
        private readonly CancellationTokenSource stopWaiting = new();
        private readonly ManualResetEventSlim commandCommitted = new();
        private volatile bool commandReturned;
        private nint dialog;
        private readonly TaskCompletionSource ready = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource completed = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private Thread? worker;
        private bool workerStarted;

        internal ProtectionDialog(ExcelProcessIdentity identity, nint editorWindow, nint excelWindow, string projectName, string password, ExcelVbaProtectionProgress progress, bool verifyOnly = false)
        {
            this.identity = identity;
            this.editorWindow = editorWindow;
            this.excelWindow = excelWindow;
            this.projectName = projectName;
            this.password = password;
            this.verifyOnly = verifyOnly;
            this.progress = progress;
            process = Process.GetProcessById(identity.ProcessId);
            try
            {
                progress.State = progress.State with { Phase = verifyOnly ? "opening-readback-dialog" : "opening-dialog" };
                _ = process.Handle;
                AssertProcess();
                using var current = Process.GetCurrentProcess();
                if (process.SessionId != current.SessionId) throw new InvalidOperationException("Excel is outside the current Windows session.");
                AssertWindow(editorWindow);
                AssertWindow(excelWindow);
            }
            catch { process.Dispose(); throw; }
        }

        public void StopWaitingForDialog() => stopWaiting.Cancel();

        public Task Start()
        {
            if (worker is not null) throw new InvalidOperationException("The native protection helper cannot be started twice.");
            worker = new Thread(RunWorker) { IsBackground = true, Name = "ExcelMCP VBA protection" };
            worker.SetApartmentState(ApartmentState.MTA);
            worker.Start();
            workerStarted = true;
            return completed.Task;
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "The initiating COM callback observes and rethrows helper failures through the completion task; readiness failures are propagated too.")]
        private void RunWorker()
        {
            try { Apply(); completed.TrySetResult(); }
            catch (Exception error) { ready.TrySetException(error); completed.TrySetException(error); }
        }

        public void WaitUntilReady() => ready.Task.WaitAsync(TimeSpan.FromSeconds(5)).GetAwaiter().GetResult();

        public void MarkCommandRequested()
        {
            if (completed.Task.IsCompleted) throw new InvalidOperationException("The native protection helper ended before the command was requested.");
            progress.State = progress.State with { DialogCleanupConfirmed = false };
            commandCommitted.Set();
        }

        public void MarkCommandReturned() => commandReturned = true;

        public void ReconcileAfterCommand()
        {
            if (commandCommitted.IsSet) CloseIntroducedDialog();
            ConfirmDialogCleanup();
        }

        private void AssertProcess()
        {
            if (process.HasExited || process.StartTime.ToUniversalTime().ToFileTimeUtc() != identity.StartedAtUtcFileTime)
                throw new InvalidOperationException("The owned Excel process is no longer available.");
        }

        private void AssertWindow(nint window)
        {
            AssertProcess();
            if (!IsWindow(window) || GetWindowThreadProcessId(window, out var owner) == 0 || owner != identity.ProcessId)
                throw new InvalidOperationException("VBA protection window ownership changed.");
        }

        private bool OwnedByEditor(nint window)
        {
            var parent = GetWindow(window, 4);
            for (var depth = 0; parent != 0 && depth < 32; depth++, parent = GetWindow(parent, 4))
                if (parent == editorWindow || parent == excelWindow) return true;
            return false;
        }

        internal void AssertNoExistingDialog()
        {
            progress.CanRestoreEditor = false;
            progress.State = progress.State with { DialogCleanupConfirmed = false };
            AssertWindow(editorWindow);
            if (Windows().Any(window => Class(window) == "#32770" && OwnedByEditor(window)))
                throw new InvalidOperationException("Close the existing Excel or VBA dialog before configuring protection.");
            progress.State = progress.State with { DialogCleanupConfirmed = true };
            progress.CanRestoreEditor = true;
        }

        private List<nint> Windows()
        {
            AssertProcess();
            var result = new List<nint>();
            if (!EnumWindows((window, _) =>
                {
                    if (GetWindowThreadProcessId(window, out var owner) != 0 && owner == identity.ProcessId) result.Add(window);
                    return true;
                }, 0)) throw new InvalidOperationException("Excel window enumeration failed.");
            return result;
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Retains the native operation failure and any exact-dialog cleanup failure and rethrows both after cleanup.")]
        internal void Apply()
        {
            Exception? primary = null;
            try
            {
                AssertProcess();
                ready.TrySetResult();
                commandCommitted.Wait(stopWaiting.Token);
                deadline.Restart();
                while (deadline.Elapsed < TimeSpan.FromSeconds(10))
                {
                    stopWaiting.Token.ThrowIfCancellationRequested();
                    var matches = Windows().Where(window => Class(window) == "#32770" && OwnedByEditor(window) &&
                        Title(window).StartsWith(projectName + " -", StringComparison.Ordinal)).ToArray();
                    if (matches.Length > 1) throw new InvalidOperationException("The VBA project dialog is ambiguous.");
                    if (matches.Length == 1)
                    {
                        if (dialog != 0 && dialog != matches[0]) throw new InvalidOperationException("The VBA project dialog changed before initialization completed.");
                        dialog = matches[0];
                        if (IsWindowVisible(dialog) && HasControl(12320, "SysTabControl32")) break;
                    }
                    Thread.Sleep(25);
                }
                if (dialog == 0 || !IsWindowVisible(dialog)) throw new TimeoutException("Excel did not open the expected VBA project dialog.");
                AssertDialog();
                var tab = Control(12320, "SysTabControl32");
                if ((GetWindowLongPtr(tab, -16).ToInt64() & 0x100) != 0)
                    throw new InvalidOperationException("This Excel project dialog uses an unsupported tab style.");
                Send(tab, 0x1330, 1);
                if (Send(tab, 0x130B) != 1) throw new InvalidOperationException("The Protection page did not become selected.");
                var checkbox = Control(5463, "Button");
                var first = Control(5461, "Edit");
                var confirmation = Control(5462, "Edit");
                if (verifyOnly)
                {
                    if (Send(checkbox, 0xF0) != 1)
                        throw new InvalidOperationException("VBA protection read-back found project locking disabled.");
                    progress.State = progress.State with { LockOptionReadback = true };
                }
                else
                {
                    foreach (var edit in new[] { first, confirmation })
                    {
                        if ((GetWindowLongPtr(edit, -16).ToInt64() & 0x20) == 0)
                            throw new InvalidOperationException("The expected password control is not masked.");
                        if ((nuint)password.Length > Send(edit, 0xD5))
                            throw new ArgumentException("The password exceeds this Excel dialog's supported length.", nameof(password));
                    }
                    foreach (var edit in new[] { first, confirmation })
                    {
                        progress.State = progress.State with { Phase = "writing-password" };
                        AssertControl(edit);
                        AssertDeadline();
                        if (SendText(edit, 0xC, 0, string.Empty, 0x23, 500, out var result) == 0 || result == 0)
                            throw new InvalidOperationException("Excel did not initialize a password field; completion is unconfirmed.");
                        // Address the password edit itself so its character handler receives the input;
                        // changing display text alone does not prove the editor updated its password state.
                        foreach (var character in password) Send(edit, 0x102, character, parameter: 1);
                        if (Send(edit, 0xE) != (nuint)password.Length)
                            throw new InvalidOperationException("Excel truncated a password field; protection was not confirmed.");
                    }
                    if (Send(checkbox, 0xF0) == 0) Send(checkbox, 0xF5);
                    if (Send(checkbox, 0xF0) != 1) throw new InvalidOperationException("Excel did not enable project locking.");
                    progress.State = progress.State with { Phase = "requesting-acceptance", AcceptanceRequested = true, ConfigurationApplied = null, SaveRequired = null, ReopenRequired = null };
                    PostClick(Control(1, "Button"));
                    var settling = Stopwatch.StartNew();
                    while (IsWindow(dialog) && settling.Elapsed < TimeSpan.FromSeconds(2)) { AssertProcess(); Thread.Sleep(25); }
                    if (IsWindow(dialog)) throw new InvalidOperationException("Excel did not confirm project protection; completion is unconfirmed.");
                    dialog = 0;
                }
            }
            catch (Exception error) { primary = error; }
            finally
            {
                // Dismiss only the exact newly opened dialog, never another matching prompt.
                try
                {
                    if (commandCommitted.IsSet)
                    {
                        CloseIntroducedDialog();
                        ConfirmDialogCleanup();
                    }
                }
                catch (Exception cleanup)
                {
                    primary = primary is null ? cleanup : new AggregateException("VBA protection and dialog cleanup failed.", primary, cleanup);
                }
            }
            if (primary is not null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(primary).Throw();
        }

        private void CloseIntroducedDialog()
        {
            var closing = Stopwatch.StartNew();
            var requested = false;
            var retry = TimeSpan.Zero;
            do
            {
                AssertProcess();
                var matches = Windows().Where(window => Class(window) == "#32770" && OwnedByEditor(window) &&
                    Title(window).StartsWith(projectName + " -", StringComparison.Ordinal)).ToArray();
                if (matches.Length > 1) throw new InvalidOperationException("The introduced VBA project dialog is ambiguous during cleanup.");
                if (matches.Length == 1)
                {
                    if (dialog != 0 && IsWindow(dialog) && dialog != matches[0])
                        throw new InvalidOperationException("The introduced VBA project dialog changed during cleanup.");
                    dialog = matches[0];
                    AssertDialog();
                    // A newly created dialog can receive WM_CLOSE before its handler
                    // is ready. Wait for its standard controls and retry cancellation
                    // within a separate cleanup budget, independent of caller waits.
                    if (HasControl(12320, "SysTabControl32") && HasControl(2, "Button") && closing.Elapsed >= retry)
                    {
                        if (!PostMessage(dialog, 0x10, 0, 0))
                            throw new InvalidOperationException("The exact introduced VBA protection dialog did not accept cancellation.");
                        requested = true;
                        retry = closing.Elapsed + TimeSpan.FromMilliseconds(250);
                    }
                }
                else if (commandReturned && (requested || dialog == 0 || !IsWindow(dialog))) return;
                Thread.Sleep(25);
            } while (closing.Elapsed < TimeSpan.FromSeconds(2));
            if (dialog != 0 && IsWindow(dialog)) throw new InvalidOperationException("The introduced VBA protection dialog remains open.");
        }

        private void ConfirmDialogCleanup()
        {
            progress.State = progress.State with { DialogCleanupConfirmed = false };
            var closing = Stopwatch.StartNew();
            do
            {
                AssertWindow(editorWindow);
                AssertWindow(excelWindow);
                if (commandReturned && !Windows().Any(window => Class(window) == "#32770" && OwnedByEditor(window)) &&
                    IsWindowEnabled(editorWindow) && IsWindowEnabled(excelWindow))
                {
                    progress.State = progress.State with { DialogCleanupConfirmed = true };
                    return;
                }
                Thread.Sleep(25);
            } while (closing.Elapsed < TimeSpan.FromSeconds(1));
            throw new InvalidOperationException("Excel or VBA still has a blocking dialog; editor restoration was not attempted.");
        }

        private void AssertDialog()
        {
            AssertWindow(dialog);
            if (Class(dialog) != "#32770" || !OwnedByEditor(dialog) ||
                !Title(dialog).StartsWith(projectName + " -", StringComparison.Ordinal))
                throw new InvalidOperationException("The VBA project dialog does not belong to the selected project.");
        }

        private nint Control(int id, string expectedClass)
        {
            var waiting = Stopwatch.StartNew();
            do
            {
                AssertDeadline(); AssertDialog();
                var matches = FindControls(id, expectedClass);
                if (matches.Count > 1) throw new InvalidOperationException($"VBA protection control {id} is ambiguous.");
                if (matches.Count == 1)
                {
                    var result = matches[0];
                    AssertWindow(result);
                    if (!IsChild(dialog, result) || !IsWindowEnabled(result)) throw new InvalidOperationException("The VBA protection control is unavailable.");
                    controls[result] = (id, expectedClass);
                    return result;
                }
                Thread.Sleep(25);
            } while (waiting.Elapsed < TimeSpan.FromSeconds(1));
            throw new InvalidOperationException($"Required VBA protection control {id} is absent.");
        }

        private bool HasControl(int id, string expectedClass) => FindControls(id, expectedClass).Count == 1;

        private List<nint> FindControls(int id, string expectedClass)
        {
            AssertDialog();
            var matches = new List<nint>();
            var children = new List<nint>();
            // Win32 documents no usable return value for EnumChildWindows.
            _ = EnumChildWindows(dialog, (window, _) => { children.Add(window); return true; }, 0);
            // Keep fallible identity checks outside the unmanaged enumeration callback.
            foreach (var window in children)
            {
                if (!IsWindow(window) || GetDlgCtrlID(window) != id || !IsWindowVisible(window)) continue;
                if (Class(window) == expectedClass)
                {
                    AssertWindow(window);
                    matches.Add(window);
                }
            }
            return matches;
        }

        private void AssertControl(nint window)
        {
            AssertDialog(); AssertWindow(window);
            if (!controls.TryGetValue(window, out var expected) || !IsChild(dialog, window) ||
                GetDlgCtrlID(window) != expected.Id || Class(window) != expected.Class || !IsWindowEnabled(window))
                throw new InvalidOperationException("The selected VBA protection control changed before input.");
        }

        private void AssertDeadline()
        {
            stopWaiting.Token.ThrowIfCancellationRequested();
            if (deadline.Elapsed > TimeSpan.FromSeconds(15)) throw new TimeoutException("The VBA protection operation exceeded its native deadline.");
        }

        private void PostClick(nint button, bool cleanup = false)
        {
            AssertControl(button);
            if (!cleanup) AssertDeadline();
            // A button can open another modal loop. Do not synchronously block its UI driver.
            if (!PostMessage(button, 0xF5, 0, 0))
                throw new InvalidOperationException("The owned VBA dialog did not accept its button request; completion is unconfirmed.");
        }

        private nuint Send(nint window, uint message, nuint value = 0, bool cleanup = false, nint parameter = 0)
        {
            AssertControl(window);
            if (!cleanup) AssertDeadline();
            if (SendMessageTimeout(window, message, value, parameter, 0x23, 500, out var result) == 0)
                throw new InvalidOperationException("A VBA protection control did not respond; completion is unconfirmed.");
            return result;
        }

        public void Dispose()
        {
            if (workerStarted && worker is not null && !worker.Join(TimeSpan.FromSeconds(2)))
                throw new InvalidOperationException("The native VBA protection helper did not finish; ownership remains unconfirmed.");
            stopWaiting.Dispose();
            commandCommitted.Dispose();
            process.Dispose();
        }
    }

    private static string Class(nint window)
    {
        var buffer = new char[128];
        var length = GetClassName(window, buffer, buffer.Length);
        if (length == 0 && IsWindow(window)) throw new InvalidOperationException("VBA window class could not be confirmed.");
        return new string(buffer, 0, length);
    }

    private static string Title(nint window)
    {
        var buffer = new char[512];
        var length = GetWindowText(window, buffer, buffer.Length);
        // Empty titles are normal while a modal window is being initialized.
        // They never satisfy the strict project identity check before input.
        return new string(buffer, 0, length);
    }

    private delegate bool WindowCallback(nint window, nint data);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool EnumWindows(WindowCallback callback, nint data);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool EnumChildWindows(nint parent, WindowCallback callback, nint data);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(nint window, out uint processId);
    [DllImport("user32.dll")] private static extern int GetDlgCtrlID(nint window);
    [DllImport("user32.dll")] private static extern nint GetWindow(nint window, uint command);
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool PostMessage(nint window, uint message, nuint wParam, nint lParam);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsChild(nint parent, nint window);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsWindow(nint window);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsWindowVisible(nint window);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsWindowEnabled(nint window);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(nint window, [Out] char[] text, int capacity);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(nint window, [Out] char[] text, int capacity);
    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] private static extern nint GetWindowLongPtr(nint window, int index);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern nint SendMessageTimeout(nint window, uint message, nuint wParam, nint lParam, uint flags, uint milliseconds, out nuint result);
    [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern nint SendText(nint window, uint message, nuint wParam, string value, uint flags, uint milliseconds, out nuint result);
}
