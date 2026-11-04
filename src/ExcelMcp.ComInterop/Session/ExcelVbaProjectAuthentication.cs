using System.Diagnostics;
using Office = Microsoft.Office.Core;
using Vbe = Microsoft.Vbe.Interop;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal static partial class ExcelVbaProjectProtection
{
    internal readonly record struct AuthenticatedProject(string ProjectName, string ProtectionBefore,
        string ProtectionAfter, bool AlreadyAccessible, ExcelVbaAuthenticationState State);

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Preserves command, native challenge and independent restoration failures with their final observed state.")]
    internal static AuthenticatedProject Authenticate(ExcelContext context, string password)
    {
        if (context.ProcessIdentity is not { } identity)
            throw new InvalidOperationException("VBA authentication requires confirmed Excel process identity.");
        Vbe.VBProject? project = null;
        Vbe.VBProject? previous = null;
        Vbe.VBE? editor = null;
        Vbe.Window? window = null;
        Office.CommandBars? bars = null;
        Office.CommandBarControl? command = null;
        var progress = new ExcelVbaAuthenticationProgress();
        bool? visibleBefore = null;
        Exception? failure = null;
        var attempted = false;
        var name = string.Empty;
        var before = string.Empty;
        var after = string.Empty;
        var alreadyAccessible = false;
        try
        {
            project = context.Book.VBProject;
            name = project.Name;
            if (project.Mode != Vbe.vbext_VBAMode.vbext_vm_Design)
                throw new InvalidOperationException("Stop VBA execution before authenticating a project.");
            var initialProtection = project.Protection;
            alreadyAccessible = initialProtection == Vbe.vbext_ProjectProtection.vbext_pp_none;
            before = initialProtection switch
            {
                Vbe.vbext_ProjectProtection.vbext_pp_none => "None",
                Vbe.vbext_ProjectProtection.vbext_pp_locked => "Locked",
                _ => "Unknown"
            };
            progress.State = progress.State with { CurrentProtection = before };
            if (alreadyAccessible)
            {
                after = before;
                progress.State = progress.State with
                {
                    Phase = "already-accessible",
                    AlreadyAccessible = true,
                    DialogCleanupConfirmed = true,
                    ActiveProjectRestored = true,
                    VisibilityRestored = true
                };
            }
            else
            {
                if (initialProtection != Vbe.vbext_ProjectProtection.vbext_pp_locked)
                    throw new InvalidOperationException("Excel returned an unsupported VBA protection state.");
                editor = project.VBE;
                window = editor.MainWindow;
                previous = editor.ActiveVBProject;
                visibleBefore = window.Visible;
                bars = editor.CommandBars;
                command = bars.FindControl(Id: 2578);
                if (command is null || !command.Enabled)
                    throw new InvalidOperationException("Excel's VBA Project Properties command is unavailable.");
                // No editor state was changed yet. A failed identity/preflight
                // must not make restoration calls into an unconfirmed modal UI.
                progress.CanRestoreEditor = false;
                using var driver = new AuthenticationDialog(identity, (nint)window.HWnd,
                    (nint)context.App.Hwnd, name, password, progress);
                driver.AssertInitiallyClear();
                window.Visible = true;
                ExecuteDialog(driver, command, editor, project, ref attempted);
                after = project.Protection switch
                {
                    Vbe.vbext_ProjectProtection.vbext_pp_none => "None",
                    Vbe.vbext_ProjectProtection.vbext_pp_locked => "Locked",
                    _ => throw new InvalidOperationException("Excel returned an unsupported VBA protection state after authentication.")
                };
                progress.State = progress.State with { CurrentProtection = after };
                if (driver.Outcome == AuthenticationOutcome.Accepted && after == "None")
                    progress.State = progress.State with { Phase = "authenticated", PasswordVerified = true, AccessUnlocked = true };
                else if (driver.Outcome == AuthenticationOutcome.Rejected && after == "Locked")
                    progress.State = progress.State with { Phase = "password-rejected", PasswordVerified = false, AccessUnlocked = false };
                else
                    throw new InvalidOperationException("The native authentication outcome did not match the selected project's protection state.");
            }
        }
        catch (Exception error) { failure = error; }
        finally
        {
            if (progress.CanRestoreEditor && (!attempted || progress.State.DialogCleanupConfirmed))
            {
                try
                {
                    if (editor is not null && previous is not null) editor.ActiveVBProject = previous;
                    progress.State = progress.State with { ActiveProjectRestored = true };
                }
                catch (Exception error) { failure = Combine(failure, error); }
                try
                {
                    if (window is not null && visibleBefore.HasValue && window.Visible != visibleBefore.Value)
                        window.Visible = visibleBefore.Value;
                    progress.State = progress.State with { VisibilityRestored = true };
                }
                catch (Exception error) { failure = Combine(failure, error); }
            }
            ComUtilities.Release(ref command);
            ComUtilities.Release(ref bars);
            ComUtilities.Release(ref window);
            ComUtilities.Release(ref previous);
            ComUtilities.Release(ref editor);
            ComUtilities.Release(ref project);
        }
        if (failure is not null) throw new ExcelVbaAuthenticationException(failure, progress.State);
        return new(name, before, after, alreadyAccessible, progress.State);
    }

    private enum AuthenticationOutcome { Unknown, Accepted, Rejected }
    private enum AuthenticationDialogKind { Password, InvalidPassword, Properties }

    private sealed class AuthenticationDialog : IProjectDialogInteraction
    {
        private static readonly AuthenticationDialogKind[] CleanupOrder =
            [AuthenticationDialogKind.InvalidPassword, AuthenticationDialogKind.Password, AuthenticationDialogKind.Properties];
        private readonly AuthenticationDialogOwner owner;
        private readonly string projectName;
        private readonly string password;
        private readonly ExcelVbaAuthenticationProgress progress;
        private readonly HashSet<nint> baseline = [];
        private readonly nint editorWindow;
        private nint challengeEditorOwner;
        private readonly Dictionary<nint, AuthenticationDialogKind> introduced = [];
        private readonly Dictionary<nint, (nint Root, int Id, string Class)> controls = [];
        private readonly CancellationTokenSource stopWaiting = new();
        private readonly ManualResetEventSlim committed = new();
        private readonly TaskCompletionSource ready = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource completed = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly Stopwatch deadline = new();
        private volatile bool returned;
        private Thread? worker;
        private bool workerStarted;
        private nint authentication;

        internal AuthenticationOutcome Outcome { get; private set; }

        internal AuthenticationDialog(ExcelProcessIdentity identity, nint editorWindow, nint excelWindow,
            string projectName, string password, ExcelVbaAuthenticationProgress progress)
        {
            owner = new AuthenticationDialogOwner(identity, editorWindow, excelWindow);
            this.editorWindow = editorWindow;
            this.projectName = projectName;
            this.password = password;
            this.progress = progress;
        }

        internal void AssertInitiallyClear()
        {
            progress.CanRestoreEditor = false;
            progress.State = progress.State with { Phase = "opening-challenge", DialogCleanupConfirmed = false };
            foreach (var window in owner.Windows()) baseline.Add(window);
            if (owner.HasBlockingDialog())
                throw new InvalidOperationException("Close the existing Excel or VBA dialog before authenticating a project.");
            progress.State = progress.State with { DialogCleanupConfirmed = true };
            progress.CanRestoreEditor = true;
        }

        public Task Start()
        {
            if (worker is not null) throw new InvalidOperationException("The native authentication helper cannot be started twice.");
            worker = new Thread(Run) { IsBackground = true, Name = "ExcelMCP VBA authentication" };
            worker.SetApartmentState(ApartmentState.MTA);
            worker.Start();
            workerStarted = true;
            return completed.Task;
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "The initiating STA observes both readiness and helper failures through the completion tasks.")]
        private void Run()
        {
            try { Apply(); completed.TrySetResult(); }
            catch (Exception error) { ready.TrySetException(error); completed.TrySetException(error); }
        }

        public void WaitUntilReady() => ready.Task.WaitAsync(TimeSpan.FromSeconds(5)).GetAwaiter().GetResult();
        public void StopWaitingForDialog() => stopWaiting.Cancel();
        public void MarkCommandReturned() => returned = true;
        public void MarkCommandRequested()
        {
            if (completed.Task.IsCompleted) throw new InvalidOperationException("The authentication helper ended before command commitment.");
            progress.State = progress.State with { DialogCleanupConfirmed = false };
            committed.Set();
        }

        public void ReconcileAfterCommand()
        {
            if (committed.IsSet)
            {
                Cleanup();
                return;
            }
            // A failed selection/readiness never introduced a challenge. Do not
            // adopt or dismiss a dialog that appeared before command commitment.
            progress.CanRestoreEditor = false;
            progress.State = progress.State with { DialogCleanupConfirmed = false };
            if (!returned || owner.HasBlockingDialog() || !owner.RootsEnabled())
                throw new InvalidOperationException("Excel has an unconfirmed dialog before authentication commitment; nothing was answered.");
            progress.State = progress.State with { DialogCleanupConfirmed = true };
            progress.CanRestoreEditor = true;
        }

        private void AssertDeadline()
        {
            stopWaiting.Token.ThrowIfCancellationRequested();
            owner.AssertProcess();
            if (deadline.Elapsed > TimeSpan.FromSeconds(20))
                throw new TimeoutException("The native VBA authentication deadline expired; outcome is unconfirmed.");
        }

        private List<nint> Dialogs() => owner.Windows().Where(window => !baseline.Contains(window)
            && Class(window) == "#32770" && owner.OwnedByEditor(window)).ToList();

        private bool Matches(nint root, AuthenticationDialogKind kind)
        {
            if (!IsWindow(root) || !IsWindowVisible(root) || baseline.Contains(root) ||
                Class(root) != "#32770" || !owner.OwnedByEditor(root)) return false;
            owner.AssertWindow(root);
            var children = owner.Children(root);
            bool Button(int id) => children.Count(child => GetDlgCtrlID(child) == id && Class(child) == "Button") == 1;
            var edits = children.Where(child => Class(child) == "Edit").ToArray();
            bool tab = children.Count(child => GetDlgCtrlID(child) == 12320 && Class(child) == "SysTabControl32") == 1;
            return kind switch
            {
                // The English caption, Edit 5470 and immediate owner were observed
                // in an owned native fixture. Other localizations fail closed.
                // A generic masked Edit is not proof of a VBA password challenge.
                AuthenticationDialogKind.Password => VbaAuthenticationChallengePolicy.CanRecognize(
                    projectName, Title(root), committed.IsSet, baseline.Contains(root), GetWindow(root, 4) == editorWindow,
                    edits.Length, edits.Length == 1 && (GetWindowLongPtr(edits[0], -16).ToInt64() & 0x20) != 0,
                    edits.Length == 1 ? GetDlgCtrlID(edits[0]) : 0,
                    children.Count(child => Class(child) == "Button"),
                    Button(1) && children.Any(child => GetDlgCtrlID(child) == 1 && Class(child) == "Button"
                        && Title(child).Replace("&", "", StringComparison.Ordinal) == "OK"),
                    Button(2) && children.Any(child => GetDlgCtrlID(child) == 2 && Class(child) == "Button"
                        && Title(child).Replace("&", "", StringComparison.Ordinal) == "Cancel"), tab),
                AuthenticationDialogKind.Properties => authentication != 0 && tab && Button(1) && Button(2)
                    && VbaAuthenticationChallengePolicy.IsPropertiesReplacement(projectName, Title(root),
                        progress.State.AuthenticationPerformed, IsWindow(authentication), challengeEditorOwner,
                        GetWindow(root, 4), Dialogs().Count),
                AuthenticationDialogKind.InvalidPassword => IsWindow(authentication)
                    && VbaAuthenticationChallengePolicy.IsInvalidPasswordOwner(authentication, GetWindow(root, 4))
                    && Title(root) == "Project Locked" && Button(2)
                    && children.Any(child => Class(child) == "Static" && Title(child) == "Invalid password")
                    && children.Any(child => GetDlgCtrlID(child) == 2 && Class(child) == "Button"
                        && Title(child).Replace("&", "", StringComparison.Ordinal) == "OK"),
                _ => false
            };
        }

        private void Remember(nint root, AuthenticationDialogKind kind)
        {
            // Only the operation's discovery phase may record a newly introduced
            // expected dialog. Cleanup must never adopt windows that appear later.
            if (!introduced.ContainsKey(root) && (!committed.IsSet ||
                (kind == AuthenticationDialogKind.Password
                    ? progress.State.Phase != "opening-challenge"
                    : !progress.State.AuthenticationPerformed)))
                throw new InvalidOperationException("The dialog appeared outside its expected authentication phase; nothing was answered.");
            if (!Matches(root, kind) || introduced.TryGetValue(root, out var previous) && previous != kind)
                throw new InvalidOperationException("The expected authentication dialog signature or identity changed.");
            if (kind == AuthenticationDialogKind.Password)
            {
                if (authentication != 0 && authentication != root)
                    throw new InvalidOperationException("The introduced authentication dialog is ambiguous.");
                if (authentication == 0)
                {
                    // Native inventory proves the challenge's immediate owner is
                    // the exact VBE main window, not an arbitrary Excel ancestor.
                    challengeEditorOwner = GetWindow(root, 4);
                    owner.AssertWindow(challengeEditorOwner);
                    if (challengeEditorOwner != editorWindow)
                        throw new InvalidOperationException("The password challenge's immediate editor owner changed.");
                }
                authentication = root;
            }
            introduced[root] = kind;
        }

        private nint Control(nint root, AuthenticationDialogKind kind, int? id, string expectedClass)
        {
            var waiting = Stopwatch.StartNew();
            do
            {
                Remember(root, kind);
                var matches = owner.Children(root).Where(child => Class(child) == expectedClass
                    && (!id.HasValue || GetDlgCtrlID(child) == id.Value)).ToArray();
                if (matches.Length > 1) throw new InvalidOperationException("The required authentication control is ambiguous.");
                if (matches.Length == 1 && IsWindowEnabled(matches[0]))
                {
                    var result = matches[0];
                    owner.AssertWindow(result);
                    controls[result] = (root, GetDlgCtrlID(result), expectedClass);
                    return result;
                }
                Thread.Sleep(25);
            } while (waiting.Elapsed < TimeSpan.FromSeconds(1));
            throw new InvalidOperationException("The required authentication control is missing or unavailable.");
        }

        private void AssertControl(nint control)
        {
            owner.AssertWindow(control);
            if (!controls.TryGetValue(control, out var expected) || !introduced.TryGetValue(expected.Root, out var kind)
                || !Matches(expected.Root, kind) || !IsChild(expected.Root, control)
                || GetDlgCtrlID(control) != expected.Id || Class(control) != expected.Class || !IsWindowEnabled(control))
                throw new InvalidOperationException("The authentication control identity changed before input.");
        }

        private nuint Send(nint control, uint message, nuint value = 0, nint parameter = 0)
        {
            AssertDeadline(); AssertControl(control);
            if (SendMessageTimeout(control, message, value, parameter, 0x23, 500, out var result) == 0)
                throw new InvalidOperationException("An authentication control did not respond; outcome is unconfirmed.");
            return result;
        }

        private void Click(nint root, AuthenticationDialogKind kind, int id)
        {
            var button = Control(root, kind, id, "Button");
            AssertControl(button);
            if (!PostMessage(button, 0xF5, 0, 0))
                throw new InvalidOperationException("The expected authentication button request was not accepted.");
        }

        private void WaitClosed(nint root)
        {
            var wait = Stopwatch.StartNew();
            while (IsWindow(root) && wait.Elapsed < TimeSpan.FromSeconds(1.5))
            {
                owner.AssertProcess(); Thread.Sleep(25);
            }
            if (IsWindow(root)) throw new InvalidOperationException("An introduced authentication dialog remains open.");
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Retains the challenge failure and every independent exact-dialog cleanup failure and rethrows them together.")]
        private void Apply()
        {
            Exception? failure = null;
            try
            {
                owner.AssertProcess();
                ready.TrySetResult();
                committed.Wait(stopWaiting.Token);
                deadline.Restart();
                var opening = Stopwatch.StartNew();
                while (opening.Elapsed < TimeSpan.FromSeconds(10))
                {
                    AssertDeadline();
                    var matches = Dialogs().Where(root => Matches(root, AuthenticationDialogKind.Password)).ToArray();
                    if (matches.Length > 1) throw new InvalidOperationException("The authentication dialog is ambiguous.");
                    if (matches.Length == 1) { Remember(matches[0], AuthenticationDialogKind.Password); break; }
                    Thread.Sleep(25);
                }
                if (authentication == 0) throw new TimeoutException("Excel did not open the selected project's authentication challenge.");
                progress.State = progress.State with { Phase = "writing-password", ChallengeObserved = true };
                var edit = Control(authentication, AuthenticationDialogKind.Password, 5470, "Edit");
                // Reject before the first character; native truncation must never
                // turn a longer submitted password into an accepted shorter one.
                if ((nuint)password.Length > Send(edit, 0xD5))
                    throw new ArgumentException("The password exceeds this Excel authentication dialog's supported length.", nameof(password));
                if (Send(edit, 0xE) != 0)
                    throw new InvalidOperationException("The new password challenge was not empty; no password was submitted.");
                foreach (var character in password) Send(edit, 0x102, character, 1);
                if (Send(edit, 0xE) != (nuint)password.Length)
                    throw new InvalidOperationException("Excel truncated authentication input; no password was submitted.");
                progress.State = progress.State with
                {
                    Phase = "requesting-authentication",
                    PasswordVerified = null,
                    AccessUnlocked = null
                };
                Click(authentication, AuthenticationDialogKind.Password, 1);
                progress.State = progress.State with { AuthenticationPerformed = true };
                var waiting = Stopwatch.StartNew();
                while (waiting.Elapsed < TimeSpan.FromSeconds(10))
                {
                    AssertDeadline();
                    var dialogs = Dialogs();
                    var accepted = dialogs.Where(root => Matches(root, AuthenticationDialogKind.Properties)).ToArray();
                    var rejected = dialogs.Where(root => Matches(root, AuthenticationDialogKind.InvalidPassword)).ToArray();
                    if (accepted.Length + rejected.Length > 1)
                        throw new InvalidOperationException("The authentication outcome is ambiguous.");
                    if (accepted.Length == 1)
                    {
                        Remember(accepted[0], AuthenticationDialogKind.Properties);
                        Outcome = AuthenticationOutcome.Accepted;
                        progress.State = progress.State with { Phase = "acceptance-observed" };
                        Click(accepted[0], AuthenticationDialogKind.Properties, 2);
                        WaitClosed(accepted[0]);
                        break;
                    }
                    if (rejected.Length == 1)
                    {
                        Remember(rejected[0], AuthenticationDialogKind.InvalidPassword);
                        Outcome = AuthenticationOutcome.Rejected;
                        progress.State = progress.State with { Phase = "rejection-observed" };
                        Click(rejected[0], AuthenticationDialogKind.InvalidPassword, 2);
                        WaitClosed(rejected[0]);
                        Click(authentication, AuthenticationDialogKind.Password, 2);
                        WaitClosed(authentication);
                        break;
                    }
                    Thread.Sleep(25);
                }
                if (Outcome == AuthenticationOutcome.Unknown)
                    throw new TimeoutException("Excel did not confirm password acceptance or a recognized rejection.");
            }
            catch (Exception error) { failure = error; }
            finally
            {
                if (committed.IsSet)
                {
                    try { Cleanup(); }
                    catch (Exception cleanup) { failure = Combine(failure, cleanup); }
                }
            }
            if (failure is not null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Independently attempts only known introduced dialogs, collecting cleanup failures without hiding the original operation error.")]
        private void Cleanup()
        {
            progress.CanRestoreEditor = false;
            progress.State = progress.State with { DialogCleanupConfirmed = false };
            List<Exception> failures = [];
            var waiting = Stopwatch.StartNew();
            while (waiting.Elapsed < TimeSpan.FromSeconds(4))
            {
                owner.AssertProcess();
                foreach (var kind in CleanupOrder)
                    foreach (var entry in introduced.ToArray())
                    {
                        var root = entry.Key;
                        try
                        {
                            if (!IsWindow(root) || entry.Value != kind) continue;
                            if (!VbaAuthenticationChallengePolicy.CanCleanUp(introduced.ContainsKey(root), Matches(root, kind)))
                                throw new InvalidOperationException("An introduced dialog changed identity; cleanup did not answer it.");
                            Click(root, kind, 2);
                            WaitClosed(root);
                        }
                        catch (Exception error) { failures.Add(error); }
                    }
                if (returned && !owner.HasBlockingDialog() && owner.RootsEnabled())
                {
                    progress.State = progress.State with { DialogCleanupConfirmed = true };
                    progress.CanRestoreEditor = true;
                    if (failures.Count != 0) throw new AggregateException("Authentication dialog cleanup encountered failures.", failures);
                    return;
                }
                Thread.Sleep(25);
            }
            failures.Add(new InvalidOperationException("Authentication dialog cleanup or command completion remains unconfirmed; editor restoration was not attempted."));
            throw new AggregateException("VBA authentication cleanup failed.", failures);
        }

        public void Dispose()
        {
            if (workerStarted && worker is not null && !worker.Join(TimeSpan.FromSeconds(2)))
                throw new InvalidOperationException("The native authentication helper did not finish; ownership remains unconfirmed.");
            committed.Dispose();
            stopWaiting.Dispose();
            owner.Dispose();
        }
    }
}
