using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.Tests.Native;

internal static class Program
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    [MTAThread]
    private static int Main(string[] arguments)
    {
        Options? options = null;
        AuthProbe? probe = null;
        var result = new ProbeResult();
        try
        {
            options = Options.Parse(arguments);
            result.Mode = options.Mode;
            if (File.Exists(options.Ready) || File.Exists(options.Result))
                throw new ProbeException("output_already_exists");
            Console.InputEncoding = new UTF8Encoding(false, true);
            var password = ReadPassword();
            if (password.Length == 0 || password.Any(char.IsControl))
                throw new ProbeException("invalid_stdin_password");
            probe = new AuthProbe(options.Identity);
            result.ProcessIdentityConfirmed = true;
            probe.AssertInitiallyClear();
            WriteNewJson(options.Ready, new { ready = true });
            probe.Authenticate(password, options.Mode);
            result.ExpectedOutcomeObserved = true;
        }
        catch (Exception error)
        {
            result.Errors.Add(ErrorCode(error));
        }
        finally
        {
            if (probe is not null)
            {
                result.PasswordControlUnicode = probe.PasswordControlUnicode;
                result.PasswordControlSupportsUnichar = probe.PasswordControlSupportsUnichar;
                result.WindowsAnsiCodePage = probe.WindowsAnsiCodePage;
                try { result.CleanupConfirmed = probe.Cleanup(result.Errors); }
                catch (Exception error) { result.Errors.Add(ErrorCode(error)); }
                try { probe.Dispose(); }
                catch (Exception error) { result.Errors.Add(ErrorCode(error)); }
            }
            result.Success = result.ExpectedOutcomeObserved && result.CleanupConfirmed && result.Errors.Count == 0;
            if (options is not null)
            {
                try { WriteNewJson(options.Result, result); }
                catch (Exception) { result.Success = false; Console.Error.WriteLine("Authentication probe result could not be written."); }
            }
            Console.WriteLine(result.Success ? "PASS: expected synthetic authentication outcome and dialog cleanup confirmed." : "FAIL: synthetic authentication outcome or dialog cleanup was not confirmed.");
        }
        return result.Success ? 0 : 1;
    }

    private static string ReadPassword()
    {
        var buffer = new char[1025];
        var count = 0;
        while (count < buffer.Length)
        {
            var read = Console.In.Read(buffer, count, buffer.Length - count);
            if (read == 0) return new string(buffer, 0, count);
            count += read;
        }
        throw new ProbeException("stdin_password_too_long");
    }

    private static string ErrorCode(Exception error) => error is ProbeException known
        ? known.Code
        : "probe_failure_" + error.GetType().Name;

    private static void WriteNewJson<T>(string path, T value)
    {
        using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read);
        JsonSerializer.Serialize(stream, value, JsonOptions);
        stream.Flush(flushToDisk: true);
    }

    private sealed record Options(string Identity, string Mode, string Ready, string Result)
    {
        internal static Options Parse(string[] arguments)
        {
            if (arguments.Length != 8) throw new ProbeException("invalid_arguments");
            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            for (var index = 0; index < arguments.Length; index += 2)
            {
                var key = arguments[index];
                if (key is not ("--identity" or "--mode" or "--ready" or "--result") ||
                    string.IsNullOrWhiteSpace(arguments[index + 1]) || !values.TryAdd(key, arguments[index + 1]))
                    throw new ProbeException("invalid_arguments");
            }
            if (!values.TryGetValue("--mode", out var mode) || mode is not ("correct" or "wrong"))
                throw new ProbeException("invalid_mode");
            var identity = Path.GetFullPath(values["--identity"]);
            var ready = Path.GetFullPath(values["--ready"]);
            var result = Path.GetFullPath(values["--result"]);
            if (string.Equals(identity, ready, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(identity, result, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(ready, result, StringComparison.OrdinalIgnoreCase))
                throw new ProbeException("overlapping_contract_paths");
            return new(identity, mode, ready, result);
        }
    }

    private sealed class ProbeResult
    {
        public int SchemaVersion { get; } = 1;
        public string Mode { get; set; } = string.Empty;
        public bool ProcessIdentityConfirmed { get; set; }
        public bool? PasswordControlUnicode { get; set; }
        public bool? PasswordControlSupportsUnichar { get; set; }
        public uint? WindowsAnsiCodePage { get; set; }
        public bool ExpectedOutcomeObserved { get; set; }
        public bool CleanupConfirmed { get; set; }
        public bool Success { get; set; }
        public List<string> Errors { get; } = [];
    }

    private sealed class ProbeException(string code) : InvalidOperationException(code)
    {
        internal string Code { get; } = code;
    }

    private enum DialogKind { Authentication, InvalidPassword, ProjectProperties }

    private sealed class AuthProbe : IDisposable
    {
        private readonly Process process;
        private readonly long started;
        private readonly uint processId;
        private readonly Dictionary<nint, DialogKind> introduced = [];
        private readonly HashSet<nint> ownerChain = [];
        private readonly Stopwatch clock = Stopwatch.StartNew();
        private nint authentication;

        internal bool? PasswordControlUnicode { get; private set; }
        internal bool? PasswordControlSupportsUnichar { get; private set; }
        internal uint? WindowsAnsiCodePage { get; private set; }

        internal AuthProbe(string identityPath)
        {
            if (!OperatingSystem.IsWindows() || !Environment.Is64BitProcess)
                throw new ProbeException("unsupported_platform");
            using var identity = JsonDocument.Parse(File.ReadAllBytes(identityPath));
            if (!identity.RootElement.GetProperty("processId").TryGetUInt32(out processId) || processId == 0 || processId > int.MaxValue ||
                !identity.RootElement.GetProperty("startedAt").TryGetInt64(out started) || started <= 0)
                throw new ProbeException("invalid_process_identity");
            process = Process.GetProcessById((int)processId);
            try
            {
                _ = process.Handle;
                CheckProcess();
                using var current = Process.GetCurrentProcess();
                if (process.SessionId != current.SessionId ||
                    !string.Equals(Path.GetFileName(process.MainModule?.FileName), "EXCEL.EXE", StringComparison.OrdinalIgnoreCase))
                    throw new ProbeException("unexpected_process_or_session");
            }
            catch { process.Dispose(); throw; }
        }

        private void CheckProcess()
        {
            if (process.HasExited || process.StartTime.ToUniversalTime().ToFileTimeUtc() != started)
                throw new ProbeException("owned_process_identity_lost");
        }

        private void CheckWindow(nint window)
        {
            CheckProcess();
            if (!IsWindow(window) || GetWindowThreadProcessId(window, out var owner) == 0 || owner != processId)
                throw new ProbeException("owned_window_identity_lost");
        }

        private List<nint> Dialogs()
        {
            CheckProcess();
            var windows = new List<nint>();
            if (!EnumWindows((window, _) =>
                {
                    if (GetWindowThreadProcessId(window, out var owner) != 0 && owner == processId) windows.Add(window);
                    return true;
                }, 0)) throw new ProbeException("window_enumeration_failed");
            return windows.Where(window => IsWindow(window) && Class(window) == "#32770").ToList();
        }

        private List<nint> Children(nint parent)
        {
            CheckWindow(parent);
            var windows = new List<nint>();
            _ = EnumChildWindows(parent, (window, _) => { windows.Add(window); return true; }, 0);
            var children = new List<nint>();
            foreach (var window in windows)
            {
                CheckWindow(window);
                if (!IsChild(parent, window)) throw new ProbeException("child_window_identity_changed");
                if (IsWindowVisible(window)) children.Add(window);
            }
            return children;
        }

        internal void AssertInitiallyClear()
        {
            if (Dialogs().Count != 0) throw new ProbeException("preexisting_dialog");
        }

        private bool Matches(nint window, DialogKind kind)
        {
            CheckWindow(window);
            if (!IsWindowVisible(window) || Class(window) != "#32770") return false;
            var children = Children(window);
            var passwordEdits = children.Count(child => Class(child) == "Edit" &&
                (GetWindowLongPtr(child, -16).ToInt64() & 0x20) != 0);
            bool Button(int id) => children.Count(child => GetDlgCtrlID(child) == id && Class(child) == "Button") == 1;
            var tabs = children.Count(child => GetDlgCtrlID(child) == 12320 && Class(child) == "SysTabControl32");
            return kind switch
            {
                DialogKind.Authentication => passwordEdits == 1 && tabs == 0 && Button(1) && Button(2),
                DialogKind.InvalidPassword => Text(window) == "Project Locked" && Button(2) &&
                    children.Any(child => Class(child) == "Static" && Text(child) == "Invalid password") &&
                    children.Any(child => GetDlgCtrlID(child) == 2 && Class(child) == "Button" && Text(child).Replace("&", "", StringComparison.Ordinal) == "OK"),
                DialogKind.ProjectProperties => tabs == 1 && Button(1) && Button(2),
                _ => false
            };
        }

        private bool HasExpectedOwner(nint window)
        {
            if (window == authentication) return true;
            var parent = GetWindow(window, 4);
            for (var depth = 0; parent != 0 && depth < 32; parent = GetWindow(parent, 4), depth++)
                if (ownerChain.Contains(parent)) return true;
            return false;
        }

        private void Remember(nint window, DialogKind kind)
        {
            if (!Matches(window, kind) || !HasExpectedOwner(window)) throw new ProbeException("unexpected_dialog_signature_or_owner");
            if (introduced.TryGetValue(window, out var previous) && previous != kind)
                throw new ProbeException("dialog_handle_reused");
            introduced[window] = kind;
        }

        private nint Button(nint dialog, DialogKind kind, int id)
        {
            Remember(dialog, kind);
            var buttons = Children(dialog).Where(window => GetDlgCtrlID(window) == id && Class(window) == "Button").ToArray();
            if (buttons.Length != 1 || !IsWindowEnabled(buttons[0])) throw new ProbeException("expected_button_unavailable");
            return buttons[0];
        }

        private void Click(nint dialog, DialogKind kind, int buttonId)
        {
            var button = Button(dialog, kind, buttonId);
            CheckWindow(button);
            if (!PostMessage(button, 0xF5, 0, 0)) throw new ProbeException("expected_button_not_queued");
        }

        private void WaitClosed(nint window)
        {
            var wait = Stopwatch.StartNew();
            while (IsWindow(window) && wait.ElapsedMilliseconds < 1500) { CheckProcess(); Thread.Sleep(20); }
            if (IsWindow(window)) throw new ProbeException("expected_dialog_close_unconfirmed");
        }

        internal void Authenticate(string password, string mode)
        {
            var wait = Stopwatch.StartNew();
            while (wait.ElapsedMilliseconds < 15000)
            {
                var all = Dialogs().Where(IsWindowVisible).ToArray();
                var candidates = all.Where(window => Matches(window, DialogKind.Authentication)).ToArray();
                if (candidates.Length > 1 || all.Length > 1) throw new ProbeException("ambiguous_authentication_dialog");
                if (candidates.Length == 1) { authentication = candidates[0]; break; }
                if (all.Length != 0) throw new ProbeException("unexpected_dialog_before_authentication");
                Thread.Sleep(25);
            }
            if (authentication == 0) throw new ProbeException("authentication_dialog_not_observed");
            ownerChain.Add(authentication);
            var owner = GetWindow(authentication, 4);
            for (var depth = 0; owner != 0 && depth < 32; owner = GetWindow(owner, 4), depth++)
            {
                CheckWindow(owner);
                ownerChain.Add(owner);
            }
            Remember(authentication, DialogKind.Authentication);
            var edit = Children(authentication).Single(window => Class(window) == "Edit" &&
                (GetWindowLongPtr(window, -16).ToInt64() & 0x20) != 0);
            CheckWindow(edit);
            if (!IsChild(authentication, edit) || Class(edit) != "Edit" || !IsWindowEnabled(edit))
                throw new ProbeException("password_control_identity_unconfirmed");
            PasswordControlUnicode = IsWindowUnicode(edit);
            WindowsAnsiCodePage = GetACP();
            // UNICODE_NOCHAR queries support without inserting text.
            if (SendMessageTimeout(edit, 0x109, 0xFFFF, 0, 0x23, 500, out var supportsUnichar) != 0)
                PasswordControlSupportsUnichar = supportsUnichar != 0;
            foreach (var character in password)
            {
                Remember(authentication, DialogKind.Authentication);
                CheckWindow(edit);
                if (!IsChild(authentication, edit) || Class(edit) != "Edit" || !IsWindowEnabled(edit) ||
                    clock.ElapsedMilliseconds > 30000 ||
                    SendMessageTimeout(edit, 0x102, character, 1, 0x23, 500, out _) == 0)
                    throw new ProbeException("password_character_not_confirmed");
            }
            Click(authentication, DialogKind.Authentication, 1);
            wait.Restart();
            while (wait.ElapsedMilliseconds < 10000)
            {
                var all = Dialogs().Where(IsWindowVisible).ToArray();
                var wanted = mode == "wrong" ? DialogKind.InvalidPassword : DialogKind.ProjectProperties;
                var candidates = all.Where(window => window != authentication && Matches(window, wanted) && HasExpectedOwner(window)).ToArray();
                if (candidates.Length > 1) throw new ProbeException("ambiguous_authentication_outcome");
                var unexpectedKind = mode == "wrong" ? DialogKind.ProjectProperties : DialogKind.InvalidPassword;
                var unexpected = all.Where(window => window != authentication && Matches(window, unexpectedKind) && HasExpectedOwner(window)).ToArray();
                if (unexpected.Length != 0)
                {
                    foreach (var window in unexpected) Remember(window, unexpectedKind);
                    throw new ProbeException("unexpected_authentication_outcome");
                }
                if (candidates.Length == 1)
                {
                    var outcome = candidates[0];
                    Remember(outcome, wanted);
                    Click(outcome, wanted, 2);
                    WaitClosed(outcome);
                    if (mode == "wrong" && IsWindow(authentication))
                    {
                        Click(authentication, DialogKind.Authentication, 2);
                        WaitClosed(authentication);
                    }
                    return;
                }
                Thread.Sleep(25);
            }
            throw new ProbeException("authentication_outcome_not_observed");
        }

        internal bool Cleanup(List<string> errors)
        {
            var initialErrors = errors.Count;
            for (var attempt = 0; attempt < 4; attempt++)
            {
                List<nint> all;
                try { all = Dialogs(); }
                catch (Exception error) { errors.Add(ErrorCode(error)); return false; }
                if (all.Count == 0) return errors.Count == initialErrors;
                var acted = false;
                foreach (var kind in new[] { DialogKind.InvalidPassword, DialogKind.Authentication, DialogKind.ProjectProperties })
                    foreach (var window in all)
                    {
                        if (!IsWindow(window)) continue;
                        try
                        {
                            if (!HasExpectedOwner(window) || !Matches(window, kind)) continue;
                            Remember(window, kind);
                            Click(window, kind, 2);
                            WaitClosed(window);
                            acted = true;
                        }
                        catch (Exception error) { errors.Add(ErrorCode(error)); }
                    }
                if (!acted) { errors.Add("unexpected_or_unclosed_dialog_during_cleanup"); return false; }
            }
            errors.Add("dialog_cleanup_budget_exhausted");
            return false;
        }

        public void Dispose() => process.Dispose();
    }

    private static string Class(nint window)
    {
        var buffer = new char[128];
        var length = GetClassName(window, buffer, buffer.Length);
        if (length == 0) throw new ProbeException("window_class_unconfirmed");
        return new string(buffer, 0, length);
    }

    private static string Text(nint window)
    {
        var buffer = new char[256];
        var length = GetWindowText(window, buffer, buffer.Length);
        return new string(buffer, 0, length);
    }

    private delegate bool WindowCallback(nint window, nint data);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool EnumWindows(WindowCallback callback, nint data);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool EnumChildWindows(nint root, WindowCallback callback, nint data);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(nint window, out uint processId);
    [DllImport("user32.dll")] private static extern int GetDlgCtrlID(nint window);
    [DllImport("user32.dll")] private static extern nint GetWindow(nint window, uint command);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsWindow(nint window);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsWindowVisible(nint window);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsWindowEnabled(nint window);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsWindowUnicode(nint window);
    [DllImport("kernel32.dll")] private static extern uint GetACP();
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsChild(nint root, nint window);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool PostMessage(nint window, uint message, nuint wParam, nint lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(nint window, [Out] char[] text, int capacity);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(nint window, [Out] char[] text, int capacity);
    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] private static extern nint GetWindowLongPtr(nint window, int index);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern nint SendMessageTimeout(nint window, uint message, nuint wParam, nint lParam, uint flags, uint milliseconds, out nuint result);
}
