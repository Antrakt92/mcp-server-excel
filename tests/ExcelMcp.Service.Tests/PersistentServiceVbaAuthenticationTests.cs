using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;
using Vbe = Microsoft.Vbe.Interop;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed class PersistentServiceVbaAuthenticationFixture : PersistentServiceWorkbookFixture
{
    public PersistentServiceVbaAuthenticationFixture() : base("KnownPasswordAuthentication.xlsm") { }
}

[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "VBA")]
public sealed class PersistentServiceVbaAuthenticationTests : PersistentServiceWorkbookTestBase,
    IClassFixture<PersistentServiceVbaAuthenticationFixture>
{
    private readonly PersistentServiceVbaAuthenticationFixture fixture;
    private readonly IPersistentVbaCommands commands;

    public PersistentServiceVbaAuthenticationTests(PersistentServiceVbaAuthenticationFixture fixture) : base(fixture)
    {
        this.fixture = fixture;
        commands = fixture.CreateCommands<IPersistentVbaCommands>();
    }

    [Fact]
    public async Task KnownPassword_RequiresARealChallenge_AndPreservesSavedProtection()
    {
        const string password = " A\"b\\c Exact!7 ";
        const string siblingPassword = "SiblingOnly!9";
        const string originalCode = "Public Function Marker() As Long\nMarker = 73\nEnd Function";
        const string replacementCode = "Public Function Marker() As Long\nMarker = 84\nEnd Function";
        const string addedCode = "Public Function AddedMarker() As Long\nAddedMarker = 92\nEnd Function";
        RequireSuccess(commands.Import(fixture.BatchToken, "AuthenticationMarker", originalCode));
        string sourceHash = RequireSuccess(commands.View(fixture.BatchToken, "AuthenticationMarker")).SourceHash;
        var projectName = RequireSuccess(commands.Status(fixture.BatchToken)).ProjectName;
        Assert.NotNull(projectName);
        string siblingName = "Sibling" + Guid.NewGuid().ToString("N")[..16];
        string siblingPath = Path.Combine(Path.GetDirectoryName(fixture.WorkbookPath)!, "AuthenticationSibling.xlsm");
        await CreateProtectedSiblingCopyAsync(siblingPath, siblingName, siblingPassword, projectName);
        byte[] siblingBytes = ReadSavedHash(siblingPath);
        var configured = RequireSuccess(commands.ProtectProject(fixture.BatchToken, password));
        Assert.Equal(projectName, configured.ProjectName);
        Assert.True(configured.ConfigurationApplied);
        await fixture.SaveAndReopenAsync();
        Assert.Equal("Locked", RequireSuccess(commands.Status(fixture.BatchToken)).Protection);
        var lockedView = await fixture.SendForFailureAsync("vba.view", new { moduleName = "AuthenticationMarker" });
        Assert.Equal("Permissions", lockedView.ErrorCategory);
        Assert.Contains("unlock-project", lockedView.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("batch JSON on stdin", lockedView.ErrorMessage, StringComparison.Ordinal);
        Assert.DoesNotContain("does not unlock projects", lockedView.ErrorMessage, StringComparison.Ordinal);
        AssertTargetCells();
        byte[] savedBytes = ReadSavedHash(fixture.WorkbookPath);
        string editedHash = string.Empty;

        await WithSiblingAsync(siblingPath, siblingName, async () =>
        {
            var before = PinSiblingAndReadView(siblingPath, siblingName, visible: false);
            var excessive = await fixture.SendForFailureAsync("vba.verify-password",
                new { password = password + new string('x', 100_000) });
            Assert.Equal("InvalidInput", excessive.ErrorCategory);
            Assert.Contains("supported length", excessive.ErrorMessage, StringComparison.Ordinal);
            AssertFailureState(excessive, performed: false, verified: null, accessible: false, protection: "Locked");
            Assert.Equal(before, ReadView(siblingPath, siblingName));
            Assert.Equal("Locked", RequireSuccess(commands.Status(fixture.BatchToken)).Protection);
            AssertTargetCells();

            string[] wrongPasswords =
            [
                password.Replace("Exact", "Other", StringComparison.Ordinal), password.Trim(),
                password.Replace("Exact", "eXact", StringComparison.Ordinal), password.Replace('\\', '/'),
                password.Replace("\"", "'", StringComparison.Ordinal), siblingPassword
            ];
            for (int index = 0; index < wrongPasswords.Length; index++)
            {
                string wrong = wrongPasswords[index];
                Assert.NotEqual(password, wrong);
                AssertPreservedAuthentication(siblingPath, siblingName, projectName, "verify-password",
                    wrong, verified: false, visible: index % 2 == 0);
            }
            before = PinSiblingAndReadView(siblingPath, siblingName, visible: true);
            var refused = await fixture.SendForFailureAsync("vba.unlock-project", new { password = wrongPasswords[0] });
            Assert.Equal("Permissions", refused.ErrorCategory);
            AssertFailureState(refused, performed: true, verified: false, accessible: false, protection: "Locked");
            Assert.Equal(before, ReadView(siblingPath, siblingName));
            Assert.Equal("Locked", RequireSuccess(commands.Status(fixture.BatchToken)).Protection);

            AssertPreservedAuthentication(siblingPath, siblingName, projectName, "verify-password",
                password, verified: true, visible: false);
            Assert.Equal(sourceHash, RequireSuccess(commands.View(fixture.BatchToken, "AuthenticationMarker")).SourceHash);
            Assert.Equal(savedBytes, ReadSavedHash(fixture.WorkbookPath));

            before = PinSiblingAndReadView(siblingPath, siblingName, visible: true);
            var inaccessible = await fixture.SendForFailureAsync("vba.verify-password", new { password = siblingPassword });
            Assert.Equal("Prerequisite", inaccessible.ErrorCategory);
            AssertFailureState(inaccessible, performed: false, verified: null, accessible: true, protection: "None");
            Assert.Equal(before, ReadView(siblingPath, siblingName));
            var accessible = RequireSuccess(commands.UnlockProject(fixture.BatchToken, siblingPassword));
            AssertAuthentication(accessible, projectName, "unlock-project", verified: null, alreadyAccessible: true);
            Assert.Equal(before, ReadView(siblingPath, siblingName));
            AssertTargetCells();

            var replaced = RequireSuccess(commands.ReplaceModule(fixture.BatchToken, "AuthenticationMarker", sourceHash, replacementCode));
            var edited = RequireSuccess(commands.View(fixture.BatchToken, "AuthenticationMarker"));
            Assert.Equal("AuthenticationMarker", replaced.ModuleName);
            Assert.Equal(edited.SourceHash, replaced.SourceHash);
            Assert.Equal(edited.LineCount, replaced.LineCount);
            Assert.NotEqual(sourceHash, edited.SourceHash);
            Assert.Contains("Marker = 84", edited.Code, StringComparison.Ordinal);
            editedHash = edited.SourceHash;
            RequireSuccess(commands.Import(fixture.BatchToken, "AuthenticationAdded", addedCode));
            Assert.Contains("AddedMarker = 92", RequireSuccess(commands.View(fixture.BatchToken, "AuthenticationAdded")).Code, StringComparison.Ordinal);
            Assert.False(fixture.ExecuteRawVerification((context, _) => context.Book.Saved));
            AssertTargetCells();
            Assert.Equal(siblingBytes, ReadSavedHash(siblingPath));
        });

        // The sibling is closed before closing/reopening the target's owned Excel session.
        await fixture.SaveAndReopenAsync();
        Assert.Equal("Locked", RequireSuccess(commands.Status(fixture.BatchToken)).Protection);
        AssertTargetCells();
        await fixture.ReopenAsync(readOnly: true, save: false);
        try
        {
            var audit = await fixture.SendForFailureAsync("vba.unlock-project", new { password });
            Assert.Equal("Permissions", audit.ErrorCategory);
            Assert.Contains("Audit session", audit.ErrorMessage, StringComparison.Ordinal);
            Assert.Equal("Locked", RequireSuccess(commands.Status(fixture.BatchToken)).Protection);
            AssertTargetCells();
        }
        finally { await fixture.ReopenAsync(readOnly: false, save: false); }

        savedBytes = ReadSavedHash(fixture.WorkbookPath);
        var reopened = RequireSuccess(commands.Status(fixture.BatchToken));
        Assert.Equal(projectName, reopened.ProjectName);
        Assert.Equal("Locked", reopened.Protection);
        await WithSiblingAsync(siblingPath, siblingName, () =>
        {
            AssertPreservedAuthentication(siblingPath, siblingName, projectName, "verify-password",
                siblingPassword, verified: false, visible: false);
            AssertPreservedAuthentication(siblingPath, siblingName, projectName, "unlock-project",
                password, verified: true, visible: true);
            Assert.Equal(savedBytes, ReadSavedHash(fixture.WorkbookPath));
            Assert.Equal(editedHash, RequireSuccess(commands.View(fixture.BatchToken, "AuthenticationMarker")).SourceHash);
            Assert.Contains("AddedMarker = 92", RequireSuccess(commands.View(fixture.BatchToken, "AuthenticationAdded")).Code, StringComparison.Ordinal);
            AssertTargetCells();
            Assert.Equal(siblingBytes, ReadSavedHash(siblingPath));
            return Task.CompletedTask;
        });
        await fixture.SaveAndReopenAsync();
        Assert.Equal("Locked", RequireSuccess(commands.Status(fixture.BatchToken)).Protection);
        AssertTargetCells();
        Assert.Equal(siblingBytes, ReadSavedHash(siblingPath));
    }

    private void AssertPreservedAuthentication(string siblingPath, string siblingName, string projectName,
        string action, string password, bool verified, bool visible)
    {
        var before = PinSiblingAndReadView(siblingPath, siblingName, visible);
        Assert.True(before.TargetSaved);
        var result = RequireSuccess(action == "verify-password"
            ? commands.VerifyPassword(fixture.BatchToken, password)
            : commands.UnlockProject(fixture.BatchToken, password));
        AssertAuthentication(result, projectName, action, verified, alreadyAccessible: false);
        Assert.Equal(before, ReadView(siblingPath, siblingName));
        var status = RequireSuccess(commands.Status(fixture.BatchToken));
        Assert.Equal(projectName, status.ProjectName);
        Assert.Equal(verified ? "None" : "Locked", status.Protection);
        AssertTargetCells();
    }

    private void AssertAuthentication(VbaProjectAuthenticationResult result, string projectName,
        string action, bool? verified, bool alreadyAccessible)
    {
        Assert.Equal(fixture.WorkbookPath, result.FilePath, ignoreCase: true);
        Assert.Equal(projectName, result.ProjectName);
        Assert.Equal(action, result.Action);
        Assert.Equal(!alreadyAccessible, result.AuthenticationPerformed);
        Assert.Equal(alreadyAccessible, result.AlreadyAccessible);
        Assert.Equal(verified, result.PasswordVerified);
        Assert.Equal(verified is true, result.AccessUnlocked);
        Assert.Equal(alreadyAccessible ? "None" : "Locked", result.ProtectionBefore);
        Assert.Equal(verified is false ? "Locked" : "None", result.ProtectionAfter);
        Assert.False(result.ProtectionSettingsChanged);
        Assert.False(result.Saved);
        Assert.False(result.Reopened);
    }

    private static void AssertFailureState(ServiceResponse response, bool performed, bool? verified,
        bool accessible, string protection)
    {
        using var state = JsonDocument.Parse(response.PartialState!);
        var root = state.RootElement;
        Assert.NotEmpty(root.GetProperty("phase").GetString()!);
        Assert.Equal(!accessible, root.GetProperty("challengeObserved").GetBoolean());
        Assert.Equal(performed, root.GetProperty("authenticationPerformed").GetBoolean());
        var outcome = root.GetProperty("passwordVerified");
        if (verified.HasValue) Assert.Equal(verified.Value, outcome.GetBoolean());
        else Assert.Equal(JsonValueKind.Null, outcome.ValueKind);
        Assert.False(root.GetProperty("accessUnlocked").GetBoolean());
        Assert.Equal(accessible, root.GetProperty("alreadyAccessible").GetBoolean());
        Assert.Equal(protection, root.GetProperty("currentProtection").GetString());
        Assert.True(root.GetProperty("dialogCleanupConfirmed").GetBoolean());
        Assert.True(root.GetProperty("activeProjectRestored").GetBoolean());
        Assert.True(root.GetProperty("visibilityRestored").GetBoolean());
        foreach (string name in new[] { "protectionSettingsChanged", "saved", "reopened", "rollbackAttempted" })
            Assert.False(root.GetProperty(name).GetBoolean());
    }

    private static byte[] ReadSavedHash(string path)
    {
        using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return SHA256.HashData(file);
    }

    private void AssertTargetCells()
    {
        var values = RequireSuccess(_commands.GetValues(fixture.BatchToken, "Sheet1", "A1:B2")).Values;
        Assert.Equal(2, values.Count);
        Assert.Equal("Target fixture", values[0][0]);
        Assert.Equal(73, Convert.ToInt32(values[0][1], CultureInfo.InvariantCulture));
        Assert.Equal("Preserve cells", values[1][0]);
        Assert.Equal(19.25, Convert.ToDouble(values[1][1], CultureInfo.InvariantCulture));
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Preserves synthetic sibling setup or SaveCopyAs failure and independently restores target name and cells; any partial copy stays in the fixture-owned directory for normal cleanup.")]
    private async Task CreateProtectedSiblingCopyAsync(string path, string siblingName, string password, string targetName)
    {
        Exception? failure = null;
        try
        {
            SetProjectName(siblingName);
            RequireSuccess(_commands.SetValues(fixture.BatchToken, "Sheet1", "A1:B2",
                [["Sibling fixture", 11], ["Preserve sibling", 6.5]]));
            var configured = RequireSuccess(commands.ProtectProject(fixture.BatchToken, password));
            Assert.Equal(siblingName, configured.ProjectName);
            Assert.True(configured.ConfigurationApplied);
            var saved = await fixture.SendAsync("session.save", new { });
            Assert.True(saved.Success, saved.ErrorMessage);
            // SaveCopyAs does not open a workbook; a failed/partial copy cannot
            // become a live sibling outside WithSiblingAsync's exact cleanup.
            fixture.ExecuteRawVerification((context, _) => context.Book.SaveCopyAs(path));
        }
        catch (Exception error) { failure = error; }
        finally
        {
            try { SetProjectName(targetName); }
            catch (Exception cleanup)
            {
                failure = failure is null ? cleanup : new AggregateException("Sibling setup and target-name restoration failed.", failure, cleanup);
            }
            try
            {
                RequireSuccess(_commands.SetValues(fixture.BatchToken, "Sheet1", "A1:B2",
                    [["Target fixture", 73], ["Preserve cells", 19.25]], overwritePolicy: OverwritePolicy.Allow));
            }
            catch (Exception cleanup)
            {
                failure = failure is null ? cleanup : new AggregateException("Sibling setup and target-cell restoration failed.", failure, cleanup);
            }
        }
        if (failure is not null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
    }

    private void SetProjectName(string name) => fixture.ExecuteRawVerification((context, _) =>
    {
        Vbe.VBProject? project = null;
        try
        {
            project = context.Book.VBProject;
            project.Name = name;
            Assert.Equal(name, project.Name);
        }
        finally { ComUtilities.Release(ref project); }
    });

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Preserves the synthetic authentication/assertion failure and independently closes the exact same-application sibling before the fixture can close its target.")]
    private async Task WithSiblingAsync(string path, string projectName, Func<Task> action)
    {
        string? name = null;
        Exception? failure = null;
        try
        {
            name = OpenSibling(path, projectName);
            await action();
        }
        catch (Exception error) { failure = error; }
        finally
        {
            if (name is not null)
            {
                try { CloseSibling(name, path); }
                catch (Exception cleanup)
                {
                    failure = failure is null ? cleanup : new AggregateException("Authentication and exact sibling cleanup failed.", failure, cleanup);
                }
            }
        }
        if (failure is not null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "A failed synthetic sibling setup closes that exact workbook independently and preserves setup and cleanup causes.")]
    private string OpenSibling(string path, string projectName) => fixture.ExecuteRawVerification((context, _) =>
    {
        Excel.Workbooks? books = null;
        Excel.Workbook? sibling = null;
        Excel.Application? application = null;
        Vbe.VBProject? project = null;
        try
        {
            books = context.App.Workbooks;
            sibling = books.Open(path, UpdateLinks: 0, ReadOnly: false);
            application = sibling.Application;
            Assert.Equal(context.App.Hwnd, application.Hwnd);
            Assert.Equal(path, sibling.FullName, ignoreCase: true);
            Assert.False(sibling.ReadOnly);
            project = sibling.VBProject;
            Assert.Equal(projectName, project.Name);
            Assert.Equal(Vbe.vbext_ProjectProtection.vbext_pp_locked, project.Protection);
            Assert.True(sibling.Saved);
            return sibling.Name;
        }
        catch (Exception primary)
        {
            try { sibling?.Close(SaveChanges: false); }
            catch (Exception cleanup) { throw new AggregateException("Sibling setup and exact cleanup failed.", primary, cleanup); }
            throw;
        }
        finally
        {
            ComUtilities.Release(ref project);
            ComUtilities.Release(ref application);
            ComUtilities.Release(ref sibling);
            ComUtilities.Release(ref books);
        }
    });

    private AuthenticationView PinSiblingAndReadView(string path, string name, bool visible)
    {
        fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Workbooks? books = null;
            Excel.Workbook? sibling = null;
            Vbe.VBProject? project = null;
            Vbe.VBE? editor = null;
            Vbe.Window? window = null;
            try
            {
                books = context.App.Workbooks;
                sibling = books[Path.GetFileName(path)];
                Assert.Equal(path, sibling.FullName, ignoreCase: true);
                project = sibling.VBProject;
                Assert.Equal(name, project.Name);
                editor = project.VBE;
                window = editor.MainWindow;
                window.Visible = visible;
                editor.ActiveVBProject = project;
            }
            finally
            {
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref editor);
                ComUtilities.Release(ref project);
                ComUtilities.Release(ref sibling);
                ComUtilities.Release(ref books);
            }
        });
        var view = ReadView(path, name);
        Assert.Equal(name, view.ActiveProject);
        Assert.Equal(path, view.ActiveProjectFile, ignoreCase: true);
        Assert.Equal(visible, view.EditorVisible);
        return view;
    }

    private AuthenticationView ReadView(string path, string name) => fixture.ExecuteRawVerification((context, _) =>
    {
        Excel.Workbooks? books = null;
        Excel.Workbook? sibling = null;
        Excel.Windows? targetWindows = null;
        Excel.Window? targetWindow = null;
        Excel.Worksheet? targetSheet = null;
        Excel.Range? selection = null;
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;
        Excel.Range? cells = null;
        Vbe.VBProject? project = null;
        Vbe.VBE? editor = null;
        Vbe.Window? editorWindow = null;
        Vbe.VBProject? active = null;
        try
        {
            books = context.App.Workbooks;
            sibling = books[Path.GetFileName(path)];
            Assert.Equal(path, sibling.FullName, ignoreCase: true);
            project = sibling.VBProject;
            Assert.Equal(name, project.Name);
            Assert.Equal(Vbe.vbext_ProjectProtection.vbext_pp_locked, project.Protection);
            Assert.True(sibling.Saved);
            sheets = sibling.Worksheets;
            sheet = (Excel.Worksheet)sheets[1];
            cells = sheet.Range["A1:B2"];
            var values = Assert.IsType<object[,]>((object)cells.Value2);
            Assert.Equal("Sibling fixture", values[1, 1]);
            Assert.Equal(11, Convert.ToInt32(values[1, 2], CultureInfo.InvariantCulture));
            Assert.Equal("Preserve sibling", values[2, 1]);
            Assert.Equal(6.5, Convert.ToDouble(values[2, 2], CultureInfo.InvariantCulture));
            editor = project.VBE;
            editorWindow = editor.MainWindow;
            active = editor.ActiveVBProject;
            targetWindows = context.Book.Windows;
            targetWindow = targetWindows[1];
            targetSheet = (Excel.Worksheet)context.Book.ActiveSheet;
            selection = targetWindow.RangeSelection;
            return new AuthenticationView(active.Name, active.FileName, editorWindow.Visible,
                context.App.Visible, context.App.WindowState, Convert.ToInt32((object)targetWindow.Zoom, CultureInfo.InvariantCulture),
                targetWindow.ScrollRow, targetWindow.ScrollColumn, selection.Address, context.Book.Saved, sibling.Saved,
                targetSheet.Name, targetWindow.DisplayGridlines, targetWindow.DisplayHeadings);
        }
        finally
        {
            ComUtilities.Release(ref selection);
            ComUtilities.Release(ref targetSheet);
            ComUtilities.Release(ref targetWindow);
            ComUtilities.Release(ref targetWindows);
            ComUtilities.Release(ref active);
            ComUtilities.Release(ref editorWindow);
            ComUtilities.Release(ref editor);
            ComUtilities.Release(ref cells);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref sheets);
            ComUtilities.Release(ref project);
            ComUtilities.Release(ref sibling);
            ComUtilities.Release(ref books);
        }
    });

    private void CloseSibling(string name, string path) => fixture.ExecuteRawVerification((context, _) =>
    {
        Excel.Workbooks? books = null;
        Excel.Workbook? sibling = null;
        try
        {
            books = context.App.Workbooks;
            sibling = books[name];
            Assert.Equal(path, sibling.FullName, ignoreCase: true);
            sibling.Close(SaveChanges: false);
        }
        finally
        {
            ComUtilities.Release(ref sibling);
            ComUtilities.Release(ref books);
        }
    });

    private sealed record AuthenticationView(string ActiveProject, string ActiveProjectFile, bool EditorVisible,
        bool ApplicationVisible, Excel.XlWindowState WindowState, int Zoom, int ScrollRow, int ScrollColumn,
        string Selection, bool TargetSaved, bool SiblingSaved, string TargetSheet, bool Gridlines, bool Headings);
}
