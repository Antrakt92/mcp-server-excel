using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;
using Vbe = Microsoft.Vbe.Interop;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed class PersistentServiceVbaProtectionFixture : PersistentServiceWorkbookFixture
{
    public PersistentServiceVbaProtectionFixture() : base("ProjectProtection.xlsm") { }
}

[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "VBA")]
public sealed class PersistentServiceVbaProtectionTests : PersistentServiceWorkbookTestBase,
    IClassFixture<PersistentServiceVbaProtectionFixture>
{
    private readonly PersistentServiceVbaProtectionFixture fixture;
    private readonly IPersistentVbaCommands commands;

    public PersistentServiceVbaProtectionTests(PersistentServiceVbaProtectionFixture fixture) : base(fixture)
    {
        this.fixture = fixture;
        commands = fixture.CreateCommands<IPersistentVbaCommands>();
    }

    [Fact]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Preserves the protection/assertion failure and independent exact-sibling cleanup failure, then rethrows both.")]
    public async Task ProtectProject_SaveAndFreshOpen_LocksTheSelectedProject()
    {
        const string password = "SyntheticProtection!7";
        var emptyProject = await fixture.SendForFailureAsync("vba.protect-project", new { password });
        Assert.Equal("Prerequisite", emptyProject.ErrorCategory);
        Assert.Contains("VBA code", emptyProject.ErrorMessage, StringComparison.Ordinal);
        RequireSuccess(commands.Import(fixture.BatchToken, "ProtectionMarker", "Public Function Marker() As Long\nMarker = 73\nEnd Function"));
        var before = RequireSuccess(commands.Status(fixture.BatchToken));
        Assert.Equal("None", before.Protection);
        var sourceHash = RequireSuccess(commands.View(fixture.BatchToken, "ProtectionMarker")).SourceHash;

        var excessive = await fixture.SendForFailureAsync("vba.protect-project", new { password = new string('x', 100_000) });
        Assert.True(excessive.ErrorMessage?.Contains("supported length", StringComparison.Ordinal) == true,
            excessive.ErrorMessage + Environment.NewLine + excessive.InnerError + Environment.NewLine + excessive.PartialState);
        using (var partial = JsonDocument.Parse(excessive.PartialState!))
        {
            Assert.False(partial.RootElement.GetProperty("acceptanceRequested").GetBoolean());
            Assert.False(partial.RootElement.GetProperty("configurationApplied").GetBoolean());
            Assert.True(partial.RootElement.GetProperty("dialogCleanupConfirmed").GetBoolean());
            Assert.True(partial.RootElement.GetProperty("activeProjectRestored").GetBoolean());
            Assert.True(partial.RootElement.GetProperty("visibilityRestored").GetBoolean());
        }
        Assert.Equal("None", RequireSuccess(commands.Status(fixture.BatchToken)).Protection);
        Assert.Equal(sourceHash, RequireSuccess(commands.View(fixture.BatchToken, "ProtectionMarker")).SourceHash);

        var siblingPath = Path.Combine(Path.GetDirectoryName(fixture.WorkbookPath)!, $"ProtectionSibling-{Guid.NewGuid():N}.xlsm");
        var siblingProjectName = "Sibling" + Guid.NewGuid().ToString("N")[..16];
        string? siblingName = null;
        Exception? protectionFailure = null;
        try
        {
            var sibling = CreateSelectedSibling(siblingPath, siblingProjectName);
            siblingName = sibling.Name;
            Assert.NotEqual(before.ProjectName, siblingProjectName);

            var result = RequireSuccess(commands.ProtectProject(fixture.BatchToken, password));
            Assert.Equal(before.ProjectName, result.ProjectName);
            Assert.Equal("None", result.CurrentProtection);
            Assert.True(result.ConfigurationApplied);
            Assert.True(result.SaveRequired);
            Assert.True(result.ReopenRequired);
            Assert.DoesNotContain(password, JsonSerializer.Serialize(result), StringComparison.Ordinal);
            fixture.ExecuteRawVerification((context, _) =>
            {
                Excel.Workbooks? books = null;
                Excel.Workbook? other = null;
                Vbe.VBProject? project = null;
                Vbe.VBE? editor = null;
                Vbe.Window? window = null;
                Vbe.VBProject? active = null;
                try
                {
                    books = context.App.Workbooks;
                    other = books[sibling.Name];
                    Assert.Equal(siblingPath, other.FullName, ignoreCase: true);
                    project = other.VBProject;
                    editor = project.VBE;
                    window = editor.MainWindow;
                    active = editor.ActiveVBProject;
                    Assert.Equal(siblingProjectName, project.Name);
                    Assert.Equal(Vbe.vbext_ProjectProtection.vbext_pp_none, project.Protection);
                    Assert.Equal(siblingProjectName, active.Name);
                    Assert.Equal(siblingPath, active.FileName, ignoreCase: true);
                    Assert.Equal(sibling.Visible, window.Visible);
                }
                finally
                {
                    ComUtilities.Release(ref active);
                    ComUtilities.Release(ref window);
                    ComUtilities.Release(ref editor);
                    ComUtilities.Release(ref project);
                    ComUtilities.Release(ref other);
                    ComUtilities.Release(ref books);
                }
            });
        }
        catch (Exception error) { protectionFailure = error; }
        finally
        {
            if (siblingName is not null)
            {
                try { CloseSibling(siblingName, siblingPath); }
                catch (Exception cleanup)
                {
                    protectionFailure = protectionFailure is null ? cleanup :
                        new AggregateException("Protection verification and exact sibling cleanup failed.", protectionFailure, cleanup);
                }
            }
        }
        if (protectionFailure is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(protectionFailure).Throw();

        await fixture.SaveAndReopenAsync();
        var persisted = RequireSuccess(commands.Status(fixture.BatchToken));
        Assert.Equal(before.ProjectName, persisted.ProjectName);
        Assert.Equal("Locked", persisted.Protection);

        await fixture.ReopenAsync(readOnly: true, save: false);
        try
        {
            var readOnly = await fixture.SendForFailureAsync("vba.protect-project", new { password });
            Assert.Equal("Permissions", readOnly.ErrorCategory);
            Assert.Contains("Audit session", readOnly.ErrorMessage, StringComparison.Ordinal);
            Assert.Equal("Locked", RequireSuccess(commands.Status(fixture.BatchToken)).Protection);
        }
        finally
        {
            await fixture.ReopenAsync(readOnly: false, save: false);
        }

        var refused = await fixture.SendForFailureAsync("vba.protect-project", new { password = "DifferentSynthetic!9" });
        Assert.Equal("Permissions", refused.ErrorCategory);
        Assert.Contains("locked", refused.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("DifferentSynthetic!9", refused.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal("Locked", RequireSuccess(commands.Status(fixture.BatchToken)).Protection);
    }

    private (string Name, bool Visible) CreateSelectedSibling(string path, string projectName) =>
        fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Workbooks? books = null;
            Excel.Workbook? sibling = null;
            Vbe.VBProject? project = null;
            Vbe.VBE? editor = null;
            Vbe.Window? window = null;
            Vbe.VBProject? active = null;
            Excel.Application? siblingApplication = null;
            var stage = "saving-owned-fixture-copy";
            var diagnostic = string.Empty;
            try
            {
                // Both projects need real VBA content; use only the owned marker
                // fixture, with opening macros still disabled by the session.
                context.Book.SaveCopyAs(path);
                books = context.App.Workbooks;
                sibling = books.Open(path, UpdateLinks: 0, ReadOnly: false);
                siblingApplication = sibling.Application;
                var sameApplication = siblingApplication.Hwnd == context.App.Hwnd;
                diagnostic = $"sameApplication={sameApplication}; targetHasVba={context.Book.HasVBProject}; siblingHasVba={sibling.HasVBProject}; siblingReadOnly={sibling.ReadOnly}; automationSecurity={(int)context.App.AutomationSecurity}";
                Assert.True(sameApplication, "The synthetic sibling opened outside the owned Excel application.");
                stage = "reading-sibling-project";
                project = sibling.VBProject;
                stage = "selecting-sibling-project";
                project.Name = projectName;
                sibling.Save();
                editor = project.VBE;
                window = editor.MainWindow;
                editor.ActiveVBProject = project;
                active = editor.ActiveVBProject;
                Assert.Equal(projectName, active.Name);
                return (sibling.Name, window.Visible);
            }
            catch (Exception primary)
            {
                try { sibling?.Close(SaveChanges: false); }
                catch (Exception cleanup)
                {
                    throw new AggregateException("Sibling setup and exact workbook cleanup failed.", primary, cleanup);
                }
                throw new InvalidOperationException($"Synthetic sibling setup failed during {stage}; {diagnostic}.", primary);
            }
            finally
            {
                ComUtilities.Release(ref active);
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref editor);
                ComUtilities.Release(ref project);
                ComUtilities.Release(ref siblingApplication);
                ComUtilities.Release(ref sibling);
                ComUtilities.Release(ref books);
            }
        });

    private void CloseSibling(string name, string path) =>
        fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Workbooks? books = null;
            Excel.Workbook? sibling = null;
            try
            {
                books = context.App.Workbooks;
                sibling = books[name];
                Assert.True(string.Equals(path, sibling.FullName, StringComparison.OrdinalIgnoreCase),
                    "Refusing to close a workbook whose saved identity does not match the synthetic sibling.");
                sibling.Close(SaveChanges: false);
            }
            finally
            {
                ComUtilities.Release(ref sibling);
                ComUtilities.Release(ref books);
            }
        });
}
