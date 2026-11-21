using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Screenshot")]
public sealed class PersistentServiceScreenshotRangeAdmissionTests(PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public async Task CaptureRange_DisjointAreas_FailsWithoutChangingSavedView()
    {
        const string targetSheet = "Sheet1";
        const string activeSheet = "UnionCaptureOtherSheet";
        const string unionAddress = "A1:B2,D4:E5";
        const string definedName = "ScreenshotDisjointSelection";
        _fixture.CreateNamedTestSheet(_fixture.BatchToken, activeSheet);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? union = null;
            Excel.Names? names = null;
            Excel.Name? created = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, targetSheet);
                Assert.NotNull(sheet);
                union = sheet.Range[unionAddress];
                names = context.Book.Names;
                created = names.Add(Name: definedName, RefersTo: union);
            }
            finally
            {
                ComUtilities.Release(ref created);
                ComUtilities.Release(ref names);
                ComUtilities.Release(ref union);
                ComUtilities.Release(ref sheet);
            }
        });
        _fixture.RegisterNamedRangeForCleanup(definedName);

        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? selection = null;
            Excel.Window? window = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, activeSheet);
                Assert.NotNull(sheet);
                sheet.Activate();
                selection = sheet.Range["K12"];
                selection.Select();
                window = context.App.ActiveWindow;
                window.Zoom = 110;
            }
            finally
            {
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref selection);
                ComUtilities.Release(ref sheet);
            }
        });
        AssertResolvesToDisjointAreas(targetSheet, unionAddress);
        AssertResolvesToDisjointAreas(targetSheet, definedName);

        var saved = _fixture.Send("session.save", new { });
        Assert.True(saved.Success, saved.ErrorMessage);
        var before = ReadHiddenWorkbookView();
        Assert.Equal(activeSheet, before.Sheet);
        Assert.Equal("$K$12", before.Selection);
        Assert.Equal(110, before.Zoom);
        Assert.True(before.Saved);
        Assert.False(before.Visible);

        foreach (string rangeAddress in new[] { unionAddress, definedName })
        {
            var failure = await _fixture.SendForFailureAsync("screenshot.capture", new { sheetName = targetSheet, rangeAddress });

            Assert.False(failure.Success);
            Assert.Equal("InvalidInput", failure.ErrorCategory);
            Assert.Contains("one rectangular range", failure.ErrorMessage, StringComparison.Ordinal);
            Assert.Equal(before, ReadHiddenWorkbookView());
        }
    }

    private void AssertResolvesToDisjointAreas(string sheetName, string rangeAddress) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Areas? areas = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                range = sheet.Range[rangeAddress];
                areas = range.Areas;
                Assert.Equal(2, areas.Count);
            }
            finally
            {
                ComUtilities.Release(ref areas);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });

    private (string Sheet, string Selection, int Zoom, int ScrollRow, int ScrollColumn, bool Saved, bool Visible) ReadHiddenWorkbookView() =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? selection = null;
            Excel.Window? window = null;
            try
            {
                sheet = (Excel.Worksheet)context.Book.ActiveSheet;
                selection = (Excel.Range)context.App.Selection;
                window = context.App.ActiveWindow;
                return (sheet.Name, selection.Address, Convert.ToInt32((object)window.Zoom, CultureInfo.InvariantCulture),
                    window.ScrollRow, window.ScrollColumn, context.Book.Saved, context.App.Visible);
            }
            finally
            {
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref selection);
                ComUtilities.Release(ref sheet);
            }
        });
}
