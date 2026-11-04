using System.Globalization;
using System.Runtime.ExceptionServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class IsolatedServiceScreenshotTests
{
    [Fact]
    public void CaptureRange_NonActiveFrozenTarget_PreservesExactBandsBothViewsAndCells()
    {
        string other = "Other" + Guid.NewGuid().ToString("N")[..16];
        var created = _fixture.Send("sheet.create", new { sheetName = other });
        Assert.True(created.Success, created.ErrorMessage);
        _fixture.RegisterSheetForCleanup(other);
        var populated = _fixture.Send("range.set-values", new
        {
            sheetName = other,
            rangeAddress = "A1:B2",
            values = new object[][] { ["Other sheet", 17d], ["Keep unchanged", 29d] }
        });
        Assert.True(populated.Success, populated.ErrorMessage);

        // This reuses every interior-row pixel assertion from the active-sheet
        // geometry suite, rather than settling for marker presence.
        AssertGeometryCapture(false, false, "D17:M26", "D17:M21", "D22:M26", "$D$17:$M$26",
            "frozen-columns", initiallyActiveOtherSheet: other);
        var final = ReadGeometryView();
        Assert.Equal(other, final.View.Sheet);
        Assert.True(final.View.Saved, final.Diagnostic);
    }

    private void PrepareInactiveGeometryTargetPose() =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? selection = null;
            Excel.Window? window = null;
            Excel.Panes? panes = null;
            Excel.Pane? pane = null;
            try
            {
                sheet = (Excel.Worksheet)context.Book.ActiveSheet;
                Assert.Equal("Sheet1", sheet.Name);
                selection = sheet.Range["R20"];
                selection.Select();
                window = context.App.ActiveWindow;
                panes = window.Panes;
                pane = panes[2];
                // Keep the selection inside BOTH the original viewport and the
                // capture viewport, outside the captured cells. Verification activation must not repair a
                // missed scroll restoration by scrolling to an offscreen selection.
                pane.ScrollRow = 10;
                pane.ScrollColumn = 4;
            }
            finally
            {
                ComUtilities.Release(ref pane);
                ComUtilities.Release(ref panes);
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref selection);
                ComUtilities.Release(ref sheet);
            }
        });

    private sealed record InactiveGeometryCaptureState(string OtherSheet, string TargetAddress,
        GeometryView OtherView, string OtherDiagnostic, string Cells);

    private InactiveGeometryCaptureState PrepareInactiveGeometryCapture(string otherSheet, string targetAddress)
    {
        string cells = ReadInactiveGeometryCells(otherSheet, targetAddress);
        ActivateGeometryVerificationSheet(otherSheet, arrangeOther: true);
        var saved = _fixture.Send("session.save", new { });
        Assert.True(saved.Success, saved.ErrorMessage);
        var other = ReadGeometryView();
        Assert.Equal(otherSheet, other.View.Sheet);
        Assert.True(other.View.Saved, other.Diagnostic);
        return new InactiveGeometryCaptureState(otherSheet, targetAddress, other.View, other.Diagnostic, cells);
    }

    private (GeometryView View, string Diagnostic) ReadInactiveGeometryTargetAndRestoreOther(InactiveGeometryCaptureState state)
    {
        Exception? failure = null;
        var restorationFailures = new List<Exception>();
        (GeometryView View, string Diagnostic) target = default;
        try
        {
            var other = ReadGeometryView();
            Assert.True(state.OtherView == other.View,
                $"Capture did not restore the previously active sheet/view. Before={state.OtherDiagnostic}; after={other.Diagnostic}");
            Assert.Equal(state.Cells, ReadInactiveGeometryCells(state.OtherSheet, state.TargetAddress));
            ActivateGeometryVerificationSheet("Sheet1", arrangeOther: false);
            target = ReadGeometryView();
            Assert.Equal("Sheet1", target.View.Sheet);
            Assert.True(target.View.Saved, target.Diagnostic);
        }
        catch (Exception error) { failure = error; }
        finally
        {
            // Verification needs to activate the target because Excel exposes
            // worksheet pane state through ActiveWindow. Restore the other sheet
            // independently even if the target/source assertions failed.
            try { ActivateGeometryVerificationSheet(state.OtherSheet, arrangeOther: false); }
            catch (Exception error) { restorationFailures.Add(error); }
            try
            {
                var restored = ReadGeometryView();
                Assert.True(state.OtherView == restored.View,
                    $"Verification changed the previously active view. Before={state.OtherDiagnostic}; after={restored.Diagnostic}");
                Assert.Equal(state.Cells, ReadInactiveGeometryCells(state.OtherSheet, state.TargetAddress));
            }
            catch (Exception error) { restorationFailures.Add(error); }
        }
        if (restorationFailures.Count > 0)
        {
            if (failure is not null) restorationFailures.Insert(0, failure);
            throw new AggregateException("Non-active screenshot target verification or restoration failed.", restorationFailures);
        }
        if (failure is not null) ExceptionDispatchInfo.Capture(failure).Throw();
        return target;
    }

    private void ActivateGeometryVerificationSheet(string name, bool arrangeOther) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? selection = null;
            Excel.Window? window = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[name];
                sheet.Activate();
                if (arrangeOther)
                {
                    window = context.App.ActiveWindow;
                    window.FreezePanes = false;
                    window.SplitRow = 0;
                    window.SplitColumn = 0;
                    window.Zoom = 110;
                    window.DisplayGridlines = true;
                    window.DisplayHeadings = false;
                    selection = sheet.Range["K12"];
                    selection.Select();
                    window.ScrollRow = 7;
                    window.ScrollColumn = 5;
                }
            }
            finally
            {
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref selection);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

    private string ReadInactiveGeometryCells(string otherSheet, string targetAddress) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? target = null;
            Excel.Worksheet? other = null;
            Excel.Range? targetCells = null;
            Excel.Range? otherCells = null;
            try
            {
                sheets = context.Book.Worksheets;
                target = (Excel.Worksheet)sheets["Sheet1"];
                other = (Excel.Worksheet)sheets[otherSheet];
                targetCells = target.Range[targetAddress];
                otherCells = other.Range["A1:B2"];
                return JsonSerializer.Serialize(new
                {
                    Target = ReadGeometryCellContent(targetCells),
                    Other = ReadGeometryCellContent(otherCells)
                });
            }
            finally
            {
                ComUtilities.Release(ref otherCells);
                ComUtilities.Release(ref targetCells);
                ComUtilities.Release(ref other);
                ComUtilities.Release(ref target);
                ComUtilities.Release(ref sheets);
            }
        });

    private sealed record GeometryCellContent(object? Value, object? Formula, int Fill);

    private static List<GeometryCellContent> ReadGeometryCellContent(Excel.Range range)
    {
        Excel.Range? cells = null;
        var snapshot = new List<GeometryCellContent>();
        try
        {
            cells = range.Cells;
            for (int index = 1; index <= cells.Count; index++)
            {
                Excel.Range? cell = null;
                Excel.Interior? fill = null;
                try
                {
                    cell = (Excel.Range)cells[index];
                    fill = cell.Interior;
                    snapshot.Add(new GeometryCellContent((object?)cell.Value2, (object?)cell.Formula,
                        Convert.ToInt32((object)fill.Color, CultureInfo.InvariantCulture)));
                }
                finally
                {
                    ComUtilities.Release(ref fill);
                    ComUtilities.Release(ref cell);
                }
            }
            return snapshot;
        }
        finally { ComUtilities.Release(ref cells); }
    }
}
