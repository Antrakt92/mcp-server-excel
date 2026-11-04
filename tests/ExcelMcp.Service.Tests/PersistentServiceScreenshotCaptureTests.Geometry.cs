using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class IsolatedServiceScreenshotTests
{
    [Theory]
    [InlineData(false, false, "A1:J10", "A1:J5", "A6:J10", "$A$1:$J$10")]
    [InlineData(true, false, "A1:J10", "A1:J5", "A6:J10", "$A$1:$J$10")]
    [InlineData(false, false, "D17:M26", "D17:M21", "D22:M26", "$D$17:$M$26")]
    [InlineData(true, false, "D17:M26", "D17:M21", "D22:M26", "$D$17:$M$26")]
    [InlineData(false, true, "A1:J10", "A1:J5", "A6:J10", "$A$1:$J$10")]
    [InlineData(false, true, "D17:M26", "D17:M21", "D22:M26", "$D$17:$M$26")]
    public void CaptureRange_DisplacedWindow_PreservesMarkerGeometryAndView(
        bool maximized, bool initiallyHidden, string address, string topAddress, string bottomAddress, string expectedAddress)
        => AssertGeometryCapture(maximized, initiallyHidden, address, topAddress, bottomAddress, expectedAddress);

    [Theory]
    [InlineData("frozen", "A1:J10", "A1:J5", "A6:J10", "$A$1:$J$10")]
    [InlineData("frozen", "D17:M26", "D17:M21", "D22:M26", "$D$17:$M$26")]
    [InlineData("split", "D17:M26", "D17:M21", "D22:M26", "$D$17:$M$26")]
    [InlineData("frozen-rows", "D17:M26", "D17:M21", "D22:M26", "$D$17:$M$26")]
    [InlineData("frozen-columns", "D17:M26", "D17:M21", "D22:M26", "$D$17:$M$26")]
    [InlineData("split-rows", "D17:M26", "D17:M21", "D22:M26", "$D$17:$M$26")]
    [InlineData("split-columns", "D17:M26", "D17:M21", "D22:M26", "$D$17:$M$26")]
    public void CaptureRange_FrozenOrSplitPanes_PreservesMarkerGeometryAndEveryPane(
        string paneMode, string address, string topAddress, string bottomAddress, string expectedAddress)
        => AssertGeometryCapture(false, false, address, topAddress, bottomAddress, expectedAddress, paneMode);

    [Fact]
    public async Task CaptureRange_ColumnWiderThanViewport_FailsAndRestoresSavedPaneView()
    {
        PrepareGeometryMarkers(false, false, "A1:J10", "A1:J5", "A6:J10", "frozen");
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? target = null;
            try
            {
                context.App.Width = 360;
                context.App.Height = 360;
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets["Sheet1"];
                target = sheet.Range["A1:J10"];
                target.ColumnWidth = 255;
            }
            finally
            {
                ComUtilities.Release(ref target);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        var saved = _fixture.Send("session.save", new { });
        Assert.True(saved.Success, saved.ErrorMessage);
        var before = ReadGeometryView();
        Assert.True(before.View.Saved, before.Diagnostic);

        var failure = await _fixture.SendForFailureAsync("screenshot.capture",
            new { sheetName = "Sheet1", rangeAddress = "A1:J10", quality = "High" });

        var after = ReadGeometryView();
        string diagnostic = $"before={before.Diagnostic}; after={after.Diagnostic}";
        Assert.False(failure.Success, $"An oversized column must not silently produce a clipped screenshot; {diagnostic}");
        Assert.Contains("larger than", failure.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("viewport", failure.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.True(before.View == after.View, $"Failed capture changed saved/pane/view state; {diagnostic}");
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task CaptureRange_FullyHiddenRowsOrColumns_FailsWithoutChangingSavedView(bool hideRows)
    {
        PrepareGeometryMarkers(false, false, "A1:J10", "A1:J5", "A6:J10", "none");
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? target = null;
            Excel.Range? axis = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets["Sheet1"];
                target = sheet.Range["A1:J10"];
                axis = hideRows ? target.EntireRow : target.EntireColumn;
                axis.Hidden = true;
                Assert.True(Convert.ToBoolean((object)axis.Hidden, CultureInfo.InvariantCulture));
            }
            finally
            {
                ComUtilities.Release(ref axis);
                ComUtilities.Release(ref target);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        var saved = _fixture.Send("session.save", new { });
        Assert.True(saved.Success, saved.ErrorMessage);
        var before = ReadGeometryView();
        Assert.True(before.View.Saved, before.Diagnostic);

        var failure = await _fixture.SendForFailureAsync("screenshot.capture",
            new { sheetName = "Sheet1", rangeAddress = "A1:J10", quality = "High" });

        var after = ReadGeometryView();
        string diagnostic = $"hideRows={hideRows}; before={before.Diagnostic}; after={after.Diagnostic}";
        Assert.False(failure.Success, diagnostic);
        Assert.Contains("no visible cell area", failure.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.True(before.View == after.View, $"Failed capture changed saved/view state; {diagnostic}");
    }

    private void AssertGeometryCapture(bool maximized, bool initiallyHidden, string address,
        string topAddress, string bottomAddress, string expectedAddress, string paneMode = "none")
    {
        PrepareGeometryMarkers(maximized, initiallyHidden, address, topAddress, bottomAddress, paneMode);
        var saved = _fixture.Send("session.save", new { });
        Assert.True(saved.Success, saved.ErrorMessage);
        var before = ReadGeometryView();
        Assert.True(before.View.Saved, before.Diagnostic);

        ScreenshotResult? result = null;
        GeometryFrozenDivider? frozenDivider = MeasureGeometryFrozenDivider(address);
        var captureFailure = Record.Exception(() => result = _screenshotCommands.CaptureRange(
            _fixture.BatchToken, "Sheet1", address, ScreenshotQuality.High));

        var after = ReadGeometryView();
        string diagnostic = $"maximized={maximized}; initiallyHidden={initiallyHidden}; paneMode={paneMode}; range={address}; before={before.Diagnostic}; after={after.Diagnostic}";
        Assert.True(captureFailure is null, $"Capture failed: {captureFailure}; {diagnostic}");
        Assert.NotNull(result);
        Assert.True(result.Success, $"{result.ErrorMessage}; {diagnostic}");
        Assert.Equal("Sheet1", result.SheetName);
        Assert.Equal(expectedAddress, result.RangeAddress);
        Assert.Equal("image/png", result.MimeType);
        Assert.True(before.View == after.View, $"View changed; {diagnostic}");
        Assert.NotEmpty(result.ImageBase64);
        using var stream = new MemoryStream(Convert.FromBase64String(result.ImageBase64));
        using var image = new Bitmap(stream);
        Assert.Equal(result.Width, image.Width);
        Assert.Equal(result.Height, image.Height);
        Assert.True(image.Width >= 100 && image.Height >= 100, diagnostic);
        string? artifactDirectory = Environment.GetEnvironmentVariable("EXCELMCP_TEST_ARTIFACT_DIRECTORY");
        if (!string.IsNullOrWhiteSpace(artifactDirectory))
        {
            Assert.True(Path.IsPathFullyQualified(artifactDirectory) && Directory.Exists(artifactDirectory),
                "Synthetic screenshot artifact directory must be an existing absolute directory.");
            string artifact = Path.Combine(artifactDirectory, $"{Guid.NewGuid():N}.png");
            using var output = new FileStream(artifact, FileMode.CreateNew, FileAccess.Write, FileShare.Read);
            output.Write(Convert.FromBase64String(result.ImageBase64));
            diagnostic += $"; syntheticImage={artifact}";
        }

        // Every interior row must contain the expected half's fill. A crop that
        // includes the ribbon, headers or bottom chrome fails even when it also
        // contains enough worksheet pixels to pass a nonblank/marker-presence test.
        int nativeDividerRows = 0;
        for (int y = 3; y < image.Height - 3; y++)
        {
            if (Math.Abs(y - image.Height / 2) <= 3) continue;
            Color expected = y < image.Height / 2 ? Color.Navy : Color.Maroon;
            int matches = 0;
            for (int x = 3; x < image.Width - 3; x++)
            {
                Color pixel = image.GetPixel(x, y);
                if (Math.Abs(pixel.R - expected.R) <= 8 && Math.Abs(pixel.G - expected.G) <= 8 &&
                    Math.Abs(pixel.B - expected.B) <= 8) matches++;
            }
            Color center = image.GetPixel(image.Width / 2, y);
            // Excel paints one gray pixel row on the frozen region's bottom edge.
            // Retain that native view detail only at the independently measured
            // COM boundary; a gap, chrome strip or second divider still fails.
            if (matches < (image.Width - 6) * 0.95 && frozenDivider is not null &&
                Math.Abs(y - frozenDivider.Row) <= 1 &&
                Enumerable.Range(3, image.Width - 6).All(x =>
                {
                    Color pixel = image.GetPixel(x, y);
                    return pixel.ToArgb() == frozenDivider.NativeColor.ToArgb();
                }))
            {
                nativeDividerRows++;
                Assert.True(nativeDividerRows <= 1, $"More than one native divider row; {diagnostic}");
                continue;
            }
            Assert.True(matches >= (image.Width - 6) * 0.95,
                $"Marker geometry mismatch at y={y}, expected={expected.Name}, matching={matches}/{image.Width - 6}, " +
                $"centerRGB=({center.R},{center.G},{center.B}), " +
                $"centerAbove={image.GetPixel(image.Width / 2, y - 1)}, centerBelow={image.GetPixel(image.Width / 2, y + 1)}, " +
                $"image={image.Width}x{image.Height}; {diagnostic}");
        }
    }

    private GeometryFrozenDivider? MeasureGeometryFrozenDivider(string address) =>
        _fixture.ExecuteRawVerification((context, _) => DpiAwareness.Execute<GeometryFrozenDivider?>(() =>
        {
            Excel.Window? window = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? target = null;
            Excel.Range? frozenRows = null;
            Excel.Range? targetRows = null;
            Excel.Panes? panes = null;
            Excel.Pane? fixedPane = null;
            Excel.Range? visible = null;
            try
            {
                window = context.App.ActiveWindow;
                if (!window.FreezePanes || window.SplitRow == 0) return null;
                sheet = (Excel.Worksheet)context.Book.ActiveSheet;
                target = sheet.Range[address];
                targetRows = target.Rows;
                int first = target.Row;
                int boundary = window.SplitRow;
                if (boundary < first || boundary >= first + targetRows.Count - 1) return null;
                frozenRows = sheet.Range[$"A{first}:A{boundary}"];
                double height = Convert.ToDouble((object)frozenRows.Height, CultureInfo.InvariantCulture);
                double zoom = Convert.ToDouble((object)window.Zoom, CultureInfo.InvariantCulture) / 100.0;
                uint dpi = GeometryNative.GetDpiForWindow((nint)context.App.Hwnd);
                Assert.True(dpi > 0);
                int row = (int)Math.Round(height * dpi / 72.0 * zoom, MidpointRounding.AwayFromZero) - 1;
                panes = window.Panes;
                fixedPane = panes[1];
                visible = fixedPane.VisibleRange;
                nint hwnd = (nint)context.App.Hwnd;
                Rectangle bounds = WindowCapture.GetWindowBounds(hwnd);
                using var native = WindowCapture.CaptureWindow(hwnd);
                int left = fixedPane.PointsToScreenPixelsX(0) - bounds.Left;
                int top = fixedPane.PointsToScreenPixelsY(0) - bounds.Top;
                int span = (int)Math.Round(Convert.ToDouble((object)visible.Width, CultureInfo.InvariantCulture) * dpi / 72.0 * zoom);
                int fromX = left + 3;
                int toX = left + span - 3;
                Assert.True(fromX >= 0 && toX > fromX && toX <= native.Width);
                var candidates = new List<Color>();
                for (int delta = -1; delta <= 1; delta++)
                {
                    int y = top + row + delta;
                    Assert.InRange(y, 1, native.Height - 1);
                    Color color = native.GetPixel((fromX + toX) / 2, y);
                    bool coherent = color.R == color.G && color.G == color.B;
                    for (int x = fromX; coherent && x < toX; x++)
                    {
                        Color above = native.GetPixel(x, y - 1);
                        coherent = native.GetPixel(x, y).ToArgb() == color.ToArgb()
                            && Math.Abs(above.R - Color.Navy.R) <= 8
                            && Math.Abs(above.G - Color.Navy.G) <= 8
                            && Math.Abs(above.B - Color.Navy.B) <= 8;
                    }
                    if (coherent) candidates.Add(color);
                }
                Assert.Single(candidates);
                return new GeometryFrozenDivider(row, candidates[0]);
            }
            finally
            {
                ComUtilities.Release(ref visible);
                ComUtilities.Release(ref fixedPane);
                ComUtilities.Release(ref panes);
                ComUtilities.Release(ref frozenRows);
                ComUtilities.Release(ref targetRows);
                ComUtilities.Release(ref target);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref window);
            }
        }));

    private sealed record GeometryFrozenDivider(int Row, Color NativeColor);

    private void PrepareGeometryMarkers(bool maximized, bool initiallyHidden, string address,
        string topAddress, string bottomAddress, string paneMode) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Window? window = null;
            Excel.Range? target = null;
            Excel.Range? top = null;
            Excel.Range? bottom = null;
            Excel.Range? parking = null;
            Excel.Interior? topFill = null;
            Excel.Interior? bottomFill = null;
            try
            {
                context.App.WindowState = Excel.XlWindowState.xlNormal;
                context.App.Left = 120;
                context.App.Top = 90;
                context.App.Width = 720;
                context.App.Height = 540;
                Assert.True(context.App.Left > 50 && context.App.Top > 50,
                    "The owned Excel window must start away from the desktop origin.");
                if (maximized) context.App.WindowState = Excel.XlWindowState.xlMaximized;
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets["Sheet1"];
                sheet.Activate();
                target = sheet.Range[address];
                target.RowHeight = 20;
                target.ColumnWidth = 8;
                top = sheet.Range[topAddress];
                bottom = sheet.Range[bottomAddress];
                topFill = top.Interior;
                bottomFill = bottom.Interior;
                topFill.Color = ColorTranslator.ToOle(Color.Navy);
                bottomFill.Color = ColorTranslator.ToOle(Color.Maroon);
                Assert.Equal(ColorTranslator.ToOle(Color.Navy), Convert.ToInt32((object)topFill.Color, CultureInfo.InvariantCulture));
                Assert.Equal(ColorTranslator.ToOle(Color.Maroon), Convert.ToInt32((object)bottomFill.Color, CultureInfo.InvariantCulture));
                parking = sheet.Range["X45"];
                parking.Select();
                window = context.App.ActiveWindow;
                window.DisplayGridlines = false;
                window.DisplayHeadings = true;
                window.Zoom = 75;
                window.ScrollRow = 40;
                window.ScrollColumn = 15;
                if (paneMode != "none") ConfigureGeometryPanes(window, paneMode);
                context.App.Visible = !initiallyHidden;
            }
            finally
            {
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref parking);
                ComUtilities.Release(ref bottomFill);
                ComUtilities.Release(ref topFill);
                ComUtilities.Release(ref bottom);
                ComUtilities.Release(ref top);
                ComUtilities.Release(ref target);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

    private static void ConfigureGeometryPanes(Excel.Window window, string paneMode)
    {
        bool frozen = paneMode.StartsWith("frozen", StringComparison.Ordinal);
        bool rows = !paneMode.EndsWith("-columns", StringComparison.Ordinal);
        bool columns = !paneMode.EndsWith("-rows", StringComparison.Ordinal);
        int count = (rows ? 2 : 1) * (columns ? 2 : 1);

        (int Row, int Column) ExpectedPosition(int index)
        {
            // Native Pane.VisibleRange readback establishes column-major order:
            // top-left, bottom-left, top-right, bottom-right in a four-pane view.
            int rowPaneCount = rows ? 2 : 1;
            bool bottom = rows && (index - 1) % rowPaneCount == 1;
            bool right = columns && (index - 1) / rowPaneCount == 1;
            return (!rows || bottom ? 40 : frozen ? 1 : 3,
                !columns || right ? 15 : frozen ? 1 : 2);
        }

        Excel.Panes? panes = null;
        Excel.Pane? pane = null;
        Excel.Range? visible = null;
        try
        {
            window.ScrollRow = 1;
            window.ScrollColumn = 1;
            window.SplitRow = rows ? 2 : 0;
            window.SplitColumn = columns ? 2 : 0;
            window.FreezePanes = frozen;
            panes = window.Panes;
            Assert.Equal(count, panes.Count);
            // Activation can scroll to the pane's selection. Complete it before
            // establishing unequal pane origins, with no later window scroll writes.
            pane = panes[count];
            pane.Activate();
            ComUtilities.Release(ref pane);
            for (int index = 1; index <= panes.Count; index++)
            {
                try
                {
                    pane = panes[index];
                    var expected = ExpectedPosition(index);
                    if (!frozen || !rows || expected.Row != 1) pane.ScrollRow = expected.Row;
                    if (!frozen || !columns || expected.Column != 1) pane.ScrollColumn = expected.Column;
                }
                finally { ComUtilities.Release(ref pane); }
            }
            for (int index = 1; index <= panes.Count; index++)
            {
                try
                {
                    pane = panes[index];
                    var expected = ExpectedPosition(index);
                    visible = pane.VisibleRange;
                    Assert.Equal(expected.Row, visible.Row);
                    Assert.Equal(expected.Column, visible.Column);
                }
                finally
                {
                    ComUtilities.Release(ref visible);
                    ComUtilities.Release(ref pane);
                }
            }
            Assert.Equal(frozen, window.FreezePanes);
            Assert.Equal(rows ? 2 : 0, window.SplitRow);
            Assert.Equal(columns ? 2 : 0, window.SplitColumn);
        }
        finally
        {
            ComUtilities.Release(ref visible);
            ComUtilities.Release(ref pane);
            ComUtilities.Release(ref panes);
        }
    }

    private (GeometryView View, string Diagnostic) ReadGeometryView() =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Window? window = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? selection = null;
            Excel.Pane? activePane = null;
            try
            {
                window = context.App.ActiveWindow;
                sheet = (Excel.Worksheet)context.Book.ActiveSheet;
                selection = (Excel.Range)context.App.Selection;
                activePane = window.ActivePane;
                var paneState = ReadGeometryPanes(window);
                var view = new GeometryView(sheet.Name, selection.Address,
                    Convert.ToInt32((object)window.Zoom, CultureInfo.InvariantCulture),
                    window.DisplayGridlines,
                    window.DisplayHeadings, context.App.WindowState, context.App.Visible,
                    activePane.Index, window.FreezePanes, window.SplitRow, window.SplitColumn, paneState.State,
                    context.Book.Saved, window.SplitHorizontal, window.SplitVertical);
                nint hwnd = (nint)context.App.Hwnd;
                Assert.True(GeometryNative.GetWindowRect(hwnd, out var bounds));
                string diagnostic = string.Create(CultureInfo.InvariantCulture,
                    $"hwnd={hwnd}; bounds=({bounds.Left},{bounds.Top},{bounds.Right},{bounds.Bottom}); " +
                    $"appLeftTop=({context.App.Left},{context.App.Top}); windowLeftTop=({window.Left},{window.Top}); " +
                    $"windowScrollGetters=({window.ScrollRow},{window.ScrollColumn}); " +
                    $"screenZero=({window.PointsToScreenPixelsX(0)},{window.PointsToScreenPixelsY(0)}); " +
                    $"dpi={GeometryNative.GetDpiForWindow(hwnd)}; panes={paneState.Diagnostic}; view={view}");
                return (view, diagnostic);
            }
            finally
            {
                ComUtilities.Release(ref activePane);
                ComUtilities.Release(ref selection);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref window);
            }
        });

    private static (string State, string Diagnostic) ReadGeometryPanes(Excel.Window window)
    {
        Excel.Panes? panes = null;
        Excel.Pane? pane = null;
        Excel.Range? visible = null;
        var state = new List<string>();
        var diagnostic = new List<string>();
        try
        {
            panes = window.Panes;
            for (int index = 1; index <= panes.Count; index++)
            {
                try
                {
                    pane = panes[index];
                    visible = pane.VisibleRange;
                    // Native Excel reports some Pane.Scroll getters in a different
                    // order from setters and VisibleRange; compare physical origins.
                    state.Add(string.Create(CultureInfo.InvariantCulture, $"{pane.Index}:{visible.Row},{visible.Column}"));
                    double left = Convert.ToDouble((object)visible.Left, CultureInfo.InvariantCulture);
                    double top = Convert.ToDouble((object)visible.Top, CultureInfo.InvariantCulture);
                    diagnostic.Add(string.Create(CultureInfo.InvariantCulture,
                        $"{pane.Index}:scrollGetters=({pane.ScrollRow},{pane.ScrollColumn}); visible={visible.Address}; visibleLeftTop=({left},{top}); " +
                        $"paneScreenZero=({pane.PointsToScreenPixelsX(0)},{pane.PointsToScreenPixelsY(0)}); " +
                        $"windowMappedVisible=({window.PointsToScreenPixelsX((int)left)},{window.PointsToScreenPixelsY((int)top)})"));
                }
                finally
                {
                    ComUtilities.Release(ref visible);
                    ComUtilities.Release(ref pane);
                }
            }
            return (string.Join(";", state), string.Join(";", diagnostic));
        }
        finally
        {
            ComUtilities.Release(ref visible);
            ComUtilities.Release(ref pane);
            ComUtilities.Release(ref panes);
        }
    }

    private sealed record GeometryView(string Sheet, string Selection, int Zoom,
        bool Gridlines, bool Headings, Excel.XlWindowState WindowState, bool Visible,
        int ActivePaneIndex, bool FreezePanes, int SplitRow, int SplitColumn, string PaneVisibleOrigins, bool Saved,
        double SplitHorizontal, double SplitVertical);

    private static class GeometryNative
    {
        [StructLayout(LayoutKind.Sequential)]
        internal struct WindowRect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetWindowRect(nint window, out WindowRect bounds);

        [DllImport("user32.dll")]
        internal static extern uint GetDpiForWindow(nint window);
    }
}
