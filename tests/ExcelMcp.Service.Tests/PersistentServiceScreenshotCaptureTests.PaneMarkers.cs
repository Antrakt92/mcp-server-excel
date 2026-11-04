using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class IsolatedServiceScreenshotTests
{
    private static readonly (string Address, Color Color, int Row, int Column)[] PaneCornerMarkers =
    [
        ("A1", Color.Red, 0, 0),
        ("J1", Color.Green, 0, 9),
        ("A10", Color.Gold, 9, 0),
        ("J10", Color.Cyan, 9, 9)
    ];

    [Fact]
    public void CaptureRange_AsymmetricFrozenPanes_PreservesCornerPositionsAndEveryPane()
    {
        PrepareGeometryMarkers(false, false, "A1:J10", "A1:J5", "A6:J10", "none");
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? target = null;
            Excel.Window? window = null;
            Excel.Panes? panes = null;
            Excel.Pane? pane = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets["Sheet1"];
                target = sheet.Range["A1:J10"];
                target.ClearContents();
                foreach (var marker in PaneCornerMarkers)
                {
                    Excel.Range? cell = null;
                    Excel.Interior? fill = null;
                    try
                    {
                        cell = sheet.Range[marker.Address];
                        fill = cell.Interior;
                        fill.Color = ColorTranslator.ToOle(marker.Color);
                    }
                    finally
                    {
                        ComUtilities.Release(ref fill);
                        ComUtilities.Release(ref cell);
                    }
                }
                window = context.App.ActiveWindow;
                window.FreezePanes = false;
                window.SplitRow = 0;
                window.SplitColumn = 0;
                window.ScrollRow = 1;
                window.ScrollColumn = 1;
                window.SplitRow = 3;
                window.SplitColumn = 2;
                window.FreezePanes = true;
                panes = window.Panes;
                Assert.Equal(4, panes.Count);
                pane = panes[4];
                pane.Activate();
                ComUtilities.Release(ref pane);
                // Physical native order is column-major. Establish independent
                // movable origins only after activation has finished scrolling.
                for (int index = 1; index <= panes.Count; index++)
                {
                    try
                    {
                        pane = panes[index];
                        if (index is 2 or 4) pane.ScrollRow = 40;
                        if (index is 3 or 4) pane.ScrollColumn = 15;
                    }
                    finally { ComUtilities.Release(ref pane); }
                }
            }
            finally
            {
                ComUtilities.Release(ref pane);
                ComUtilities.Release(ref panes);
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref target);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        var saved = _fixture.Send("session.save", new { });
        Assert.True(saved.Success, saved.ErrorMessage);
        var before = ReadGeometryView();
        Assert.True(before.View.Saved && before.View.Visible && before.View.FreezePanes, before.Diagnostic);
        Assert.Equal(Excel.XlWindowState.xlNormal, before.View.WindowState);
        Assert.Equal(75, before.View.Zoom);
        Assert.Equal(3, before.View.SplitRow);
        Assert.Equal(2, before.View.SplitColumn);
        Assert.Equal("1:1,1;2:40,1;3:1,15;4:40,15", before.View.PaneVisibleOrigins);
        var sourceBefore = ReadPaneCornerColors();
        Assert.Equal(PaneCornerMarkers.Select(marker => ColorTranslator.ToOle(marker.Color)), sourceBefore);

        ScreenshotResult? result = null;
        var failure = Record.Exception(() => result = _screenshotCommands.CaptureRange(
            _fixture.BatchToken, "Sheet1", "A1:J10", ScreenshotQuality.High));

        var after = ReadGeometryView();
        string diagnostic = $"before={before.Diagnostic}; after={after.Diagnostic}";
        Assert.True(before.View == after.View, $"Capture changed saved state, selection, zoom or physical panes; {diagnostic}");
        Assert.Equal(sourceBefore, ReadPaneCornerColors());
        Assert.True(failure is null, $"Capture failed: {failure}; {diagnostic}");
        Assert.NotNull(result);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal("Sheet1", result.SheetName);
        Assert.Equal("$A$1:$J$10", result.RangeAddress);
        Assert.Equal("image/png", result.MimeType);
        Assert.NotEmpty(result.ImageBase64);
        using var stream = new MemoryStream(Convert.FromBase64String(result.ImageBase64));
        using var image = new Bitmap(stream);
        Assert.Equal(result.Width, image.Width);
        Assert.Equal(result.Height, image.Height);
        Assert.True(image.Width >= 100 && image.Height >= 100, diagnostic);
        // All ten rows/columns have equal source dimensions. Sample each corner
        // cell's center at its expected position, never search for marker colors.
        foreach (var marker in PaneCornerMarkers)
        {
            int x = (int)Math.Round((marker.Column + 0.5) * image.Width / 10);
            int y = (int)Math.Round((marker.Row + 0.5) * image.Height / 10);
            Color pixel = image.GetPixel(x, y);
            Assert.True(Math.Abs(pixel.R - marker.Color.R) <= 8
                        && Math.Abs(pixel.G - marker.Color.G) <= 8
                        && Math.Abs(pixel.B - marker.Color.B) <= 8,
                $"{marker.Address} center at ({x},{y}) expected {marker.Color.Name}, got {pixel}; " +
                $"image={image.Width}x{image.Height}; {diagnostic}");
        }
    }

    private int[] ReadPaneCornerColors() => _fixture.ExecuteRawVerification((context, _) =>
    {
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;
        try
        {
            sheets = context.Book.Worksheets;
            sheet = (Excel.Worksheet)sheets["Sheet1"];
            var colors = new int[PaneCornerMarkers.Length];
            for (int index = 0; index < PaneCornerMarkers.Length; index++)
            {
                Excel.Range? cell = null;
                Excel.Interior? fill = null;
                try
                {
                    cell = sheet.Range[PaneCornerMarkers[index].Address];
                    fill = cell.Interior;
                    colors[index] = Convert.ToInt32((object)fill.Color, CultureInfo.InvariantCulture);
                    Assert.Null((object?)cell.Value2);
                }
                finally
                {
                    ComUtilities.Release(ref fill);
                    ComUtilities.Release(ref cell);
                }
            }
            return colors;
        }
        finally
        {
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref sheets);
        }
    });
}
