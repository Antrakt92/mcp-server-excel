using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Screenshot;

public partial class ScreenshotCommands
{
    private sealed record CapturePane(int Index, int Row, int Column, int Rows, int Columns,
        bool FixedRows, bool FixedColumns, double Area);

    private sealed record PreparedPaneRegion(CapturePane Pane, int Row, int Column, int Rows, int Columns,
        CapturePlanner.CapturePlan Plan, int[] Xs, int[] Ys, Point Destination);

    private static Bitmap CapturePaneTiles(Excel.Application app, Excel.Window window,
        Excel.Range range, ref CapturePlanner.CapturePlan plan)
    {
        var hwnd = GetExcelWindowHandle(app);
        double deviceScale = WindowCapture.GetWindowDpi(hwnd) / 72.0;
        double scale = deviceScale * plan.Zoom / 100.0;
        int firstRow = range.Row, firstColumn = range.Column;
        int rows = plan.RowSegments[^1].Start + plan.RowSegments[^1].Count - 1;
        int columns = plan.ColumnSegments[^1].Start + plan.ColumnSegments[^1].Count - 1;
        bool frozen = window.FreezePanes;
        int splitRow = frozen ? window.SplitRow : 0;
        int splitColumn = frozen ? window.SplitColumn : 0;
        var panes = ReadCapturePanes(window, splitRow, splitColumn);
        var preferred = panes.OrderByDescending(pane => pane.Area).First();
        var rowParts = DivideCaptureAxis(firstRow, rows, splitRow);
        var columnParts = DivideCaptureAxis(firstColumn, columns, splitColumn);
        Excel.Worksheet? sheet = null;
        Bitmap? canvas = null;
        try
        {
            sheet = range.Worksheet;
            var prepared = new List<PreparedPaneRegion>();
            int plannedTiles = 0, destinationY = 0, width = 0;
            foreach (var (row, rowCount) in rowParts)
            {
                int destinationX = 0, regionHeight = 0;
                foreach (var (column, columnCount) in columnParts)
                {
                    var selected = frozen
                        ? panes.FirstOrDefault(pane => pane.FixedRows == (splitRow > 0 && row <= splitRow)
                            && pane.FixedColumns == (splitColumn > 0 && column <= splitColumn))
                            ?? throw new InvalidOperationException("Excel did not expose the required frozen screenshot pane.")
                        : preferred;
                    using var region = new CaptureRangeLease(GetRange(sheet, row, column,
                        row + rowCount - 1, column + columnCount - 1));
                    var ready = PreparePaneRegion(window, selected, region.Range, deviceScale, scale,
                        plan.Zoom, new Point(destinationX, destinationY));
                    plannedTiles = checked(plannedTiles + ready.Plan.RowSegments.Count * ready.Plan.ColumnSegments.Count);
                    if (plannedTiles > 64)
                        throw new InvalidOperationException("The pane layout requires more than 64 screenshot tiles. Capture a smaller range.");
                    prepared.Add(ready);
                    regionHeight = ready.Ys[^1];
                    destinationX = checked(destinationX + ready.Xs[^1]);
                }
                width = destinationX;
                destinationY = checked(destinationY + regionHeight);
            }
            if (width <= 0 || destinationY <= 0)
                throw new InvalidOperationException("The requested screenshot range has no visible cell area.");
            canvas = new Bitmap(width, destinationY, PixelFormat.Format32bppArgb);
            using var graphics = Graphics.FromImage(canvas);
            graphics.Clear(Color.White);
            int tiles = 0;
            foreach (var ready in prepared)
            {
                using var region = new CaptureRangeLease(GetRange(sheet, ready.Row, ready.Column,
                    ready.Row + ready.Rows - 1, ready.Column + ready.Columns - 1));
                CapturePaneRegion(window, ready, region.Range, graphics, hwnd, deviceScale,
                    scale, plan.Zoom, ref tiles);
            }
            plan = plan with { CapturedTiles = tiles };
            var result = canvas;
            canvas = null;
            return result;
        }
        finally
        {
            canvas?.Dispose();
            ComUtilities.Release(ref sheet);
        }
    }

    private static List<(int Start, int Count)> DivideCaptureAxis(int start, int count, int split)
    {
        int last = checked(start + count - 1);
        return split >= start && split < last
            ? [(start, split - start + 1), (split + 1, last - split)]
            : [(start, count)];
    }

    private static List<CapturePane> ReadCapturePanes(Excel.Window window, int splitRow, int splitColumn)
    {
        Excel.Panes? panes = null;
        try
        {
            panes = window.Panes;
            var result = new List<CapturePane>();
            for (int index = 1; index <= panes.Count; index++)
            {
                Excel.Pane? pane = null;
                Excel.Range? visible = null;
                Excel.Range? rows = null;
                Excel.Range? columns = null;
                try
                {
                    pane = panes[index];
                    visible = pane.VisibleRange;
                    rows = visible.Rows;
                    columns = visible.Columns;
                    result.Add(new CapturePane(index, visible.Row, visible.Column, rows.Count, columns.Count,
                        splitRow > 0 && visible.Row <= splitRow,
                        splitColumn > 0 && visible.Column <= splitColumn,
                        Convert.ToDouble((object)visible.Width, CultureInfo.InvariantCulture)
                            * Convert.ToDouble((object)visible.Height, CultureInfo.InvariantCulture)));
                }
                finally
                {
                    ComUtilities.Release(ref columns);
                    ComUtilities.Release(ref rows);
                    ComUtilities.Release(ref visible);
                    ComUtilities.Release(ref pane);
                }
            }
            if (result.Count == 0)
                throw new InvalidOperationException("Excel did not expose a screenshot pane.");
            return result;
        }
        finally { ComUtilities.Release(ref panes); }
    }

    private static PreparedPaneRegion PreparePaneRegion(Excel.Window window, CapturePane selected,
        Excel.Range region, double deviceScale, double scale, int zoom, Point destination)
    {
        Excel.Panes? panes = null;
        Excel.Pane? pane = null;
        try
        {
            panes = window.Panes;
            pane = panes[selected.Index];
            if (selected.FixedRows && (region.Row < selected.Row
                || region.Row + GetPaneAxisCount(region, true) > selected.Row + selected.Rows)
                || selected.FixedColumns && (region.Column < selected.Column
                || region.Column + GetPaneAxisCount(region, false) > selected.Column + selected.Columns))
                throw new InvalidOperationException("The requested fixed screenshot cells are not visible in the frozen pane.");
            if (!selected.FixedRows) pane.ScrollRow = region.Row;
            if (!selected.FixedColumns) pane.ScrollColumn = region.Column;
            Thread.Sleep(ViewSettleMs);
            var usable = MeasureCapturePane(window, pane, selected, zoom);
            int paneWidth = (int)Math.Round(usable.Width * deviceScale);
            int paneHeight = (int)Math.Round(usable.Height * deviceScale);
            var regionPlan = CapturePlanner.ReplanPixels(region, zoom, usable, scale, paneWidth, paneHeight);
            if (regionPlan.Truncated)
                throw new InvalidOperationException("The pane layout requires more than 64 screenshot tiles. Capture a smaller range.");
            int[] xs = BuildPixelOffsets(region, false, regionPlan.ColumnSegments, scale, paneWidth);
            int[] ys = BuildPixelOffsets(region, true, regionPlan.RowSegments, scale, paneHeight);
            return new PreparedPaneRegion(selected, region.Row, region.Column,
                GetPaneAxisCount(region, true), GetPaneAxisCount(region, false), regionPlan, xs, ys, destination);
        }
        finally
        {
            ComUtilities.Release(ref pane);
            ComUtilities.Release(ref panes);
        }
    }

    private static void CapturePaneRegion(Excel.Window window, PreparedPaneRegion prepared, Excel.Range region,
        Graphics graphics, IntPtr hwnd, double deviceScale, double scale, int zoom, ref int tileCount)
    {
        var selected = prepared.Pane;
        var regionPlan = prepared.Plan;
        var xs = prepared.Xs;
        var ys = prepared.Ys;
        var destination = prepared.Destination;
        Excel.Panes? panes = null;
        Excel.Pane? pane = null;
        Excel.Worksheet? sheet = null;
        try
        {
            panes = window.Panes;
            pane = panes[selected.Index];
            sheet = region.Worksheet;
            int fixedX = selected.FixedColumns ? FixedRegionAnchor(pane, sheet, selected, region, false, scale) : 0;
            int fixedY = selected.FixedRows ? FixedRegionAnchor(pane, sheet, selected, region, true, scale) : 0;
            foreach (var row in regionPlan.RowSegments.Select((segment, index) => (segment, index)))
            {
                foreach (var column in regionPlan.ColumnSegments.Select((segment, index) => (segment, index)))
                {
                    Excel.Range? tile = null;
                    try
                    {
                        tile = GetTileRange(sheet, region, row.segment, column.segment);
                        if (!selected.FixedRows) pane.ScrollRow = tile.Row;
                        if (!selected.FixedColumns) pane.ScrollColumn = tile.Column;
                        Thread.Sleep(ViewSettleMs);
                        double left = Convert.ToDouble((object)tile.Left, CultureInfo.InvariantCulture);
                        double top = Convert.ToDouble((object)tile.Top, CultureInfo.InvariantCulture);
                        int x = selected.FixedColumns ? checked(fixedX + xs[column.index]) : PaneScreenPoint(pane, left, false);
                        int y = selected.FixedRows ? checked(fixedY + ys[row.index]) : PaneScreenPoint(pane, top, true);
                        int width = xs[column.index + 1] - xs[column.index];
                        int height = ys[row.index + 1] - ys[row.index];
                        var paneBounds = CapturePaneBounds(window, pane, selected, zoom, scale, deviceScale);
                        if (!paneBounds.Contains(new Rectangle(x, y, width, height)))
                            throw new InvalidOperationException(
                                $"The complete screenshot tile is not visible inside its Excel pane. Tile={new Rectangle(x, y, width, height)}; pane={paneBounds}.");
                        var bounds = WindowCapture.GetWindowBounds(hwnd);
                        using var shot = WindowCapture.CaptureWindow(hwnd);
                        EnsureWindowRendered(shot);
                        var source = new Rectangle(x - bounds.X, y - bounds.Y, width, height);
                        if (width <= 0 || height <= 0 || !new Rectangle(0, 0, shot.Width, shot.Height).Contains(source))
                            throw new InvalidOperationException("The complete screenshot tile does not fit inside the rendered Excel window.");
                        DrawTile(graphics, shot, source,
                            new Point(destination.X + xs[column.index], destination.Y + ys[row.index]));
                        tileCount++;
                    }
                    finally { ComUtilities.Release(ref tile); }
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref pane);
            ComUtilities.Release(ref panes);
        }
    }

    private static int PaneScreenPoint(Excel.Pane pane, double points, bool vertical)
    {
        int whole = checked((int)Math.Round(points, MidpointRounding.AwayFromZero));
        return vertical ? pane.PointsToScreenPixelsY(whole) : pane.PointsToScreenPixelsX(whole);
    }

    private static int FixedRegionAnchor(Excel.Pane pane, Excel.Worksheet sheet, CapturePane selected,
        Excel.Range region, bool rows, double scale)
    {
        Excel.Range? visible = null;
        Excel.Range? prefix = null;
        try
        {
            visible = pane.VisibleRange;
            double points = Convert.ToDouble(rows ? (object)visible.Top : (object)visible.Left, CultureInfo.InvariantCulture);
            int origin = PaneScreenPoint(pane, points, rows);
            int first = rows ? selected.Row : selected.Column;
            int start = rows ? region.Row : region.Column;
            if (start == first) return origin;
            prefix = rows
                ? GetRange(sheet, first, region.Column, start - 1, region.Column)
                : GetRange(sheet, region.Row, first, region.Row, start - 1);
            return checked(origin + PanePixelExtent(prefix, rows, scale));
        }
        finally
        {
            ComUtilities.Release(ref prefix);
            ComUtilities.Release(ref visible);
        }
    }

    private static Rectangle CapturePaneBounds(Excel.Window window, Excel.Pane pane,
        CapturePane selected, int zoom, double scale, double deviceScale)
    {
        Excel.Range? visible = null;
        try
        {
            visible = pane.VisibleRange;
            double left = Convert.ToDouble((object)visible.Left, CultureInfo.InvariantCulture);
            double top = Convert.ToDouble((object)visible.Top, CultureInfo.InvariantCulture);
            double width = MeasurePaneAxis(visible, false, selected.FixedColumns, zoom, window.UsableWidth);
            double height = MeasurePaneAxis(visible, true, selected.FixedRows, zoom, window.UsableHeight);
            int pixelWidth = selected.FixedColumns
                ? Math.Min(PanePixelExtent(visible, false, scale), (int)Math.Round((window.UsableWidth - 15) * deviceScale))
                : (int)Math.Round(width * deviceScale);
            int pixelHeight = selected.FixedRows
                ? Math.Min(PanePixelExtent(visible, true, scale), (int)Math.Round((window.UsableHeight - 15) * deviceScale))
                : (int)Math.Round(height * deviceScale);
            return new Rectangle(PaneScreenPoint(pane, left, false),
                PaneScreenPoint(pane, top, true), pixelWidth, pixelHeight);
        }
        finally { ComUtilities.Release(ref visible); }
    }

    private static CapturePlanner.UsablePane MeasureCapturePane(Excel.Window window, Excel.Pane pane,
        CapturePane selected, int zoom)
    {
        Excel.Range? visible = null;
        try
        {
            visible = pane.VisibleRange;
            return new CapturePlanner.UsablePane(
                MeasurePaneAxis(visible, false, selected.FixedColumns, zoom, window.UsableWidth),
                MeasurePaneAxis(visible, true, selected.FixedRows, zoom, window.UsableHeight));
        }
        finally { ComUtilities.Release(ref visible); }
    }

    private static double MeasurePaneAxis(Excel.Range visible, bool rows, bool fixedAxis, int zoom, double workspace)
    {
        Excel.Range? items = null;
        Excel.Range? first = null;
        Excel.Range? last = null;
        try
        {
            items = rows ? visible.Rows : visible.Columns;
            first = (Excel.Range)items[1];
            last = (Excel.Range)items[items.Count];
            double start = Convert.ToDouble(rows ? (object)first.Top : (object)first.Left, CultureInfo.InvariantCulture);
            double end = Convert.ToDouble(rows ? (object)last.Top : (object)last.Left, CultureInfo.InvariantCulture);
            if (fixedAxis || items.Count == 1)
                end += Convert.ToDouble(rows ? (object)last.Height : (object)last.Width, CultureInfo.InvariantCulture);
            double extent = (end - start) * zoom / 100.0 - (fixedAxis ? 0 : 15);
            extent = Math.Min(extent, workspace - 15);
            if (extent <= 0)
                throw new InvalidOperationException("The Excel pane has no measurable screenshot viewport.");
            return extent;
        }
        finally
        {
            ComUtilities.Release(ref last);
            ComUtilities.Release(ref first);
            ComUtilities.Release(ref items);
        }
    }

    private static int GetPaneAxisCount(Excel.Range range, bool rows)
    {
        Excel.Range? items = null;
        try
        {
            items = rows ? range.Rows : range.Columns;
            return items.Count;
        }
        finally { ComUtilities.Release(ref items); }
    }

    private static int PanePixelExtent(Excel.Range range, bool rows, double scale) =>
        BuildPixelOffsets(range, rows,
            [new CapturePlanner.Segment(1, GetPaneAxisCount(range, rows), 0, 0)], scale, int.MaxValue)[1];

    private sealed class CaptureRangeLease(Excel.Range range) : IDisposable
    {
        private Excel.Range? _range = range;
        internal Excel.Range Range => _range ?? throw new ObjectDisposedException(nameof(CaptureRangeLease));
        public void Dispose() => ComUtilities.Release(ref _range);
    }
}
