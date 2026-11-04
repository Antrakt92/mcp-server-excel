using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace Sbroenne.ExcelMcp.Core.Commands.Screenshot;

/// <summary>
/// Captures the pixels of a top-level window using the Win32 <c>PrintWindow</c> API.
///
/// PrintWindow asks the window to render itself into a device context, so it works even when the
/// window is partially covered by other windows. Copying from the screen instead would capture
/// whatever happens to be on top.
/// </summary>
[SupportedOSPlatform("windows")]
internal static class WindowCapture
{
    private const uint PwRenderFullContent = 0x00000002;
    private const int DefaultDpi = 96;

    /// <summary>
    /// Gets the DPI of the monitor the window is on, falling back to 96 when unavailable.
    /// </summary>
    public static int GetWindowDpi(IntPtr hwnd)
    {
        try
        {
            uint dpi = GetDpiForWindow(hwnd);
            return dpi == 0 ? DefaultDpi : (int)dpi;
        }
        catch (EntryPointNotFoundException)
        {
            return DefaultDpi;
        }
    }

    /// <summary>
    /// Gets the window rectangle in physical screen pixels.
    /// </summary>
    public static Rectangle GetWindowBounds(IntPtr hwnd)
    {
        if (!GetWindowRect(hwnd, out Rect rect))
        {
            throw new InvalidOperationException(
                "Could not read the Excel window bounds needed for the screenshot. " +
                "The Excel window may have been closed. Retry the capture.");
        }

        return Rectangle.FromLTRB(rect.Left, rect.Top, rect.Right, rect.Bottom);
    }

    /// <summary>
    /// Renders the window into a bitmap.
    /// </summary>
    /// <param name="hwnd">Top-level window handle.</param>
    /// <param name="expectedBounds">Window bounds used to calculate the screenshot crop.</param>
    /// <returns>Bitmap of the whole window, in physical pixels. Caller owns the bitmap.</returns>
    public static Bitmap CaptureWindow(IntPtr hwnd, Rectangle expectedBounds)
    {
        // A standard render updates the viewport that full-content capture can otherwise
        // retain from before scrolling. Its bitmap can be incomplete, so discard it before
        // capturing the full-content frame that supplies the actual screenshot pixels.
        CaptureWindowCore(hwnd, 0, expectedBounds).Dispose();
        return CaptureWindowCore(hwnd, PwRenderFullContent, expectedBounds);
    }

    private static Bitmap CaptureWindowCore(IntPtr hwnd, uint flags, Rectangle expectedBounds)
    {
        Rectangle bounds = GetWindowBounds(hwnd);
        if (bounds != expectedBounds)
            throw new InvalidOperationException(
                $"The Excel window moved or resized before screenshot rendering. Expected {expectedBounds}; observed {bounds}. Retry the capture.");

        if (bounds.Width <= 0 || bounds.Height <= 0)
        {
            throw new InvalidOperationException(
                $"The Excel window has no capturable area ({bounds.Width}x{bounds.Height}). " +
                "Restore the Excel window to a normal size and retry the capture.");
        }

        var bitmap = new Bitmap(bounds.Width, bounds.Height, PixelFormat.Format32bppArgb);

        try
        {
            using var graphics = Graphics.FromImage(bitmap);
            IntPtr hdc = graphics.GetHdc();

            bool captured;
            try
            {
                captured = PrintWindow(hwnd, hdc, flags);
            }
            finally
            {
                graphics.ReleaseHdc(hdc);
            }

            if (!captured)
            {
                throw new InvalidOperationException(
                    "Windows could not render the Excel window for the screenshot. " +
                    "This happens when the desktop is locked or a Remote Desktop session is disconnected. " +
                    "Reconnect to an interactive desktop session and retry the capture.");
            }

            Rectangle renderedBounds = GetWindowBounds(hwnd);
            if (renderedBounds != expectedBounds)
                throw new InvalidOperationException(
                    $"The Excel window moved or resized during screenshot rendering. Expected {expectedBounds}; observed {renderedBounds}. Retry the capture.");

            return bitmap;
        }
        catch
        {
            bitmap.Dispose();
            throw;
        }
    }

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint flags);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr hwnd, out Rect lpRect);

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hwnd);

    [StructLayout(LayoutKind.Sequential)]
    private struct Rect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
}
