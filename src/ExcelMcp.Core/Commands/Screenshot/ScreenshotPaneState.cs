using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Screenshot;

/// <summary>Retains pane layout and scroll positions without changing the layout.</summary>
internal sealed class ScreenshotPaneState
{
    private sealed record PaneState(int Index, int ScrollRow, int ScrollColumn);

    private readonly bool _split;
    private readonly bool _frozen;
    private readonly int _splitRow;
    private readonly int _splitColumn;
    private readonly double _splitHorizontal;
    private readonly double _splitVertical;
    private readonly int _activePane;
    private readonly PaneState[] _panes;

    private ScreenshotPaneState(Excel.Window window)
    {
        _split = window.Split;
        _frozen = window.FreezePanes;
        _splitRow = window.SplitRow;
        _splitColumn = window.SplitColumn;
        _splitHorizontal = window.SplitHorizontal;
        _splitVertical = window.SplitVertical;
        Excel.Panes? panes = null;
        Excel.Pane? active = null;
        try
        {
            panes = window.Panes;
            active = window.ActivePane;
            _activePane = active.Index;
            _panes = new PaneState[panes.Count];
            for (int index = 1; index <= _panes.Length; index++)
            {
                Excel.Pane? pane = null;
                Excel.Range? visible = null;
                try
                {
                    pane = panes[index];
                    visible = pane.VisibleRange;
                    // In four-pane Excel views ScrollRow/Column getters can describe
                    // a different pane. VisibleRange identifies the actual painted origin.
                    _panes[index - 1] = new PaneState(pane.Index, visible.Row, visible.Column);
                }
                finally
                {
                    ComUtilities.Release(ref visible);
                    ComUtilities.Release(ref pane);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref active);
            ComUtilities.Release(ref panes);
        }
    }

    internal static ScreenshotPaneState Capture(Excel.Window window) => new(window);

    internal bool HasMultiplePanes => _frozen || _split || _panes.Length > 1;

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Attempts every pane restoration and readback, then throws all failures; no incomplete restoration is reported as successful.")]
    internal void Restore(Excel.Window window)
    {
        var failures = new List<Exception>();
        void Attempt(Action action)
        {
            try { action(); }
            catch (Exception error) { failures.Add(error); }
        }

        // The caller restores zoom first. Avoid redundant activation, which can move the
        // selection or scroll to it when it was outside the original viewport.
        Attempt(() =>
        {
            Excel.Pane? active = null;
            try
            {
                active = window.ActivePane;
                if (active.Index != _activePane)
                    WithPane(window, _activePane, pane => pane.Activate());
            }
            finally { ComUtilities.Release(ref active); }
        });
        foreach (var saved in _panes)
        {
            Attempt(() => WithPane(window, saved.Index, pane =>
            {
                if (ReadOrigin(pane).Row != saved.ScrollRow) pane.ScrollRow = saved.ScrollRow;
            }));
            Attempt(() => WithPane(window, saved.Index, pane =>
            {
                if (ReadOrigin(pane).Column != saved.ScrollColumn) pane.ScrollColumn = saved.ScrollColumn;
            }));
        }
        Attempt(() => Verify(window));
        if (failures.Count > 0)
            throw new AggregateException("Screenshot pane restoration failed.", failures);
    }

    private void Verify(Excel.Window window)
    {
        var actual = Capture(window);
        if (actual._split != _split || actual._frozen != _frozen
            || actual._splitRow != _splitRow || actual._splitColumn != _splitColumn
            || actual._splitHorizontal != _splitHorizontal || actual._splitVertical != _splitVertical
            || actual._activePane != _activePane || !actual._panes.SequenceEqual(_panes))
            throw new InvalidOperationException(
                "Excel did not restore the complete screenshot pane layout and scroll state. " +
                $"Expected {Describe()}; observed {actual.Describe()}.");
    }

    private string Describe() =>
        $"split={_split}, frozen={_frozen}, boundary=({_splitRow},{_splitColumn}), " +
        $"points=({_splitHorizontal},{_splitVertical}), activePane={_activePane}, " +
        $"scroll=[{string.Join(';', _panes.Select(pane => $"{pane.Index}:{pane.ScrollRow},{pane.ScrollColumn}"))}]";

    private static (int Row, int Column) ReadOrigin(Excel.Pane pane)
    {
        Excel.Range? visible = null;
        try
        {
            visible = pane.VisibleRange;
            return (visible.Row, visible.Column);
        }
        finally { ComUtilities.Release(ref visible); }
    }

    private static void WithPane(Excel.Window window, int index, Action<Excel.Pane> action)
    {
        Excel.Panes? panes = null;
        Excel.Pane? pane = null;
        try
        {
            panes = window.Panes;
            pane = panes[index];
            action(pane);
        }
        finally
        {
            ComUtilities.Release(ref pane);
            ComUtilities.Release(ref panes);
        }
    }
}
