using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Slicer;

internal static class SlicerSelection
{
    internal static void SetNonOlapSelection(Excel.SlicerCache cache,
        List<string> selectedItems, bool clearFirst, CancellationToken ct)
    {
        var requested = new HashSet<string>(selectedItems, StringComparer.OrdinalIgnoreCase);
        bool selectAll = selectedItems.Count == 0;
        Excel.SlicerItems? items = null;
        try
        {
            items = cache.SlicerItems;
            if (!selectAll)
                RequireMatchingItem(items, requested, ct);

            // Excel resets the filter if its last selected item is deselected.
            // Select replacements first so removing old items never empties a valid selection.
            int passes = clearFirst && !selectAll ? 2 : 1;
            for (int pass = 0; pass < passes; pass++)
            {
                ct.ThrowIfCancellationRequested();
                for (int index = 1; index <= items.Count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.SlicerItem? item = null;
                    try
                    {
                        item = items.Item[index];
                        bool matches = requested.Contains(item.Name);
                        if (pass == 0 && (selectAll || matches))
                        {
                            ct.ThrowIfCancellationRequested();
                            item.Selected = true;
                        }
                        else if (pass == 1 && !matches)
                        {
                            ct.ThrowIfCancellationRequested();
                            item.Selected = false;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref item);
                    }
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref items);
        }
    }

    private static void RequireMatchingItem(Excel.SlicerItems items, HashSet<string> selectedItems, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();
        int count = items.Count;
        for (int index = 1; index <= count; index++)
        {
            ct.ThrowIfCancellationRequested();
            Excel.SlicerItem? item = null;
            try
            {
                item = items.Item[index];
                if (selectedItems.Contains(item.Name))
                    return;
            }
            finally
            {
                ComUtilities.Release(ref item);
            }
        }

        throw new ArgumentException("None of the requested slicer items were found. Use item names from availableItems.", nameof(selectedItems));
    }
}
