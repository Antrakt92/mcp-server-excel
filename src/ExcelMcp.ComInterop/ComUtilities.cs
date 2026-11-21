using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// Low-level COM interop utilities for Excel automation.
/// Provides helpers for finding Excel objects and managing COM object lifecycle.
/// </summary>
public static class ComUtilities
{
    internal static ADODB._Connection? ReadAdoConnection(Excel.ModelConnection connection) =>
        TakeAdoConnection((object?)connection.ADOConnection);

    internal static ADODB._Connection? ReadAdoConnection(Excel.OLEDBConnection connection) =>
        TakeAdoConnection((object?)connection.ADOConnection);

    private static ADODB._Connection? TakeAdoConnection(object? ownedConnection)
    {
        try
        {
            var connection = (ADODB._Connection?)ownedConnection;
            ownedConnection = null;
            return connection;
        }
        finally
        {
            Release(ref ownedConnection);
        }
    }

    internal static int? ReadStyleIndentLevel(Excel.Style style) =>
        ReadStyleIndentLevel(() => style.IndentLevel);

    internal static int? ReadStyleIndentLevel(Func<int> readIndentLevel)
    {
        try
        {
            return readIndentLevel();
        }
        catch (InvalidCastException ex) when (ex.HResult == unchecked((int)0x80004002))
        {
            // Only this getter: a new style's unset/initial zero is VT_NULL, but the PIA requires int.
            return null;
        }
    }

    /// <summary>
    /// Safely releases a COM object and sets the reference to null
    /// </summary>
    /// <param name="comObject">The COM object to release</param>
    /// <remarks>
    /// Use this helper to release intermediate COM objects (like ranges, worksheets, queries)
    /// to prevent Excel process from staying open. This is especially important when
    /// iterating through collections or accessing multiple COM properties.
    /// </remarks>
    /// <example>
    /// <code>
    /// Excel.Queries? queries = null;
    /// try
    /// {
    ///     queries = workbook.Queries;
    ///     // Use queries...
    /// }
    /// finally
    /// {
    ///     ComUtilities.Release(ref queries);
    /// }
    /// </code>
    /// </example>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ApiDesign", "RS0030",
        Justification = "The central COM release helper owns exactly one acquired RCW reference, checks non-COM values, and clears the caller's reference in finally. Callers must use this helper.")]
    public static void Release<T>(ref T? comObject) where T : class
    {
        if (comObject != null)
        {
            try
            {
                if (Marshal.IsComObject(comObject))
                {
                    Marshal.ReleaseComObject(comObject);
                }
            }
            catch (InvalidComObjectException)
            {
                // An already released RCW has no remaining reference to release.
                return;
            }
            finally
            {
                comObject = null;
            }
        }
    }

    /// <summary>
    /// Finds a worksheet by name in the workbook
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="sheetName">Name of the worksheet to find</param>
    /// <returns>The worksheet COM object if found, null otherwise</returns>
    /// <remarks>
    /// CRITICAL: Caller is responsible for releasing the returned COM object.
    /// Use ComUtilities.Release(ref sheet) when done with the object.
    /// </remarks>
    public static Excel.Worksheet? FindSheet(Excel.Workbook workbook, string sheetName)
    {
        Excel.Sheets? sheetsCollection = null;
        try
        {
            sheetsCollection = workbook.Worksheets;
            int count = sheetsCollection.Count;
            for (int i = 1; i <= count; i++)
            {
                Excel.Worksheet? sheet = null;
                try
                {
                    sheet = (Excel.Worksheet)sheetsCollection[i];
                    string currentName = sheet.Name;

                    if (currentName == sheetName)
                    {
                        // Found match - return it (caller owns it now)
                        var result = sheet;
                        sheet = null; // Prevent cleanup in finally block
                        return result;
                    }
                }
                finally
                {
                    // Only release if not returning
                    if (sheet != null)
                    {
                        Release(ref sheet);
                    }
                }
            }

            return null; // Not found
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to search for worksheet '{sheetName}'.", ex);
        }
        finally
        {
            Release(ref sheetsCollection);
        }
    }

    /// <summary>
    /// Safely iterates through all columns in a model table with automatic COM cleanup
    /// </summary>
    /// <param name="table">Model table COM object</param>
    /// <param name="action">Action to perform on each column (receives the column COM object and 1-based index)</param>
    /// <remarks>
    /// The column is passed as <see cref="object"/> rather than <c>Excel.ModelTableColumn</c> on purpose:
    /// the Excel interop types are embedded (EmbedInteropTypes), and an embedded interop type cannot be used
    /// as a generic type argument across assembly boundaries (compiler error CS1769). Callers cast the
    /// object to <c>Excel.ModelTableColumn</c> before reading its properties.
    /// </remarks>
    public static void ForEachColumn(Excel.ModelTable table, Action<object, int> action)
    {
        Excel.ModelTableColumns? columns = null;
        try
        {
            columns = table.ModelTableColumns;
            int count = columns.Count;

            for (int i = 1; i <= count; i++)
            {
                Excel.ModelTableColumn? column = null;
                try
                {
                    column = columns.Item(i);
                    action(column, i);
                }
                finally
                {
                    Release(ref column);
                }
            }
        }
        finally
        {
            Release(ref columns);
        }
    }

    [DllImport("kernel32.dll")]
    private static extern void Sleep(uint dwMilliseconds);

    /// <summary>
    /// Kernel-level sleep that does NOT pump the STA COM message queue.
    /// Unlike Thread.Sleep (which uses CoWaitForMultipleHandles internally and wakes early on
    /// every incoming COM event), this calls Win32 Sleep() directly via NtDelayExecution —
    /// the thread genuinely sleeps for the full interval regardless of COM callbacks.
    /// Safe to use in WaitForRefreshCompletion: Power Query refresh completion is driven by
    /// Excel's own internals (MashupHost.exe → Excel's STA). Our polling thread does not need
    /// to service any callbacks for connection.Refreshing to become false.
    /// </summary>
    public static void KernelSleep(int milliseconds) =>
        Sleep((uint)Math.Max(0, milliseconds));
}
