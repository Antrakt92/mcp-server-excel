using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal static class WorkbookRefreshProbe
{
    private const int ExcelObjectRequired = unchecked((int)0x800A01A8);

    internal static WorkbookRefreshState Read(Excel.Application application, Excel.Workbook workbook)
    {
        if (!application.Ready || application.CalculationState == Excel.XlCalculationState.xlCalculating)
        {
            return WorkbookRefreshState.Busy;
        }

        var connections = ReadConnectionRefreshState(workbook);
        return connections == WorkbookRefreshState.Refreshing || HasRefreshingQueryTable(workbook)
            ? WorkbookRefreshState.Refreshing
            : connections;
    }

    private static WorkbookRefreshState ReadConnectionRefreshState(Excel.Workbook workbook)
    {
        var state = WorkbookRefreshState.Ready;
        Excel.Connections? connections = null;
        try
        {
            connections = workbook.Connections;
            for (var i = 1; i <= connections.Count; i++)
            {
                Excel.WorkbookConnection? connection = null;
                Excel.OLEDBConnection? oledb = null;
                Excel.ODBCConnection? odbc = null;
                Excel.DataFeedConnection? dataFeed = null;
                try
                {
                    connection = connections.Item(i);
                    if (connection.Type == Excel.XlConnectionType.xlConnectionTypeOLEDB)
                    {
                        oledb = connection.OLEDBConnection;
                        var observed = ReadOleDbConnection(oledb);
                        if (observed == WorkbookRefreshState.Refreshing) return observed;
                        if (observed == WorkbookRefreshState.Unknown) state = observed;
                    }
                    else if (connection.Type == Excel.XlConnectionType.xlConnectionTypeODBC)
                    {
                        odbc = connection.ODBCConnection;
                        if (odbc.Refreshing) return WorkbookRefreshState.Refreshing;
                    }
                    else if (connection.Type == Excel.XlConnectionType.xlConnectionTypeDATAFEED)
                    {
                        dataFeed = connection.DataFeedConnection;
                        if (dataFeed.Refreshing) return WorkbookRefreshState.Refreshing;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref dataFeed);
                    ComUtilities.Release(ref odbc);
                    ComUtilities.Release(ref oledb);
                    ComUtilities.Release(ref connection);
                }
            }
            return state;
        }
        finally
        {
            ComUtilities.Release(ref connections);
        }
    }

    internal static WorkbookRefreshState ReadOleDbConnection(Excel.OLEDBConnection connection)
    {
        try
        {
            return connection.Refreshing ? WorkbookRefreshState.Refreshing : WorkbookRefreshState.Ready;
        }
        catch (COMException ex) when (ex.HResult == ExcelObjectRequired)
        {
            // Excel can deny this getter while a background refresh is settling.
            return WorkbookRefreshState.Unknown;
        }
    }

    private static bool HasRefreshingQueryTable(Excel.Workbook workbook)
    {
        Excel.Sheets? sheets = null;
        try
        {
            sheets = workbook.Worksheets;
            for (var i = 1; i <= sheets.Count; i++)
            {
                Excel.Worksheet? sheet = null;
                Excel.QueryTables? queries = null;
                Excel.ListObjects? tables = null;
                try
                {
                    sheet = (Excel.Worksheet)sheets.Item[i];
                    queries = sheet.QueryTables;
                    for (var j = 1; j <= queries.Count; j++)
                    {
                        Excel.QueryTable? query = null;
                        try
                        {
                            query = queries.Item(j);
                            if (query.Refreshing) return true;
                        }
                        finally
                        {
                            ComUtilities.Release(ref query);
                        }
                    }

                    // Power Query tables can be absent from Worksheet.QueryTables.
                    tables = sheet.ListObjects;
                    for (var j = 1; j <= tables.Count; j++)
                    {
                        Excel.ListObject? table = null;
                        Excel.QueryTable? query = null;
                        try
                        {
                            table = tables[j];
                            if (table.SourceType is Excel.XlListObjectSourceType.xlSrcQuery
                                or Excel.XlListObjectSourceType.xlSrcExternal)
                            {
                                query = table.QueryTable;
                                if (query.Refreshing) return true;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref query);
                            ComUtilities.Release(ref table);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref tables);
                    ComUtilities.Release(ref queries);
                    ComUtilities.Release(ref sheet);
                }
            }
            return false;
        }
        finally
        {
            ComUtilities.Release(ref sheets);
        }
    }
}
