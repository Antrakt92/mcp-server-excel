using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query LoadTo operations - STANDALONE implementation.
/// Uses ListObjects.Add() pattern (same as Create) for consistency.
/// Based on Microsoft WorkbookQuery and ListObject API.
/// </summary>
public partial class PowerQueryCommands
{
    /// <summary>
    /// Applies load destination to an existing Power Query.
    /// Uses ListObjects.Add() for worksheet loading (consistent with Create).
    /// </summary>
    /// <remarks>
    /// Microsoft Docs Reference:
    /// - ListObjects.Add method - Creates Excel Table with external data source
    /// - QueryTable properties - Configure refresh behavior and formatting
    /// - Connections.Add2 method - Load to Data Model with CreateModelConnection=true
    ///
    /// IMPORTANT: Uses ListObjects.Add() (not QueryTables.Add()) for worksheet loading.
    /// This is the CORRECT approach per Microsoft docs and matches Create() behavior.
    /// </remarks>
    public PowerQueryLoadResult LoadTo(
        IExcelBatch batch,
        string queryName,
        PowerQueryLoadMode loadMode,
        string? targetSheet = null,
        string? targetCellAddress = null)
    {
        if (!Enum.IsDefined(loadMode))
            throw new ArgumentOutOfRangeException(nameof(loadMode));

        bool requiresWorksheet = loadMode == PowerQueryLoadMode.LoadToTable || loadMode == PowerQueryLoadMode.LoadToBoth;

        if (requiresWorksheet && string.IsNullOrWhiteSpace(targetSheet))
        {
            targetSheet = queryName; // Default to query name
        }

        if (!string.IsNullOrWhiteSpace(targetCellAddress) && !requiresWorksheet)
        {
            throw new ArgumentException("targetCellAddress is only supported when loadMode is 'LoadToTable' or 'LoadToBoth'.", nameof(targetCellAddress));
        }

        targetCellAddress ??= "A1"; // Default cell address

        var timeout = ComInteropConstants.DataOperationTimeout;

        var started = System.Diagnostics.Stopwatch.GetTimestamp();
        return ExcelBatchWait.Execute(batch, timeout, (ctx, ct) =>
        {
            Excel.Queries? queries = null;
            Excel.WorkbookQuery? query = null;
            bool destinationsChanging = false;
            var result = new PowerQueryLoadResult
            {
                FilePath = batch.WorkbookPath,
                QueryName = queryName,
                LoadDestination = loadMode,
                WorksheetName = requiresWorksheet ? targetSheet : null,
                TargetCellAddress = requiresWorksheet ? targetCellAddress : null
            };

            try
            {
                // STEP 1: Find the Power Query
                queries = ctx.Book.Queries;
                query = null;
                for (int i = 1; i <= queries.Count; i++)
                {
                    Excel.WorkbookQuery? q = null;
                    try
                    {
                        q = queries.Item(i);
                        string qName = q.Name?.ToString() ?? "";
                        if (qName.Equals(queryName, StringComparison.OrdinalIgnoreCase))
                        {
                            query = q;
                            q = null; // Don't release - keeping reference
                            break;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref q!);
                    }
                }

                if (query == null)
                {
                    throw new InvalidOperationException($"Query '{queryName}' not found.");
                }

                if (requiresWorksheet)
                    PreflightWorksheetLoad(ctx.Book, queryName, targetSheet!, targetCellAddress, ct);

                // STEP 2: Always clean up current destinations before applying the new load mode.
                // This ensures clean state transitions in all directions:
                // LoadToTable → DataModel (removes old ListObject + worksheet connection)
                // DataModel → LoadToTable (removes old Data Model connection)
                // LoadToBoth → ConnectionOnly (removes both destinations)
                // etc.
                destinationsChanging = true;
                UnloadFromDestinations(ctx.Book, queryName);

                // STEP 3: Apply load destination based on mode
                switch (loadMode)
                {
                    case PowerQueryLoadMode.LoadToTable:
                        LoadQueryToWorksheet(ctx.Book, queryName, targetSheet!, targetCellAddress, result, ct);
                        break;

                    case PowerQueryLoadMode.LoadToDataModel:
                        LoadQueryToDataModel(ctx.Book, queryName, result, cancellationToken: ct);
                        break;

                    case PowerQueryLoadMode.LoadToBoth:
                        // For LoadToBoth, create TWO separate properly-named connections:
                        // 1. Worksheet connection: "Query - {name}" (created by LoadQueryToWorksheet)
                        // 2. Data Model connection: "Query - {name} (Data Model)" (with suffix to avoid conflict)
                        LoadQueryToWorksheet(ctx.Book, queryName, targetSheet!, targetCellAddress, result, ct);
                        LoadQueryToDataModel(ctx.Book, queryName, result, " (Data Model)", ct);
                        break;

                    case PowerQueryLoadMode.ConnectionOnly:
                        // UnloadFromDestinations already called above; just set result properties.
                        result.DataRefreshed = false;
                        result.RowsLoaded = 0;
                        result.TargetCellAddress = null;
                        result.Success = true;
                        break;
                }

                // Set additional result properties
                if (result.Success)
                {
                    result.ConfigurationApplied = true;
                    result.DataRefreshed = (loadMode != PowerQueryLoadMode.ConnectionOnly);
                }

                if (loadMode != PowerQueryLoadMode.ConnectionOnly)
                    ReadLoadMetrics(ctx.Book, queryName, result, ct);
                result.DurationMs = System.Diagnostics.Stopwatch.GetElapsedTime(started).TotalMilliseconds;
                return result;
            }
            catch (Exception ex) when (destinationsChanging && !ct.IsCancellationRequested && CanReportMutationFailure(ex))
            {
                throw CreateMutationFailure(ex, queryName, queryCreated: false,
                    state => InspectMutationState(ctx.Book, state, ct), loadDestinationChanged: true);
            }
            finally
            {
                ComUtilities.Release(ref query!);
                ComUtilities.Release(ref queries!);
            }
        });
    }

    private static void PreflightWorksheetLoad(Excel.Workbook workbook, string queryName, string sheetName,
        string targetCellAddress, CancellationToken cancellationToken)
    {
        Excel.Sheets? worksheets = null;
        Excel.Worksheet? sheet = null;
        Excel.Range? destination = null;
        try
        {
            worksheets = workbook.Worksheets;
            sheet = FindLoadWorksheet(worksheets, sheetName, cancellationToken);
            if (sheet is null)
            {
                SheetCommands.ValidateNewSheetName(sheetName);
                SheetCommands.EnsureSheetNameAvailable(workbook, sheetName, cancellationToken);
                if (workbook.ProtectStructure)
                    throw new InvalidOperationException($"Workbook structure is protected. Worksheet '{sheetName}' cannot be added.");
                // The worksheet does not exist yet, so only a bounded cell address can be resolved on it.
                Range.RangeCommands.ParseCellAddress(targetCellAddress);
                return;
            }

            destination = ResolveWorksheetLoadCell(sheet, targetCellAddress);
            ValidateWorksheetLoadCell(sheet, destination, targetCellAddress, queryName, cancellationToken);
        }
        finally
        {
            ComUtilities.Release(ref destination);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref worksheets);
        }
    }

    private static Excel.Worksheet? FindLoadWorksheet(
        Excel.Sheets worksheets, string sheetName, CancellationToken cancellationToken)
    {
        int count = worksheets.Count;
        for (int index = 1; index <= count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = (Excel.Worksheet)worksheets[index];
                if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    var match = sheet;
                    sheet = null;
                    return match;
                }
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        }

        return null;
    }

    private static Excel.Range ResolveWorksheetLoadCell(Excel.Worksheet sheet, string targetCellAddress)
    {
        Excel.Range? destination = null;
        Excel.Worksheet? destinationSheet = null;
        try
        {
            destination = sheet.Range[targetCellAddress];
            destinationSheet = destination.Worksheet;
            if (Convert.ToInt64((object)destination.CountLarge, CultureInfo.InvariantCulture) != 1 ||
                !ReferenceEquals(destinationSheet, sheet))
            {
                throw new ArgumentException(
                    $"Target cell '{targetCellAddress}' must be one cell on the destination worksheet.",
                    nameof(targetCellAddress));
            }

            var cell = destination;
            destination = null;
            return cell;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException(
                $"Could not resolve target cell '{targetCellAddress}' on the destination worksheet: {ex.Message}", ex);
        }
        finally
        {
            ComUtilities.Release(ref destinationSheet);
            ComUtilities.Release(ref destination);
        }
    }

    private static void ValidateWorksheetLoadCell(Excel.Worksheet sheet, Excel.Range destination,
        string targetCellAddress, string? replacingQueryName, CancellationToken cancellationToken)
    {
        if (sheet.ProtectContents && !sheet.ProtectionMode)
        {
            throw new InvalidOperationException(
                $"Worksheet '{sheet.Name}' is protected. Unprotect it or choose another destination before loading the query.");
        }

        Excel.ListObjects? tables = null;
        try
        {
            int row = destination.Row;
            int column = destination.Column;
            tables = sheet.ListObjects;
            int count = tables.Count;
            for (int index = 1; index <= count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.ListObject? table = null;
                Excel.Range? range = null;
                Excel.QueryTable? queryTable = null;
                try
                {
                    table = tables[index];
                    range = table.Range;
                    var (rows, columns) = Range.RangeHelpers.GetDimensions(range);
                    if (row < range.Row || row >= range.Row + rows ||
                        column < range.Column || column >= range.Column + columns)
                        continue;

                    // Allow an existing table only when unload will remove this exact query's table.
                    if (replacingQueryName is not null &&
                        table.SourceType is Excel.XlListObjectSourceType.xlSrcExternal or Excel.XlListObjectSourceType.xlSrcQuery)
                    {
                        queryTable = table.QueryTable;
                        if (queryTable is not null && MatchesWorksheetQueryTable(queryTable, replacingQueryName))
                            return;
                    }

                    throw new InvalidOperationException(
                        $"Cell {targetCellAddress} on sheet '{sheet.Name}' overlaps with an existing table.");
                }
                finally
                {
                    ComUtilities.Release(ref queryTable);
                    ComUtilities.Release(ref range);
                    ComUtilities.Release(ref table);
                }
            }

            if ((object)destination.HasFormula is not bool hasFormula)
                throw new InvalidOperationException($"Could not inspect formulas in target cell '{targetCellAddress}'.");

            if (hasFormula || (object?)destination.Value2 is not null)
            {
                throw new InvalidOperationException(
                    $"Target cell '{targetCellAddress}' on worksheet '{sheet.Name}' already contains data. " +
                    "Choose a different targetCellAddress or clear the existing data first.");
            }
        }
        finally
        {
            ComUtilities.Release(ref tables);
        }
    }

    /// <summary>
    /// Loads query data to a worksheet using ListObjects.Add (correct approach for Power Query).
    /// SHARED IMPLEMENTATION - Used by both Create and LoadTo.
    /// </summary>
    /// <remarks>
    /// This is extracted from Create.cs for reuse. Both Create and LoadTo should use
    /// the same ListObjects.Add() pattern for consistency.
    /// Matches Excel UI behavior: Creates worksheet if it doesn't exist, or loads to existing worksheet.
    /// </remarks>
    private static bool LoadQueryToWorksheet(
        Excel.Workbook workbook,
        string queryName,
        string sheetName,
        string targetCellAddress,
        PowerQueryOperationResult result,
        CancellationToken cancellationToken)
    {
        Excel.Sheets? worksheets = null;
        Excel.Worksheet? sheet = null;
        Excel.Range? destination = null;
        Excel.Connections? connections = null;
        Excel.WorkbookConnection? connection = null;
        Excel.ListObjects? listObjects = null;
        Excel.ListObject? listObject = null;
        Excel.QueryTable? queryTable = null;
        Excel.Range? listObjectRange = null;

        try
        {
            worksheets = workbook.Worksheets;

            sheet = FindLoadWorksheet(worksheets, sheetName, cancellationToken);
            bool worksheetExists = sheet is not null;
            if (sheet is null)
            {
                sheet = (Excel.Worksheet)worksheets.Add();
                sheet.Name = sheetName;
            }

            destination = ResolveWorksheetLoadCell(sheet, targetCellAddress);
            if (worksheetExists)
                ValidateWorksheetLoadCell(sheet, destination, targetCellAddress, null, cancellationToken);

            // This ensures the connection is named "Query - {queryName}" instead of generic "Connection", "Connection1", etc.
            connections = workbook.Connections;
            string connectionName = $"Query - {queryName}";
            string connectionDescription = $"Connection to the '{queryName}' query in the workbook.";
            string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
            string commandText = $"SELECT * FROM [{queryName}]";

            connection = connections.Add2(
                Name: connectionName,
                Description: connectionDescription,
                ConnectionString: connectionString,
                CommandText: commandText,
                lCmdtype: 2,                    // xlCmdSql
                CreateModelConnection: false,   // Worksheet loading, NOT Data Model
                ImportRelationships: false
            );

            // Step 2: Add ListObject using the connection object (not connection string)
            // This reuses the properly-named connection instead of creating a new generic one
            listObjects = sheet.ListObjects;
            listObject = listObjects.Add(
                Excel.XlListObjectSourceType.xlSrcExternal,
                connection,         // Source: connection object (reuses existing named connection)
                Type.Missing,       // LinkSource
                Excel.XlYesNoGuess.xlYes,
                destination         // Destination: starting cell
            );

            // Configure the QueryTable behind the ListObject
            queryTable = listObject.QueryTable;
            queryTable.CommandType = Excel.XlCmdType.xlCmdSql;
            queryTable.CommandText = $"SELECT * FROM [{queryName}]";
            queryTable.AdjustColumnWidth = true;
            queryTable.PreserveFormatting = true;
            queryTable.BackgroundQuery = false; // Synchronous
            queryTable.RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells;
            // Excel needs column bindings to remove fields when the source schema shrinks.
            queryTable.PreserveColumnInfo = true;

            // Refresh to materialize the table.
            // Do NOT use EnterLongOperation here: synchronous QueryTable refresh depends on inbound
            // Excel callbacks to complete. Rejecting those callbacks can deadlock the load.
            QueryTableRefreshHelper.RefreshSynchronously(
                queryTable,
                cancellationToken,
                $"Power Query load for '{queryName}'");

            // Name the table after the query for predictable M-code referencing.
            // Without this, Excel auto-assigns a generic name ("Table1", "Table2", etc.)
            // which makes Excel.CurrentWorkbook(){[Name=queryName]}[Content] lookups unreliable.
            string createdTableName = listObject.Name;
            try
            {
                listObject.Name = queryName;
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException(
                    $"Query '{queryName}' loaded into table '{createdTableName}' on worksheet '{sheetName}', " +
                    $"but renaming the table failed: {ex.Message}. The loaded table remains; no changes were rolled back.", ex);
            }

            // Capture results - use ListObject Range for total rows, subtract header
            listObjectRange = listObject.Range;
            int totalRows = listObjectRange is not null ? Range.RangeHelpers.GetDimensions(listObjectRange).Rows : 0;
            result.RowsLoaded = totalRows > 0 ? totalRows - 1 : 0; // Subtract header row
            result.Success = true;

            return true;
        }
        finally
        {
            ComUtilities.Release(ref listObjectRange);
            ComUtilities.Release(ref queryTable);
            ComUtilities.Release(ref listObject);
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref connection);
            ComUtilities.Release(ref connections);
            ComUtilities.Release(ref destination);
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref worksheets);
        }
    }

    /// <summary>
    /// Loads query data to the Data Model using Connections.Add2.
    /// SHARED IMPLEMENTATION - Used by both Create and LoadTo.
    /// </summary>
    /// <param name="workbook">The workbook to load into.</param>
    /// <param name="queryName">The Power Query name.</param>
    /// <param name="result">Result object to populate.</param>
    /// <param name="connectionNameSuffix">Optional suffix for connection name to avoid conflicts (e.g., " (Data Model)").</param>
    private static bool LoadQueryToDataModel(
        Excel.Workbook workbook,
        string queryName,
        PowerQueryOperationResult result,
        string? connectionNameSuffix = null,
        CancellationToken cancellationToken = default)
    {
        Excel.Connections? connections = null;
        Excel.WorkbookConnection? connection = null;

        try
        {
            connections = workbook.Connections;

            // Use suffix if provided (for LoadToBoth to avoid conflict with worksheet connection)
            string connectionName = string.IsNullOrEmpty(connectionNameSuffix)
                ? $"Query - {queryName}"
                : $"Query - {queryName}{connectionNameSuffix}";
            string description = $"Connection to the '{queryName}' query in the workbook.";
            string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
            string commandText = $"\"{queryName}\"";

            connection = connections.Add2(
                Name: connectionName,
                Description: description,
                ConnectionString: connectionString,
                CommandText: commandText,
                lCmdtype: 6, // Data Model command type
                CreateModelConnection: true, // CRITICAL: This loads to Data Model
                ImportRelationships: false
            );

            // Refresh the connection to actually load data into the Data Model.
            // Without this call, the connection is registered but no data is materialized —
            // the table never appears in the Data Model even though success is returned.
            // Do NOT use EnterLongOperation here: synchronous connection.Refresh() can require
            // inbound callbacks from Excel/MashupHost to complete.
            OleMessageFilter.SetPendingCancellationToken(cancellationToken);
            try
            {
                connection.Refresh();
            }
            finally
            {
                OleMessageFilter.ClearPendingCancellationToken();
            }

            result.Success = true;
            return true;
        }
        finally
        {
            ComUtilities.Release(ref connection);
            ComUtilities.Release(ref connections);
        }
    }

}
