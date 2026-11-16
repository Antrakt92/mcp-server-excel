using System.Runtime.InteropServices;
using System.Runtime.ExceptionServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query evaluate operation - executes M code and returns results without creating a permanent query
/// </summary>
public partial class PowerQueryCommands
{
    private const string EvaluateResultLimitGuidance =
        "Shorten the M source or limit its output with Table.FirstN or Table.SelectColumns.";

    /// <inheritdoc />
    public PowerQueryEvaluateResult Evaluate(IExcelBatch batch, string mCode)
    {
        var result = new PowerQueryEvaluateResult
        {
            FilePath = batch.WorkbookPath,
            MCode = mCode
        };

        // Validate M code
        if (string.IsNullOrWhiteSpace(mCode))
        {
            throw new ArgumentException("M code is required for evaluate action", nameof(mCode));
        }

        var budget = new RangeReadBudget(EvaluateResultLimitGuidance);
        budget.AddScalar(result.FilePath);
        budget.AddScalar(result.MCode);

        // Generate unique temporary names
        var uniqueId = Guid.NewGuid().ToString("N")[..8];
        var tempQueryName = $"__pq_eval_{uniqueId}";
        var tempSheetName = $"__pq_eval_{uniqueId}";

        using var timeoutCts = new CancellationTokenSource(batch.OperationTimeout);

        PowerQueryEvaluateResult Execute() => batch.Execute(EvaluateTemporaryQuery, timeoutCts.Token);

        [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
            Justification = "Retains the evaluation failure and every temporary-object cleanup failure, releases COM references, then throws the original or aggregate error; no failed evaluation returns success.")]
        PowerQueryEvaluateResult EvaluateTemporaryQuery(ExcelContext ctx, CancellationToken ct)
        {
            Excel.Queries? queriesCollection = null;
            Excel.WorkbookQuery? query = null;
            Excel.Sheets? worksheets = null;
            Excel.Worksheet? tempSheet = null;
            Excel.ListObjects? listObjects = null;
            Excel.ListObject? listObject = null;
            Excel.QueryTable? queryTable = null;
            Excel.Range? range = null;
            Exception? operationError = null;

            try
            {
                ct.ThrowIfCancellationRequested();

                // STEP 1: Create temporary query with the M code
                queriesCollection = ctx.Book.Queries;
                query = queriesCollection.Add(tempQueryName, mCode);

                // STEP 2: Create temporary worksheet and load data to it
                worksheets = ctx.Book.Worksheets;
                tempSheet = (Excel.Worksheet)worksheets.Add();
                tempSheet.Name = tempSheetName;

                // STEP 3: Load query to worksheet using QueryTable
                // Connection string for Power Query - must include Extended Properties
                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={tempQueryName};Extended Properties=\"\"";
                range = tempSheet.Range["A1"];

                listObjects = tempSheet.ListObjects;
                listObject = listObjects.Add(
                    Excel.XlListObjectSourceType.xlSrcExternal,
                    connectionString,   // Source: connection string
                    Type.Missing,       // LinkSource
                    Excel.XlYesNoGuess.xlYes,
                    range               // Destination: starting cell
                );

                // Get the QueryTable to refresh (this executes the M code)
                queryTable = listObject.QueryTable;

                // Configure QueryTable to select from the query
                queryTable.CommandType = Excel.XlCmdType.xlCmdSql;
                queryTable.CommandText = $"SELECT * FROM [{tempQueryName}]";
                queryTable.BackgroundQuery = false; // Synchronous

                // STEP 4: Refresh to execute the M code (errors will throw via QueryTable.Refresh)
                // This path uses the same synchronous worksheet refresh mechanism as PowerQuery refresh/load.
                // Do NOT use EnterLongOperation here: rejecting inbound COM callbacks can deadlock the
                // MashupHost → Excel → STA callback chain that completes the synchronous refresh.
                try
                {
                    QueryTableRefreshHelper.RefreshSynchronously(queryTable, ct, "Power Query evaluation");
                }
                catch (COMException ex) when (TryWrapPowerQueryException(ex, out var queryError))
                {
                    throw queryError!;
                }

                ct.ThrowIfCancellationRequested();

                ReadEvaluationResult(listObject, result, budget, ct);
            }
            catch (Exception ex)
            {
                operationError = ex;
            }

            var cleanupErrors = new List<Exception>();
            try
            {
                if (listObject != null)
                {
                    listObject.Delete();
                }
            }
            catch (COMException ex)
            {
                cleanupErrors.Add(new InvalidOperationException(
                    $"Failed to delete temporary evaluate table for '{tempQueryName}'.",
                    ex));
            }

            ComUtilities.Release(ref queryTable);
            ComUtilities.Release(ref listObject);
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref range);

            try
            {
                if (tempSheet != null)
                {
                    // Suppress alerts to avoid "Are you sure you want to delete?" prompt
                    Excel.Application? app = null;
                    try
                    {
                        app = ctx.Book.Application;
                        bool originalAlerts = app.DisplayAlerts;
                        app.DisplayAlerts = false;
                        try
                        {
                            tempSheet.Delete();
                        }
                        finally
                        {
                            app.DisplayAlerts = originalAlerts;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref app);
                    }
                }
            }
            catch (COMException ex)
            {
                cleanupErrors.Add(new InvalidOperationException(
                    $"Failed to delete temporary evaluate worksheet '{tempSheetName}'.",
                    ex));
            }

            ComUtilities.Release(ref tempSheet);
            ComUtilities.Release(ref worksheets);

            try
            {
                PowerQuery.PowerQueryHelpers.RemoveConnectionsByMashupLocation(
                    ctx.Book,
                    tempQueryName);
            }
            catch (Exception ex) when (ex is COMException or InvalidOperationException)
            {
                cleanupErrors.Add(new InvalidOperationException(
                    $"Failed to delete temporary evaluate connections for '{tempQueryName}'.",
                    ex));
            }

            try
            {
                if (query != null)
                {
                    query.Delete();
                }
            }
            catch (COMException ex)
            {
                cleanupErrors.Add(new InvalidOperationException(
                    $"Failed to delete temporary evaluate query '{tempQueryName}'.",
                    ex));
            }

            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queriesCollection);

            if (cleanupErrors.Count > 0)
            {
                if (operationError != null)
                {
                    cleanupErrors.Insert(0, operationError);
                }

                var cleanupFailure = new PowerQueryCommandException(
                    $"Power Query evaluate cleanup failed for temporary identity '{tempQueryName}'. " +
                    "Inspect powerquery list and connection list before saving, or close without saving and retry.",
                    "Cleanup",
                    new AggregateException(cleanupErrors));
                if (operationError is OperationCanceledException canceled)
                    throw new OperationCanceledException(cleanupFailure.Message, cleanupFailure, canceled.CancellationToken);
                throw cleanupFailure;
            }

            if (operationError != null)
            {
                ExceptionDispatchInfo.Capture(operationError).Throw();
            }

            result.Success = true;
            return result;
        }

        try
        {
            return Execute();
        }
        catch (OperationCanceledException ex) when (timeoutCts.IsCancellationRequested)
        {
            throw new TimeoutException(
                $"Power Query evaluate timed out after {batch.OperationTimeout.TotalSeconds} seconds.",
                ex);
        }
    }

    private static void ReadEvaluationResult(
        Excel.ListObject table, PowerQueryEvaluateResult result, RangeReadBudget budget, CancellationToken ct)
    {
        Excel.Range? dataRange = null;
        Excel.Range? headerRange = null;
        try
        {
            dataRange = table.DataBodyRange;
            var dataSize = dataRange is null ? default : RangeHelpers.GetDimensions(dataRange);
            RangeReadWindow.EnsureCellLimit(dataSize.Rows, dataSize.Columns, RangeReadWindow.MaximumReadCells,
                EvaluateResultLimitGuidance + " No values were read.");

            headerRange = table.HeaderRowRange;
            if (headerRange is not null)
            {
                var size = RangeHelpers.GetDimensions(headerRange);
                var header = BoundedRangeValueReader.Read(headerRange, size, budget, ct,
                    convertValue: static value => value?.ToString())[0];
                if (header.Count != 1 || header[0] is not null)
                {
                    for (int index = 0; index < header.Count; index++)
                    {
                        ct.ThrowIfCancellationRequested();
                        var name = (string?)header[index];
                        if (name is null)
                        {
                            name = $"Column{index + 1}";
                            budget.AddScalar(name);
                        }
                        result.Columns.Add(name);
                    }
                }
            }
            result.ColumnCount = result.Columns.Count;

            if (dataRange is not null)
            {
                // Excel can return a single blank placeholder for an empty table.
                bool blankSingleton = false;
                result.Rows = BoundedRangeValueReader.Read(dataRange, dataSize, budget, ct,
                    convertValue: dataSize == (1, 1) ? ConvertSingleton : ConvertCellValue);
                if (blankSingleton)
                    result.Rows.Clear();

                object? ConvertSingleton(object? value)
                {
                    blankSingleton = value is null;
                    return ConvertCellValue(value);
                }
            }
            result.RowCount = result.Rows.Count;
        }
        finally
        {
            ComUtilities.Release(ref headerRange);
            ComUtilities.Release(ref dataRange);
        }
    }

    /// <summary>
    /// Converts Excel cell values to JSON-friendly types
    /// </summary>
    private static object? ConvertCellValue(object? value)
    {
        if (value == null || value == DBNull.Value)
            return null;

        return value switch
        {
            DateTime dt => dt.ToString("O"), // ISO 8601 format
            double d => d,
            int i => i,
            long l => l,
            bool b => b,
            string s => s,
            _ => value.ToString()
        };
    }
}
