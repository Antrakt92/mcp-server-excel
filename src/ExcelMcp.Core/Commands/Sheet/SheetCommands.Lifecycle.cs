using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet lifecycle operations (List, Create, Rename, Copy, Delete)
/// </summary>
public partial class SheetCommands
{
    /// <inheritdoc />
    public WorksheetListResult List(IExcelBatch batch, string? filePath = null)
    {
        var result = new WorksheetListResult { FilePath = filePath ?? batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            // Get the workbook to list from
            Excel.Workbook workbook = filePath != null ? batch.GetWorkbook(filePath) : ctx.Book;

            Excel.Sheets? sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int i = 1; i <= sheets.Count; i++)
                {
                    Excel.Worksheet? sheet = null;
                    try
                    {
                        sheet = (Excel.Worksheet)sheets[i];
                        result.Worksheets.Add(new WorksheetInfo
                        {
                            Name = sheet.Name,
                            Index = i,
                            Visible = Convert.ToInt32((object)sheet.Visible, System.Globalization.CultureInfo.InvariantCulture) == -1
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref sheet);
                    }
                }
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref sheets);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Create(IExcelBatch batch, string sheetName, string? filePath = null)
    {
        ValidateNewSheetName(sheetName);
        return batch.Execute((ctx, ct) =>
        {
            // Get the workbook to create sheet in
            Excel.Workbook workbook = filePath != null ? batch.GetWorkbook(filePath) : ctx.Book;

            Excel.Sheets? sheets = null;
            Excel.Worksheet? newSheet = null;
            try
            {
                EnsureSheetNameAvailable(workbook, sheetName, ct);
                sheets = workbook.Worksheets;
                newSheet = (Excel.Worksheet)sheets.Add();
                var addedSheetCurrentName = newSheet.Name;
                SetSheetNameWithContext(newSheet, sheetName, addedSheetCurrentName, isNewSheet: true);
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref newSheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Rename(IExcelBatch batch, string oldName, string newName)
    {
        ValidateNewSheetName(newName);
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, oldName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{oldName}' not found.");
                }
                EnsureSheetNameAvailable(ctx.Book, newName, ct, sheet.Name);
                SetSheetNameWithContext(sheet, newName, sheet.Name, isNewSheet: false);
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Copy(IExcelBatch batch, string sourceName, string targetName)
    {
        ValidateNewSheetName(targetName);
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sourceSheet = null;
            Excel.Sheets? sheets = null;
            Excel.Worksheet? lastSheet = null;
            try
            {
                sourceSheet = ComUtilities.FindSheet(ctx.Book, sourceName);
                if (sourceSheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sourceName}' not found.");
                }
                EnsureSheetNameAvailable(ctx.Book, targetName, ct);
                sheets = ctx.Book.Worksheets;
                lastSheet = (Excel.Worksheet)sheets.Item[sheets.Count];
                CopyWorksheet(sourceSheet, ctx.Book, lastSheet, before: false, targetName, ct);
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref lastSheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref sourceSheet);
            }
        });
    }

    private static void CopyWorksheet(
        Excel.Worksheet source, Excel.Workbook destination, object anchor,
        bool before, string? targetName, CancellationToken cancellationToken)
    {
        Excel.Sheets? worksheets = null;
        Excel.Worksheet? copied = null;
        var existing = new HashSet<Excel.Worksheet>(ReferenceEqualityComparer.Instance);
        try
        {
            if (targetName is not null)
            {
                // Excel can insert after the last visible tab when the requested anchor is hidden.
                // Keep the original worksheet identities alive across the native copy.
                worksheets = destination.Worksheets;
                int count = worksheets.Count;
                for (int index = 1; index <= count; index++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    Excel.Worksheet? worksheet = null;
                    try
                    {
                        worksheet = (Excel.Worksheet)worksheets[index];
                        if (!existing.Add(worksheet))
                            throw new InvalidOperationException("Excel returned a worksheet twice while preparing the copy.");
                        worksheet = null;
                    }
                    finally { ComUtilities.Release(ref worksheet); }
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            if (before)
                source.Copy(Before: anchor);
            else
                source.Copy(After: anchor);

            if (worksheets is not null && targetName is not null)
            {
                string copiedName;
                try
                {
                    copied = FindCopiedWorksheet(worksheets, existing);
                    copiedName = copied.Name;
                }
                catch (Exception ex)
                {
                    throw new OperationCompletionUnconfirmedException(
                        "Could not identify the copied worksheet. Inspect the destination workbook before retrying.", ex);
                }
                SetSheetNameWithContext(copied, targetName, copiedName, isNewSheet: true);
            }
        }
        finally
        {
            ComUtilities.Release(ref copied);
            foreach (var original in existing.Reverse())
            {
                Excel.Worksheet? worksheet = original;
                ComUtilities.Release(ref worksheet);
            }
            ComUtilities.Release(ref worksheets);
        }
    }

    private static Excel.Worksheet FindCopiedWorksheet(
        Excel.Sheets worksheets, HashSet<Excel.Worksheet> existing)
    {
        Excel.Worksheet? copied = null;
        try
        {
            int count = worksheets.Count;
            if (count != existing.Count + 1)
                throw new InvalidOperationException("Excel did not add exactly one worksheet.");

            for (int index = 1; index <= count; index++)
            {
                Excel.Worksheet? worksheet = null;
                try
                {
                    worksheet = (Excel.Worksheet)worksheets[index];
                    if (existing.Contains(worksheet))
                        continue;
                    if (copied is not null)
                        throw new InvalidOperationException("Excel returned more than one new worksheet.");
                    copied = worksheet;
                    worksheet = null;
                }
                finally { ComUtilities.Release(ref worksheet); }
            }

            var result = copied ?? throw new InvalidOperationException("Excel did not return the copied worksheet.");
            copied = null;
            return result;
        }
        finally { ComUtilities.Release(ref copied); }
    }

    private static void SetSheetNameWithContext(
        Excel.Worksheet sheet, string name, string existingName, bool isNewSheet)
    {
        try
        {
            sheet.Name = name;
        }
        catch (System.Runtime.InteropServices.COMException namingError)
        {
            var message = isNewSheet
                ? $"Excel rejected worksheet name '{name}'. The new worksheet '{existingName}' remains in the workbook; inspect it and remove it if appropriate."
                : $"Excel rejected worksheet name '{name}'. Worksheet '{existingName}' was not renamed.";
            throw new InvalidOperationException(message, namingError);
        }
    }

    internal static void ValidateNewSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name) ||
            name.Length > 31 ||
            name.AsSpan().IndexOfAny(":\\/?*[]".AsSpan()) >= 0 ||
            name.StartsWith('\'') ||
            name.EndsWith('\'') ||
            string.Equals(name, "History", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException(
                "Worksheet names must be nonblank, contain at most 31 characters, not be 'History', " +
                "not contain : \\ / ? * [ ], and not begin or end with an apostrophe.",
                nameof(name));
        }
    }

    internal static void EnsureSheetNameAvailable(
        Excel.Workbook workbook, string name, CancellationToken cancellationToken, string? currentName = null)
    {
        Excel.Sheets? sheets = null;
        try
        {
            sheets = workbook.Sheets;
            for (var index = 1; index <= sheets.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                object? candidate = null;
                try
                {
                    candidate = sheets.Item[index];
                    var existingName = candidate switch
                    {
                        Excel.Worksheet worksheet => worksheet.Name,
                        Excel.Chart chart => chart.Name,
                        _ => throw new InvalidOperationException("Excel returned an unsupported sheet type during name validation.")
                    };
                    if (string.Equals(existingName, name, StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(existingName, currentName, StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException($"Sheet '{name}' already exists.");
                    }
                }
                finally { ComUtilities.Release(ref candidate); }
            }
        }
        finally { ComUtilities.Release(ref sheets); }
    }

    /// <inheritdoc />
    public OperationResult Delete(IExcelBatch batch, string sheetName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }
                ct.ThrowIfCancellationRequested();
                sheet.Delete();
                // Excel's COM interop drops the Boolean result when the confirmation is cancelled.
                bool originalStillPresent;
                try
                {
                    originalStillPresent = ContainsWorksheet(ctx.Book, sheet);
                }
                catch (Exception ex)
                {
                    throw new OperationCompletionUnconfirmedException(
                        $"Could not confirm that Excel deleted worksheet '{sheetName}'. Inspect the workbook before retrying.", ex);
                }
                if (originalStillPresent)
                {
                    throw new InvalidOperationException($"Excel did not delete worksheet '{sheetName}'. The original worksheet is still present.");
                }
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static bool ContainsWorksheet(Excel.Workbook workbook, Excel.Worksheet target)
    {
        Excel.Sheets? sheets = null;
        try
        {
            sheets = workbook.Worksheets;
            var count = sheets.Count;
            for (var index = 1; index <= count; index++)
            {
                Excel.Worksheet? candidate = null;
                try
                {
                    candidate = (Excel.Worksheet)sheets[index];
                    if (ReferenceEquals(candidate, target))
                    {
                        return true;
                    }
                }
                finally { ComUtilities.Release(ref candidate); }
            }
            return false;
        }
        finally { ComUtilities.Release(ref sheets); }
    }

    /// <inheritdoc />
    public OperationResult Move(IExcelBatch batch, string sheetName, string? beforeSheet = null, string? afterSheet = null)
    {
        // Validate parameters
        if (!string.IsNullOrWhiteSpace(beforeSheet) && !string.IsNullOrWhiteSpace(afterSheet))
        {
            throw new ArgumentException("Cannot specify both beforeSheet and afterSheet");
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Worksheet? targetSheet = null;
            Excel.Sheets? sheets = null;
            object? lastTab = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                // If no position specified, move to end
                if (string.IsNullOrWhiteSpace(beforeSheet) && string.IsNullOrWhiteSpace(afterSheet))
                {
                    sheets = ctx.Book.Sheets;
                    lastTab = sheets[sheets.Count];
                    if (!ReferenceEquals(sheet, lastTab))
                        sheet.Move(After: lastTab);
                }
                else
                {
                    // Find target sheet for positioning
                    string targetName = beforeSheet ?? afterSheet!;
                    targetSheet = ComUtilities.FindSheet(ctx.Book, targetName);
                    if (targetSheet == null)
                    {
                        throw new InvalidOperationException($"Target sheet '{targetName}' not found.");
                    }

                    // Move using Excel COM API
                    if (!string.IsNullOrWhiteSpace(beforeSheet))
                    {
                        sheet.Move(Before: targetSheet);
                    }
                    else
                    {
                        sheet.Move(After: targetSheet);
                    }
                }
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref lastTab);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref targetSheet);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    // === ATOMIC CROSS-FILE OPERATIONS ===

    /// <inheritdoc />
    public OperationResult CopyToFile(string sourceFile, string sourceSheet, string targetFile, string? targetSheetName = null, string? beforeSheet = null, string? afterSheet = null)
    {
        if (targetSheetName is not null)
        {
            ValidateNewSheetName(targetSheetName);
        }

        // Validate positioning parameters
        if (!string.IsNullOrWhiteSpace(beforeSheet) && !string.IsNullOrWhiteSpace(afterSheet))
        {
            throw new ArgumentException("Cannot specify both beforeSheet and afterSheet. Choose one or neither.");
        }

        // Validate file paths
        if (string.IsNullOrWhiteSpace(sourceFile))
            throw new ArgumentException("sourceFile is required", nameof(sourceFile));
        if (string.IsNullOrWhiteSpace(targetFile))
            throw new ArgumentException("targetFile is required", nameof(targetFile));
        if (!File.Exists(sourceFile))
            throw new FileNotFoundException($"Source file not found: {sourceFile}");
        if (!File.Exists(targetFile))
            throw new FileNotFoundException($"Target file not found: {targetFile}");

        // Normalize paths for comparison
        string normalizedSource = Path.GetFullPath(sourceFile);
        string normalizedTarget = Path.GetFullPath(targetFile);
        if (string.Equals(normalizedSource, normalizedTarget, StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("Source and target files must be different. For same-file copy, use the 'copy' action.");
        }

        // Create a batch with both files open in the same Excel instance
        using var batch = ExcelSession.BeginBatch(sourceFile, targetFile);
        WorkbookAccessGuard.EnsureWritable(batch, normalizedTarget);

        return batch.Execute((ctx, ct) =>
        {
            Excel.Workbook? sourceWb = null;
            Excel.Workbook? targetWb = null;
            Excel.Worksheet? sourceSheetObj = null;
            Excel.Sheets? targetSheets = null;
            object? targetPositionSheet = null;

            try
            {
                // Get both workbooks from the batch
                sourceWb = batch.GetWorkbook(normalizedSource);
                targetWb = batch.GetWorkbook(normalizedTarget);

                // Find source sheet
                sourceSheetObj = ComUtilities.FindSheet(sourceWb, sourceSheet);
                if (sourceSheetObj == null)
                {
                    throw new InvalidOperationException($"Source sheet '{sourceSheet}' not found in '{Path.GetFileName(sourceFile)}'");
                }

                if (targetSheetName is not null)
                {
                    EnsureSheetNameAvailable(
                        targetWb, targetSheetName, ct);
                }

                bool before = !string.IsNullOrWhiteSpace(beforeSheet);
                string? anchorName = before ? beforeSheet : afterSheet;
                if (!string.IsNullOrWhiteSpace(anchorName))
                {
                    targetPositionSheet = ComUtilities.FindSheet(targetWb, anchorName);
                    if (targetPositionSheet == null)
                    {
                        throw new InvalidOperationException($"Target sheet '{anchorName}' not found in '{Path.GetFileName(targetFile)}'");
                    }
                }
                else
                {
                    // The last tab can be a chart sheet.
                    targetSheets = targetWb.Sheets;
                    targetPositionSheet = targetSheets[targetSheets.Count];
                }

                CopyWorksheet(sourceSheetObj, targetWb, targetPositionSheet, before, targetSheetName, ct);

                // Save the target workbook (source unchanged, only target modified)
                try
                {
                    ExcelShutdownService.SaveWorkbookWithTimeout(
                        targetWb, Path.GetFileName(targetFile), batch.Logger, ct);
                }
                catch (InvalidOperationException ex)
                {
                    throw new InvalidOperationException(
                        "Worksheet copy changed the target in memory, but its save was not confirmed. " +
                        "Inspect the reported partial state and any recovery session before retrying. " +
                        ex.Message, ex);
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref targetPositionSheet);
                ComUtilities.Release(ref targetSheets);
                ComUtilities.Release(ref sourceSheetObj);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult MoveToFile(string sourceFile, string sourceSheet, string targetFile, string? beforeSheet = null, string? afterSheet = null)
    {
        // Validate positioning parameters
        if (!string.IsNullOrWhiteSpace(beforeSheet) && !string.IsNullOrWhiteSpace(afterSheet))
        {
            throw new ArgumentException("Cannot specify both beforeSheet and afterSheet. Choose one or neither.");
        }

        // Validate file paths
        if (string.IsNullOrWhiteSpace(sourceFile))
            throw new ArgumentException("sourceFile is required", nameof(sourceFile));
        if (string.IsNullOrWhiteSpace(targetFile))
            throw new ArgumentException("targetFile is required", nameof(targetFile));
        if (!File.Exists(sourceFile))
            throw new FileNotFoundException($"Source file not found: {sourceFile}");
        if (!File.Exists(targetFile))
            throw new FileNotFoundException($"Target file not found: {targetFile}");

        // Normalize paths for comparison
        string normalizedSource = Path.GetFullPath(sourceFile);
        string normalizedTarget = Path.GetFullPath(targetFile);
        if (string.Equals(normalizedSource, normalizedTarget, StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("Source and target files must be different. For same-file move, use the 'move' action.");
        }

        // Create a batch with both files open in the same Excel instance
        using var batch = ExcelSession.BeginBatch(sourceFile, targetFile);
        WorkbookAccessGuard.EnsureWritable(batch, normalizedSource);
        WorkbookAccessGuard.EnsureWritable(batch, normalizedTarget);

        return batch.Execute((ctx, ct) =>
        {
            Excel.Workbook? sourceWb = null;
            Excel.Workbook? targetWb = null;
            Excel.Worksheet? sourceSheetObj = null;
            Excel.Sheets? targetSheets = null;
            object? targetPositionSheet = null;

            try
            {
                // Get both workbooks from the batch
                sourceWb = batch.GetWorkbook(normalizedSource);
                targetWb = batch.GetWorkbook(normalizedTarget);

                // Find source sheet
                sourceSheetObj = ComUtilities.FindSheet(sourceWb, sourceSheet);
                if (sourceSheetObj == null)
                {
                    throw new InvalidOperationException($"Source sheet '{sourceSheet}' not found in '{Path.GetFileName(sourceFile)}'");
                }

                if (!string.IsNullOrWhiteSpace(beforeSheet))
                {
                    targetPositionSheet = ComUtilities.FindSheet(targetWb, beforeSheet);
                    if (targetPositionSheet == null)
                    {
                        throw new InvalidOperationException($"Target sheet '{beforeSheet}' not found in '{Path.GetFileName(targetFile)}'");
                    }
                    sourceSheetObj.Move(Before: targetPositionSheet);
                }
                else if (!string.IsNullOrWhiteSpace(afterSheet))
                {
                    targetPositionSheet = ComUtilities.FindSheet(targetWb, afterSheet);
                    if (targetPositionSheet == null)
                    {
                        throw new InvalidOperationException($"Target sheet '{afterSheet}' not found in '{Path.GetFileName(targetFile)}'");
                    }
                    sourceSheetObj.Move(After: targetPositionSheet);
                }
                else
                {
                    // The final workbook tab can be a chart sheet.
                    targetSheets = targetWb.Sheets;
                    targetPositionSheet = targetSheets[targetSheets.Count];
                    sourceSheetObj.Move(After: targetPositionSheet);
                }

                // Persist the destination before saving the deletion in the source.
                var targetSaveCompleted = false;
                try
                {
                    ExcelShutdownService.SaveWorkbookWithTimeout(
                        targetWb, Path.GetFileName(targetFile), batch.Logger, ct);
                    targetSaveCompleted = true;
                    ExcelShutdownService.SaveWorkbookWithTimeout(
                        sourceWb, Path.GetFileName(sourceFile), batch.Logger, ct);
                }
                catch (InvalidOperationException ex)
                {
                    throw new InvalidOperationException(
                        (targetSaveCompleted
                            ? "Worksheet move: Excel confirmed the target save, but the source save was not confirmed. Both files may contain the worksheet. "
                            : "Worksheet move changed both workbooks in memory, but the target save was not confirmed. The source save was not attempted. ") +
                        "Inspect the reported partial state and any recovery session before retrying. " +
                        "No rollback was performed. " + ex.Message, ex);
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref targetPositionSheet);
                ComUtilities.Release(ref targetSheets);
                ComUtilities.Release(ref sourceSheetObj);
            }
        });
    }
}
