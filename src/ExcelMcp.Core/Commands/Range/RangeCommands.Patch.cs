using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    /// <inheritdoc />
    public RangePatchResult PatchCells(IExcelBatch batch, string sheetName, RangePatchOptions patchOptions)
    {
        var edits = PrepareCellEdits(patchOptions);
        return batch.Execute((context, cancellationToken) =>
        {
            var result = new RangePatchResult { FilePath = batch.WorkbookPath, SheetName = sheetName };
            foreach (var edit in edits)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!WithPatchCell(context, sheetName, edit.Address, cell =>
                {
                    ValidatePatchTarget(context, cell, edit.Address);
                    return PatchCellMatches(context, cell, edit.Expected);
                }))
                    result.ConflictAddresses.Add(edit.Address);
            }

            if (result.ConflictAddresses.Count > 0)
                return PatchConflict(result);

            foreach (var edit in edits)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!WithPatchCell(context, sheetName, edit.Address, cell =>
                {
                    ValidatePatchTarget(context, cell, edit.Address);
                    if (!PatchCellMatches(context, cell, edit.Expected))
                        return false;

                    if (!Equals(edit.Expected, edit.Replacement))
                    {
                        result.UnconfirmedAddress = edit.Address;
                        WritePatchCell(cell, edit.Replacement);
                        result.WriteReturnedAddresses.Add(edit.Address);
                        if (!PatchCellMatches(context, cell, edit.Replacement))
                        {
                            result.ReadbackMismatch = true;
                            throw new InvalidOperationException($"Excel readback differs from the requested content at {edit.Address}.");
                        }
                        result.UnconfirmedAddress = null;
                    }

                    result.ConfirmedAddresses.Add(edit.Address);
                    return true;
                }, result))
                {
                    result.ConflictAddresses.Add(edit.Address);
                    return PatchConflict(result);
                }
            }

            // A later edit or a workbook event can change a previously read-back cell.
            foreach (var edit in edits)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!WithPatchCell(context, sheetName, edit.Address,
                    cell => PatchCellMatches(context, cell, edit.Replacement), result))
                {
                    result.ConfirmedAddresses.Remove(edit.Address);
                    result.ConflictAddresses.Add(edit.Address);
                }
            }

            if (result.ConflictAddresses.Count > 0)
                return PatchConflict(result);

            result.Success = true;
            return result;
        });
    }

    private static RangePatchResult PatchConflict(RangePatchResult result)
    {
        result.ErrorCategory = "Conflict";
        result.ErrorMessage = $"Cell contents differ at {string.Join(", ", result.ConflictAddresses)}. Inspect confirmed addresses before continuing.";
        return result;
    }

    private static List<PreparedCellEdit> PrepareCellEdits(RangePatchOptions options)
    {
        if (options?.Edits is null || options.Edits.Count is < 1 or > 100)
            throw new ArgumentException("Supply 1–100 single-cell edits.", nameof(options));

        var addresses = new HashSet<(int Row, int Column)>();
        var edits = new List<PreparedCellEdit>(options.Edits.Count);
        int textLength = 0;
        foreach (var edit in options.Edits)
        {
            if (edit?.Address is null || !RangeHelpers.TryParseCellReference(edit.Address, out int column, out int row))
                throw new ArgumentException("Every edit needs one sheet-local A1 cell address.", nameof(options));

            string address = $"{GetColumnLetter(column)}{row}";
            if (!addresses.Add((row, column)))
                throw new ArgumentException($"Duplicate target {address}.", nameof(options));

            var expected = PrepareCellState(edit.Expected);
            var replacement = PrepareCellState(edit.Replacement);
            if (!replacement.IsFormula && replacement.Content is string replacementText)
                RangeHelpers.ValidateLiteralCellText(replacementText);
            textLength += (expected.Content as string)?.Length ?? 0;
            textLength += (replacement.Content as string)?.Length ?? 0;
            if (textLength > 131_072)
                throw new ArgumentException("The combined patch text exceeds 131,072 characters.", nameof(options));
            edits.Add(new PreparedCellEdit(address, expected, replacement));
        }
        return edits;
    }

    private static PreparedCellState PrepareCellState(RangeCellState? state)
    {
        if (state is null)
            throw new ArgumentException("Every edit needs expected and replacement states.");
        if (state.Formula is not null)
        {
            if (state.Value.ValueKind != JsonValueKind.Undefined || !state.Formula.StartsWith('=')
                || state.Formula.Length is < 2 or > 8192 || state.Formula.Contains('\0'))
                throw new ArgumentException("Specify one valid A1 formula or a value for each state.");
            return new(true, state.Formula);
        }

        return state.Value.ValueKind switch
        {
            JsonValueKind.Null => new(false, null),
            JsonValueKind.True => new(false, true),
            JsonValueKind.False => new(false, false),
            JsonValueKind.Number when state.Value.TryGetDouble(out double number) && double.IsFinite(number) => new(false, number),
            JsonValueKind.String when state.Value.GetString() is { Length: <= 32767 } text && !text.Contains('\0') => new(false, text),
            _ => throw new ArgumentException("A value must be null, a boolean, a finite number, or Excel text.")
        };
    }

    private static bool WithPatchCell(ExcelContext context, string sheetName, string address,
        Func<Excel.Range, bool> action, RangePatchResult? progress = null)
    {
        Excel.Range? cell = null;
        try
        {
            cell = RangeHelpers.ResolveRange(context.Book, sheetName, address, out var error)
                ?? throw new InvalidOperationException(error ?? RangeHelpers.GetResolveError(sheetName, address));
            return action(cell);
        }
        catch (Exception error) when (progress is not null && error is not OperationCanceledException
            && (progress.UnconfirmedAddress is not null || progress.ConfirmedAddresses.Count > 0))
        {
            throw new RangePatchException(progress, error);
        }
        finally
        {
            ComUtilities.Release(ref cell);
        }
    }

    private static void ValidatePatchTarget(ExcelContext context, Excel.Range cell, string address)
    {
        Excel.Worksheet? sheet = null;
        Excel.ListObject? table = null;
        try
        {
            sheet = cell.Worksheet;
            table = cell.ListObject;
            if (table is not null || cell.MergeCells is true || cell.HasArray
                || (context.Capabilities.SupportsFormula2 && cell.HasSpill is not false)
                || (sheet.ProtectContents && cell.Locked is not false))
                throw new ArgumentException($"Cell {address} is in a table, merge, array, spill, or protected locked area.");
        }
        finally
        {
            ComUtilities.Release(ref table);
            ComUtilities.Release(ref sheet);
        }
    }

    private static bool PatchCellMatches(ExcelContext context, Excel.Range cell, PreparedCellState expected)
    {
        if (cell.HasFormula is true)
            return expected.IsFormula && Equals(ReadFormulas(context, cell), expected.Content);
        if (expected.IsFormula)
            return false;
        object? actual = cell.Value2;
        return Equals(actual, expected.Content)
            && (!ExcelErrorMapper.TryGet(actual, out _, out _) || !IsNativeError(context, cell, 1, 1));
    }

    private static void WritePatchCell(Excel.Range cell, PreparedCellState replacement)
    {
        if (replacement.IsFormula)
            cell.Formula = replacement.Content!;
        else if (replacement.Content is null)
            cell.ClearContents();
        else
            cell.Value2 = replacement.Content is string text ? "'" + text : replacement.Content;
    }

    private sealed record PreparedCellEdit(string Address, PreparedCellState Expected, PreparedCellState Replacement);
    private sealed record PreparedCellState(bool IsFormula, object? Content);
}
