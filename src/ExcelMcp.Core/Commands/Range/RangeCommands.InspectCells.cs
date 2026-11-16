using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    /// <inheritdoc />
    public RangeCellInspectionResult InspectCells(IExcelBatch batch, string sheetName, string rangeAddress)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(rangeAddress);
        return batch.Execute((context, cancellationToken) =>
        {
            Excel.Range? range = null;
            Excel.Worksheet? sheet = null;
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                range = RangeHelpers.ResolveRange(context.Book, sheetName, rangeAddress, out string? error);
                if (range is null)
                    throw new ArgumentException(error ?? RangeHelpers.GetResolveError(sheetName, rangeAddress));

                var (rows, columns) = RequireReadRectangle(range);
                RangeReadWindow.EnsureCellLimit(rows, columns, 64,
                    "Select a smaller range or use get-values-page for data without displayed text.");
                sheet = range.Worksheet;
                var result = new RangeCellInspectionResult
                {
                    FilePath = batch.WorkbookPath,
                    SheetName = sheet.Name,
                    RangeAddress = range.Address,
                    Action = "inspect-cells"
                };
                RangeHelpers.VisitCells(range, cancellationToken,
                    cell => result.Cells.Add(ReadInspectedCell(context, cell)));
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref range);
            }
        });
    }

    private static InspectedRangeCell ReadInspectedCell(ExcelContext context, Excel.Range cell)
    {
        const int maxTextLength = 512;
        string address = cell.Address;
        object? value = cell.Value2;
        int? errorCode = null;
        if (ExcelErrorMapper.TryGet(value, out var code, out var error) && IsNativeError(context, cell, 1, 1))
        {
            value = error.Name;
            errorCode = code;
        }

        if ((object?)cell.HasFormula is not bool hasFormula)
            throw new InvalidOperationException($"Excel returned an indeterminate formula state for {address}.");
        string? formula = hasFormula
            ? ReadFormulas(context, cell) as string ??
                throw new InvalidOperationException($"Excel did not return formula text for {address}.")
            : null;
        if ((object?)cell.Text is not string displayedText)
            throw new InvalidOperationException($"Excel did not return displayed text for {address}.");

        object? previewValue = value is string text ? LimitCellPreviewText(text, maxTextLength) : value;
        string? previewFormula = LimitCellPreviewText(formula, maxTextLength);
        string previewDisplayedText = LimitCellPreviewText(displayedText, maxTextLength)!;
        return new InspectedRangeCell(address, previewValue, previewFormula, previewDisplayedText, errorCode,
            !Equals(previewValue, value), previewFormula != formula, previewDisplayedText != displayedText);
    }
}
