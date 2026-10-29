using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

public partial class VbaCommands
{
    /// <inheritdoc />
    public VbaInvokeResult Invoke(IExcelBatch batch, string procedureName, TimeSpan? timeout, List<VbaArgument>? arguments = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(procedureName);
        if (procedureName.Contains('!'))
            throw new ArgumentException("Use an unqualified Module.Procedure name; the selected session determines the workbook.", nameof(procedureName));
        object[] values = VbaInvocationValues.Arguments(arguments);
        EnsureVbaFile(batch.WorkbookPath);
        using var timeoutCts = timeout.HasValue ? new CancellationTokenSource(timeout.Value) : null;
        try
        {
            return batch.Execute((ctx, ct) =>
            {
                EnsureMacroApproval(batch, ctx.Book);
                var culture = CultureInfo.CurrentCulture;
                var uiCulture = CultureInfo.CurrentUICulture;
                object? returned = null;
                try
                {
                    CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
                    CultureInfo.CurrentUICulture = CultureInfo.CurrentCulture;
                    string qualified = $"'{ctx.Book.Name.Replace("'", "''", StringComparison.Ordinal)}'!{procedureName}";
                    returned = ctx.App.Run(qualified,
                        values[0], values[1], values[2], values[3], values[4], values[5], values[6], values[7], values[8], values[9],
                        values[10], values[11], values[12], values[13], values[14], values[15], values[16], values[17], values[18], values[19],
                        values[20], values[21], values[22], values[23], values[24], values[25], values[26], values[27], values[28], values[29]);
                    var result = VbaInvocationValues.Result(returned);
                    result.FilePath = batch.WorkbookPath;
                    result.Action = "invoke";
                    return result;
                }
                finally
                {
                    CultureInfo.CurrentCulture = culture;
                    CultureInfo.CurrentUICulture = uiCulture;
                    VbaInvocationValues.ReleaseReturnedObjects(returned);
                }
            }, timeoutCts?.Token ?? default);
        }
        catch (OperationCanceledException ex) when (timeoutCts?.IsCancellationRequested == true)
        {
            throw new TimeoutException("VBA procedure execution exceeded its requested timeout; workbook effects were not rolled back.", ex);
        }
    }

    private static void EnsureMacroApproval(IExcelBatch batch, Excel.Workbook workbook)
    {
        var owner = batch is AttachedWorkbookBatch attached ? attached.Owner : batch as ExcelBatch;
        if (owner?.RequiresVbaApproval(workbook) == true)
            throw new InvalidOperationException(batch is AttachedWorkbookBatch
                ? "VBA macros were disabled when this workbook opened. Review macro permissions in Excel; an attached workbook cannot be reopened by this session. No procedure was run."
                : "VBA macros were disabled when this workbook opened. Obtain approval for file enable-macros with approved:true and allow_reopen:true (CLI: session enable-macros --approved --allow-reopen) before running a procedure. Unsaved changes are not discarded.");
    }
}
