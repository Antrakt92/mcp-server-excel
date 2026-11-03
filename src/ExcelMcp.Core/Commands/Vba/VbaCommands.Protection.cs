using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

public partial class VbaCommands
{
    /// <inheritdoc />
    public VbaProjectProtectionResult ProtectProject(IExcelBatch batch, string password)
    {
        ArgumentException.ThrowIfNullOrEmpty(password);
        if (password.Any(char.IsControl))
            throw new ArgumentException("VBA project passwords cannot contain control characters.", nameof(password));
        // VBA hashes MBCS bytes in PROJECTCODEPAGE, which the public object model
        // does not expose. Different Unicode passwords can otherwise become the
        // same native password. Printable ASCII avoids that lossy conversion.
        if (password.Any(character => character > 0x7E))
            throw new ArgumentException("VBA project passwords require printable ASCII characters. Excel's legacy project password encoding cannot be verified losslessly for non-ASCII input.", nameof(password));
        EnsureVbaFile(batch.WorkbookPath);

        return batch.Execute((ctx, _) =>
        {
            if (ctx.Book.ReadOnly)
                throw new OperationFailureException(OperationFailureCategory.Permissions, "VBA protection cannot be configured in a read-only workbook.");
            if (!ctx.Book.HasVBProject)
                throw new OperationFailureException(OperationFailureCategory.Prerequisite, "The workbook must contain VBA code or a custom form before Excel can retain project password protection.");
            ExcelVbaProjectProtection.AppliedProtection applied;
            try { applied = ExcelVbaProjectProtection.Protect(ctx, password); }
            catch (ExcelVbaProtectionException error)
            {
                var state = error.State;
                var cause = error.InnerException ?? error;
                if (cause is UnauthorizedAccessException)
                    cause = new OperationFailureException(OperationFailureCategory.Permissions, cause.Message, cause);
                throw new VbaProjectProtectionException(error.Message, cause,
                    new VbaProjectProtectionState
                    {
                        Phase = state.Phase,
                        AcceptanceRequested = state.AcceptanceRequested,
                        ConfigurationApplied = state.ConfigurationApplied,
                        LockOptionReadback = state.LockOptionReadback,
                        DialogCleanupConfirmed = state.DialogCleanupConfirmed,
                        ActiveProjectRestored = state.ActiveProjectRestored,
                        VisibilityRestored = state.VisibilityRestored,
                        SaveRequired = state.SaveRequired,
                        ReopenRequired = state.ReopenRequired
                    });
            }
            return new VbaProjectProtectionResult
            {
                Success = true,
                FilePath = batch.WorkbookPath,
                ProjectName = applied.ProjectName,
                CurrentProtection = applied.CurrentProtection,
                ConfigurationApplied = true,
                SaveRequired = true,
                ReopenRequired = true
            };
        });
    }
}
