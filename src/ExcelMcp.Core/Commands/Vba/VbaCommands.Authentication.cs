using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

public partial class VbaCommands
{
    /// <inheritdoc />
    public VbaProjectAuthenticationResult UnlockProject(IExcelBatch batch, string password) =>
        AuthenticateProject(batch, password, verifyOnly: false);

    /// <inheritdoc />
    public VbaProjectAuthenticationResult VerifyPassword(IExcelBatch batch, string password) =>
        AuthenticateProject(batch, password, verifyOnly: true);

    private static VbaProjectAuthenticationResult AuthenticateProject(IExcelBatch batch, string password, bool verifyOnly)
    {
        ArgumentException.ThrowIfNullOrEmpty(password);
        if (password.Any(character => character < 0x20 || character > 0x7E))
            throw new ArgumentException("VBA project authentication requires printable ASCII characters; non-ASCII and control characters cannot be verified losslessly.", nameof(password));
        if (batch is AttachedWorkbookBatch)
            throw new OperationFailureException(OperationFailureCategory.Permissions,
                "VBA password authentication is supported only in an owned workbook session. Attached workbooks were not changed.");
        EnsureVbaFile(batch.WorkbookPath);
        return batch.Execute((context, _) =>
        {
            if (context.Book.ReadOnly)
                throw new OperationFailureException(OperationFailureCategory.Permissions,
                    "VBA password authentication is disabled for a read-only workbook.");
            if (!context.Book.HasVBProject)
                throw new OperationFailureException(OperationFailureCategory.Prerequisite, "The workbook has no VBA project to authenticate.");
            ExcelVbaProjectProtection.AuthenticatedProject native;
            try { native = ExcelVbaProjectProtection.Authenticate(context, password); }
            catch (ExcelVbaAuthenticationException error)
            {
                throw new VbaProjectAuthenticationException(error.Message, error.InnerException ?? error,
                    MapAuthenticationState(error.State));
            }
            var state = MapAuthenticationState(native.State);
            if (verifyOnly && native.AlreadyAccessible)
                throw new VbaProjectAuthenticationException(
                    "The project is already accessible, so the supplied password was not verified. A locked project is required; nothing was reopened.",
                    new OperationFailureException(OperationFailureCategory.Prerequisite, "An actual locked-project password challenge is required."), state);
            if (!verifyOnly && native.State.PasswordVerified is false)
                throw new VbaProjectAuthenticationException("Excel rejected the supplied project password. The project remains locked.",
                    new OperationFailureException(OperationFailureCategory.Permissions, "The known password was rejected by Excel."), state);
            return new VbaProjectAuthenticationResult
            {
                Success = true,
                FilePath = batch.WorkbookPath,
                Action = verifyOnly ? "verify-password" : "unlock-project",
                ProjectName = native.ProjectName,
                ProtectionBefore = native.ProtectionBefore,
                ProtectionAfter = native.ProtectionAfter,
                AlreadyAccessible = native.AlreadyAccessible,
                AuthenticationPerformed = native.State.AuthenticationPerformed,
                PasswordVerified = native.State.PasswordVerified,
                AccessUnlocked = native.State.AccessUnlocked is true
            };
        });
    }

    private static VbaProjectAuthenticationState MapAuthenticationState(ExcelVbaAuthenticationState state) => new()
    {
        Phase = state.Phase,
        ChallengeObserved = state.ChallengeObserved,
        AuthenticationPerformed = state.AuthenticationPerformed,
        PasswordVerified = state.PasswordVerified,
        AccessUnlocked = state.AccessUnlocked,
        AlreadyAccessible = state.AlreadyAccessible,
        CurrentProtection = state.CurrentProtection,
        DialogCleanupConfirmed = state.DialogCleanupConfirmed,
        ActiveProjectRestored = state.ActiveProjectRestored,
        VisibilityRestored = state.VisibilityRestored
    };
}
