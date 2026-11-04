namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>Fail-closed recognition policy; window/process ownership is checked separately.</summary>
internal static class VbaAuthenticationChallengePolicy
{
    internal static bool CanRecognize(string expectedProject, string caption, bool commandCommitted,
        bool preexisting, bool ownedByEditor, int editCount, bool passwordMasked, int editId,
        int buttonCount, bool standardOk, bool standardCancel, bool hasPropertiesTab) =>
        commandCommitted && !preexisting && ownedByEditor &&
        caption.Equals(expectedProject + " Password", StringComparison.Ordinal) &&
        editCount == 1 && passwordMasked && editId == 5470 && buttonCount == 2 &&
        standardOk && standardCancel && !hasPropertiesTab;

    internal static bool IsInvalidPasswordOwner(nint challenge, nint directOwner) =>
        challenge != 0 && directOwner == challenge;

    internal static bool IsPropertiesReplacement(string expectedProject, string caption,
        bool authenticationPerformed, bool challengeStillExists, nint capturedEditorOwner,
        nint directOwner, int ownedModalCount) =>
        authenticationPerformed && !challengeStillExists && capturedEditorOwner != 0 &&
        directOwner == capturedEditorOwner && ownedModalCount == 1 &&
        caption.Equals(expectedProject + " - Project Properties", StringComparison.Ordinal);

    internal static bool CanCleanUp(bool recordedDuringCommand, bool stillMatchesRecordedSignature) =>
        recordedDuringCommand && stillMatchesRecordedSignature;
}
