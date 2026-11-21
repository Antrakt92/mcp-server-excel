namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal sealed class WorkbookSaveIncompleteException(
    string fileName, bool? excelReportedSaveSucceeded, string? observationError,
    bool? hasUnsavedChanges = true, Exception? innerException = null)
    : InvalidOperationException(CreateMessage(fileName, excelReportedSaveSucceeded, hasUnsavedChanges), innerException)
{
    internal bool? ExcelReportedSaveSucceeded { get; } = excelReportedSaveSucceeded;
    internal string? ObservationError { get; } = observationError;
    internal bool? HasUnsavedChanges { get; } = hasUnsavedChanges;

    private static string CreateMessage(string fileName, bool? excelReportedSaveSucceeded, bool? hasUnsavedChanges) =>
        (hasUnsavedChanges == true ? $"Workbook '{fileName}' still has unsaved changes. " :
            $"The save call for workbook '{fileName}' failed and its unsaved state could not be confirmed. ") +
        (excelReportedSaveSucceeded switch
        {
            true => "Excel reported a successful write during the save call. ",
            false => "Excel reported an unsuccessful save. ",
            _ => "Whether Excel wrote the file is unconfirmed. "
        }) + "Inspect the workbook and saved file before retrying.";
}
