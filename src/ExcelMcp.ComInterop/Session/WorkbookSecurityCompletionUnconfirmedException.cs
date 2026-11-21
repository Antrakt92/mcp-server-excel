namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>Excel may have applied a workbook security change before its outcome could be confirmed.</summary>
internal sealed class WorkbookSecurityCompletionUnconfirmedException(string message, Exception? cause = null)
    : InvalidOperationException(message, cause);
