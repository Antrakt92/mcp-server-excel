namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>Excel returned from a workbook change, but a follow-up check or cleanup failed.</summary>
/// <param name="message">What happened and which workbook state needs inspection.</param>
/// <param name="cause">The original follow-up failure.</param>
public sealed class OperationCompletionUnconfirmedException(string message, Exception cause)
    : InvalidOperationException(message, cause);
