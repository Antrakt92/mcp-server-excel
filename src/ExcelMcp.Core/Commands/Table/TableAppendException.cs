namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>Excel failed after table expansion began.</summary>
public sealed class TableAppendException : InvalidOperationException
{
    internal TableAppendException(string message, TableAppendPartialState partialState, Exception cause)
        : base(message, cause) => PartialState = partialState;

    /// <summary>The last verified stage of the append, with no implication of rollback.</summary>
    public TableAppendPartialState PartialState { get; }
}

/// <summary>State observed between table expansion and the attempted value write.</summary>
/// <param name="TableName">Target table.</param>
/// <param name="OriginalRowCount">Table row count observed before expansion.</param>
/// <param name="ExpectedRowCount">Requested final table row count.</param>
/// <param name="ExpansionVerified">Excel's expanded row count and data range were verified.</param>
/// <param name="ValueWriteAttempted">The provided rows may have been written.</param>
public sealed record TableAppendPartialState(string TableName, int OriginalRowCount, int ExpectedRowCount,
    bool ExpansionVerified, bool ValueWriteAttempted);
