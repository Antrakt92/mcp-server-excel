namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>Bounded, content-free facts attached only at a known execution boundary.</summary>
internal sealed record ExcelExecutionFailureContext
{
    private const int MaxNodes = 32;
    private const int MaxEdges = 64;

    private ExcelExecutionFailureContext(Exception? operationFailure, Exception snapshotFailure, bool callbackReturned)
    {
        CallbackReturned = callbackReturned;
        OperationExceptionType = operationFailure is null ? null : GetTypeName(operationFailure);
        OperationHResult = operationFailure?.HResult;
        SnapshotExceptionType = GetTypeName(snapshotFailure);
        SnapshotHResult = snapshotFailure.HResult;
    }

    internal bool CallbackReturned { get; }
    internal string? OperationExceptionType { get; }
    internal int? OperationHResult { get; }
    internal string SnapshotExceptionType { get; }
    internal int SnapshotHResult { get; }

    /// <summary>Annotates the existing wire exception without replacing its type, causes, or cancellation token.</summary>
    internal static void Attach(Exception wireException, Exception? operationFailure,
        Exception snapshotFailure, bool callbackReturned)
    {
        ArgumentNullException.ThrowIfNull(wireException);
        ArgumentNullException.ThrowIfNull(snapshotFailure);
        if (callbackReturned && operationFailure is not null)
            throw new ArgumentException("A returned callback cannot also have an operation failure.", nameof(callbackReturned));
        if (wireException.Data[typeof(ExcelExecutionFailureContext)] is not ExcelExecutionFailureContext)
            wireException.Data[typeof(ExcelExecutionFailureContext)] =
                new ExcelExecutionFailureContext(operationFailure, snapshotFailure, callbackReturned);
    }

    /// <summary>Finds explicit boundary facts; incomplete traversal never establishes the absence of a boundary.</summary>
    internal static bool TryFind(Exception error, out ExcelExecutionFailureContext? context, out bool complete)
    {
        ArgumentNullException.ThrowIfNull(error);
        context = null;
        var traversalComplete = true;
        var seen = new HashSet<Exception>(ReferenceEqualityComparer.Instance) { error };
        var pending = new Queue<Exception>();
        pending.Enqueue(error);
        var edges = 0;
        while (pending.TryDequeue(out var current))
        {
            if (current.Data[typeof(ExcelExecutionFailureContext)] is ExcelExecutionFailureContext candidate)
            {
                if (context is not null && !ReferenceEquals(context, candidate))
                {
                    context = null;
                    complete = false;
                    return false;
                }
                context ??= candidate;
            }
            if (current is AggregateException aggregate)
            {
                foreach (var child in aggregate.InnerExceptions)
                {
                    if (edges == MaxEdges)
                    {
                        traversalComplete = false;
                        pending.Clear();
                        break;
                    }
                    edges++;
                    Enqueue(child);
                }
            }
            else if (current.InnerException is { } inner)
            {
                if (edges == MaxEdges)
                {
                    traversalComplete = false;
                    break;
                }
                edges++;
                Enqueue(inner);
            }
        }
        complete = traversalComplete;
        return context is not null;

        void Enqueue(Exception child)
        {
            if (seen.Contains(child))
                return;
            if (seen.Count >= MaxNodes)
            {
                traversalComplete = false;
                return;
            }
            seen.Add(child);
            pending.Enqueue(child);
        }
    }

    private static string GetTypeName(Exception error)
    {
        var name = error.GetType().Name;
        return name.Length <= 128 ? name : name[..128];
    }
}
