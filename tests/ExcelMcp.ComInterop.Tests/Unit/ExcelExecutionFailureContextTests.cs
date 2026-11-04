using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ExcelExecutionFailureContextTests
{
    [Fact]
    public void Attach_PreservesWireFailuresAndCopiesOnlyBoundedFacts()
    {
        const string secret = "synthetic-password-and-workbook-content";
        var operation = new InvalidOperationException(secret);
#pragma warning disable CA2201 // Synthetic native failure; this test never accesses Excel.
        var snapshot = new COMException(secret, unchecked((int)0x800706BE));
#pragma warning restore CA2201
        var wire = new AggregateException(secret, operation, snapshot);
        wire.Data["secret"] = secret;
        ExcelExecutionFailureContext.Attach(wire, operation, snapshot, callbackReturned: false);

        Assert.True(ExcelExecutionFailureContext.TryFind(wire, out var context, out var complete));
        Assert.True(complete);
        Assert.NotNull(context);
        Assert.False(context.CallbackReturned);
        Assert.Equal(nameof(InvalidOperationException), context.OperationExceptionType);
        Assert.Equal(operation.HResult, context.OperationHResult);
        Assert.Equal(nameof(COMException), context.SnapshotExceptionType);
        Assert.Equal(unchecked((int)0x800706BE), context.SnapshotHResult);
        Assert.DoesNotContain(secret, context.ToString(), StringComparison.Ordinal);
        Assert.Same(operation, wire.InnerExceptions[0]);
        Assert.Same(snapshot, wire.InnerExceptions[1]);
        Assert.Equal(secret, wire.Data["secret"]);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SnapshotOnly_RecordsWhetherCallbackReturned(bool callbackReturned)
    {
        var snapshot = new InvalidOperationException("snapshot");
        ExcelExecutionFailureContext.Attach(snapshot, null, snapshot, callbackReturned);
        Assert.True(ExcelExecutionFailureContext.TryFind(snapshot, out var context, out var complete));
        Assert.True(complete);
        Assert.NotNull(context);
        Assert.Equal(callbackReturned, context.CallbackReturned);
        Assert.Null(context.OperationExceptionType);
        Assert.Null(context.OperationHResult);
        Assert.Equal(nameof(InvalidOperationException), context.SnapshotExceptionType);
    }

    [Fact]
    public void NestedWrappersAndRepeatedReferences_FindSameExplicitBoundary()
    {
        var marked = new InvalidOperationException("snapshot");
        ExcelExecutionFailureContext.Attach(marked, null, marked, callbackReturned: true);
        var root = new TargetInvocationException(new AggregateException(marked, marked,
            new InvalidOperationException("outer", marked)));
        Assert.True(ExcelExecutionFailureContext.TryFind(root, out var context, out var complete));
        Assert.True(complete);
        Assert.True(ExcelExecutionFailureContext.TryFind(marked, out var original, out _));
        Assert.Same(original, context);
    }

    [Fact]
    public void UnmarkedAggregate_DoesNotInventBoundaryPhases()
    {
        var error = new AggregateException(new InvalidOperationException("operation"),
            new InvalidOperationException("snapshot"));
        Assert.False(ExcelExecutionFailureContext.TryFind(error, out var context, out var complete));
        Assert.Null(context);
        Assert.True(complete);
    }

    [Fact]
    public void Cancellation_RemainsSameExceptionWithOriginalTokenAndCauses()
    {
        using var source = new CancellationTokenSource();
        source.Cancel();
        var snapshot = new InvalidOperationException("snapshot");
        var operation = new OperationCanceledException(source.Token);
        var causes = new AggregateException(operation, snapshot);
        var wire = new OperationCanceledException("cancelled", causes, source.Token);
        ExcelExecutionFailureContext.Attach(wire, operation, snapshot, callbackReturned: false);
        Assert.Equal(source.Token, wire.CancellationToken);
        Assert.Same(causes, wire.InnerException);
        Assert.True(ExcelExecutionFailureContext.TryFind(wire, out var context, out var complete));
        Assert.True(complete);
        Assert.Equal(nameof(OperationCanceledException), context!.OperationExceptionType);
    }

    [Theory]
    [InlineData(31, true)]
    [InlineData(40, false)]
    public void DeepBoundary_ReportsWhetherSearchReachedItsMarker(int wrapperCount, bool expectedFound)
    {
        Exception error = new InvalidOperationException("snapshot");
        ExcelExecutionFailureContext.Attach(error, null, error, callbackReturned: true);
        for (var index = 0; index < wrapperCount; index++)
            error = new InvalidOperationException("wrapper", error);
        Assert.Equal(expectedFound, ExcelExecutionFailureContext.TryFind(error, out var context, out var complete));
        Assert.Equal(expectedFound, context is not null);
        Assert.Equal(expectedFound, complete);
    }

    [Theory]
    [InlineData(64, true)]
    [InlineData(65, false)]
    public void EdgeBound_StopsGloballyBeforeProcessingQueuedChildren(int edgeCount, bool expectedFound)
    {
        var child = new InvalidOperationException("child");
        ExcelExecutionFailureContext.Attach(child, null, child, callbackReturned: true);
        var root = new AggregateException(Enumerable.Repeat(child, edgeCount));
        Assert.Equal(expectedFound, ExcelExecutionFailureContext.TryFind(root, out var context, out var complete));
        Assert.Equal(expectedFound, context is not null);
        Assert.Equal(expectedFound, complete);
    }

    [Fact]
    public void MultipleDistinctBoundaries_DoNotSelectAnAuthoritativeContext()
    {
        var first = new InvalidOperationException("first");
        var second = new InvalidOperationException("second");
        ExcelExecutionFailureContext.Attach(first, null, first, callbackReturned: true);
        ExcelExecutionFailureContext.Attach(second, null, second, callbackReturned: false);
        Assert.False(ExcelExecutionFailureContext.TryFind(new AggregateException(first, second), out var context, out var complete));
        Assert.Null(context);
        Assert.False(complete);
    }

    [Fact]
    public void RepeatedAttachment_RetainsOriginalBoundaryFacts()
    {
        var wire = new InvalidOperationException("first");
        ExcelExecutionFailureContext.Attach(wire, null, wire, callbackReturned: true);
        Assert.True(ExcelExecutionFailureContext.TryFind(wire, out var first, out _));
        ExcelExecutionFailureContext.Attach(wire, new ArgumentException("later"), wire, callbackReturned: false);
        Assert.True(ExcelExecutionFailureContext.TryFind(wire, out var second, out var complete));
        Assert.True(complete);
        Assert.Same(first, second);
        Assert.True(second!.CallbackReturned);
    }
}
