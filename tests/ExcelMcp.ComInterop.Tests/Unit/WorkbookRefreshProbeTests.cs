using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ErrorHandling")]
[Trait("RequiresExcel", "false")]
[System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2201",
    Justification = "Typed COM doubles reproduce the exact Excel HRESULT without opening a workbook.")]
public sealed class WorkbookRefreshProbeTests
{
    [Fact]
    public void ReadOleDbConnection_ExcelObjectIsUnavailable_LeavesRefreshStateUnknown()
    {
        var connection = CreateConnection(() =>
            throw new COMException("Cannot read OLE DB refresh state.", unchecked((int)0x800A01A8)));

        Assert.Equal(WorkbookRefreshState.Unknown, WorkbookRefreshProbe.ReadOleDbConnection(connection));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReadOleDbConnection_ObservedState_IsPreserved(bool refreshing)
    {
        var expected = refreshing ? WorkbookRefreshState.Refreshing : WorkbookRefreshState.Ready;
        Assert.Equal(expected, WorkbookRefreshProbe.ReadOleDbConnection(CreateConnection(() => refreshing)));
    }

    [Theory]
    [InlineData(unchecked((int)0x800706BE))]
    [InlineData(unchecked((int)0x80070005))]
    public void ReadOleDbConnection_UnexpectedFailure_PreservesOriginalError(int hresult)
    {
        var failure = new COMException("Unexpected OLE DB failure.", hresult);
        var connection = CreateConnection(() => throw failure);

        Assert.Same(failure, Assert.Throws<COMException>(() => WorkbookRefreshProbe.ReadOleDbConnection(connection)));
    }

    private static Excel.OLEDBConnection CreateConnection(Func<bool> read)
    {
        var connection = DispatchProxy.Create<Excel.OLEDBConnection, OleDbReadinessProxy>();
        ((OleDbReadinessProxy)(object)connection).Read = read;
        return connection;
    }

    public class OleDbReadinessProxy : DispatchProxy
    {
        internal Func<bool> Read { get; set; } = null!;

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args) =>
            targetMethod?.Name == "get_Refreshing" ? Read() : throw new NotSupportedException(targetMethod?.Name);
    }
}
