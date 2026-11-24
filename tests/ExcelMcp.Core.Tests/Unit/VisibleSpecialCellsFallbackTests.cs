using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "false")]
public sealed class VisibleSpecialCellsFallbackTests
{
    [Theory]
    [InlineData(true, false, 2, 3, 2, 3)]
    [InlineData(false, true, 2, 3, 2, 3)]
    [InlineData(null, null, 2, 3, 2, 3)]
    [InlineData(false, true, null, 3, 2, 0)]
    [InlineData(true, false, 1, null, 1, 3)]
    public void GetSpecialCells_VerifiesActualMembersDespiteMisleadingAggregate(
        bool? rowsHidden, bool? columnsHidden, int? visibleRow, int? visibleColumn, int rowReads, int columnReads)
    {
        var fixture = new VisibilityProbe(rowsHidden, columnsHidden, visibleRow, visibleColumn);
        if (visibleRow.HasValue && visibleColumn.HasValue)
        {
            var error = Assert.Throws<COMException>(fixture.Run);
            Assert.Same(fixture.NativeFailure, error);
        }
        else
        {
            var result = fixture.Run();
            Assert.True(result.Success);
            Assert.Empty(result.Areas);
            Assert.Equal(0, result.CellCount);
        }

        Assert.Equal(0, fixture.AggregateReads);
        Assert.Equal(Enumerable.Range(1, rowReads), fixture.RowMemberIndexes);
        Assert.Equal(Enumerable.Range(1, columnReads), fixture.ColumnMemberIndexes);
    }

    [Fact]
    public void GetSpecialCells_ChecksSmallerAxisFirstAndStopsAtProvenAbsence()
    {
        var fixture = new VisibilityProbe(true, false, visibleRow: 3, visibleColumn: null, rowCount: 5);
        var result = fixture.Run();
        Assert.True(result.Success);
        Assert.Equal(0, result.CellCount);
        Assert.Empty(result.Areas);
        Assert.Empty(fixture.RowMemberIndexes);
        Assert.Equal([1, 2, 3], fixture.ColumnMemberIndexes);
    }

    [Fact]
    public void GetSpecialCells_PreservesNativeErrorWhenBoundedProbeCannotVerifyAbsence()
    {
        var fixture = new VisibilityProbe(false, true, visibleRow: 1, visibleColumn: null,
            columnCount: RangeCommands.MaxVisibleFallbackProbes);
        var failure = Assert.Throws<AggregateException>(fixture.Run);
        Assert.Collection(failure.InnerExceptions,
            native => Assert.Same(fixture.NativeFailure, native),
            verification => Assert.Contains("visibility check limit", verification.Message, StringComparison.Ordinal));
        Assert.Equal([1], fixture.RowMemberIndexes);
        Assert.Equal(Enumerable.Range(1, RangeCommands.MaxVisibleFallbackProbes - 1), fixture.ColumnMemberIndexes);
    }

    [Theory]
    [InlineData("EntireRow")]
    [InlineData("EntireColumn")]
    [InlineData("RowMemberHidden")]
    [InlineData("ColumnMemberHidden")]
    public void GetSpecialCells_ReportsBothExcelErrorsWhenVisibilityCannotBeVerified(string failedMember)
    {
        var fixture = new VisibilityProbe(null, null, failedMember: failedMember);
        var error = Assert.Throws<AggregateException>(fixture.Run);

        Assert.Collection(error.InnerExceptions,
            native => Assert.Same(fixture.NativeFailure, native),
            verification => Assert.Same(fixture.ProbeFailure, verification));
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2201",
        Justification = "Exercises the ambiguous native SpecialCells error without launching Excel.")]
    private sealed class VisibilityProbe(bool? rowsHidden, bool? columnsHidden,
        int? visibleRow = 2, int? visibleColumn = 3, int rowCount = 2, int columnCount = 3,
        string? failedMember = null)
    {
        public COMException NativeFailure { get; } = new("Synthetic SpecialCells no-match code.", unchecked((int)0x800A03EC));
        public COMException ProbeFailure { get; } = new("Synthetic visibility read error.", unchecked((int)0x800A03EC));
        public List<int> RowMemberIndexes { get; } = [];
        public List<int> ColumnMemberIndexes { get; } = [];
        public int AggregateReads { get; private set; }

        public SpecialCellsResult Run()
        {
            const string address = "A1:C2";
            Excel.Range? scope = null;
            var sheet = ExcelInterfaceProxy.Create<Excel.Worksheet>((method, _) => method.Name switch
            {
                "get_Name" => "Sheet1",
                "get_Range" => scope,
                _ => throw new NotSupportedException(method.Name)
            });
            var areas = ExcelInterfaceProxy.Create<Excel.Areas>((method, _) => method.Name switch
            {
                "get_Count" => 1,
                "get_Item" or "get__Default" => scope,
                _ => throw new NotSupportedException(method.Name)
            });
            var rows = Dimension(rows: true, rowsHidden, rowCount, visibleRow, RowMemberIndexes);
            var columns = Dimension(rows: false, columnsHidden, columnCount, visibleColumn, ColumnMemberIndexes);
            scope = ExcelInterfaceProxy.Create<Excel.Range>((method, _) => method.Name switch
            {
                "get_Parent" => sheet,
                "get_Address" => address,
                "get_Areas" => areas,
                "get_CountLarge" => (double)rowCount * columnCount,
                "get_EntireRow" => Probe("EntireRow", rows),
                "get_EntireColumn" => Probe("EntireColumn", columns),
                "SpecialCells" => throw NativeFailure,
                _ => throw new NotSupportedException(method.Name)
            });
            var sheets = Properties<Excel.Sheets>(("get_Count", 1), ("get__Default", sheet));
            var workbook = Properties<Excel.Workbook>(("get_Worksheets", sheets));
            var app = ExcelInterfaceProxy.Create<Excel.Application>((method, _) => method.Name == "get_International"
                ? "General" : throw new NotSupportedException(method.Name), typeof(ExcelContext).Assembly);
            var context = new ExcelContext("fixture.xlsx", app, workbook);
            using var batch = ExcelInterfaceProxy.Create<IExcelBatch>((method, args) => method.Name switch
            {
                nameof(IDisposable.Dispose) => null,
                "get_WorkbookPath" => context.WorkbookPath,
                nameof(IExcelBatch.Execute) => Assert.IsType<Func<ExcelContext, CancellationToken, SpecialCellsResult>>(
                    args![0])(context, Assert.IsType<CancellationToken>(args[1])),
                _ => throw new NotSupportedException(method.Name)
            });

            return new RangeCommands().GetSpecialCells(batch, "Sheet1", address, SpecialCellKind.Visible);
        }

        private Excel.Range Dimension(bool rows, bool? hidden, int count, int? visible, List<int> memberIndexes)
        {
            var visibleMember = Member(false);
            var hiddenMember = Member(true);
            var members = ExcelInterfaceProxy.Create<Excel.Range>((method, args) =>
            {
                if (method.Name == "get_Count")
                    return count;
                Assert.Equal("get__Default", method.Name);
                int index = Assert.IsType<int>(args![0]);
                memberIndexes.Add(index);
                return index == visible ? visibleMember : hiddenMember;
            });
            return ExcelInterfaceProxy.Create<Excel.Range>((method, _) => method.Name switch
            {
                "get_Hidden" => ReadAggregate(hidden),
                "get_Rows" when rows => members,
                "get_Columns" when !rows => members,
                _ => throw new NotSupportedException(method.Name)
            });

            Excel.Range Member(bool memberHidden) => ExcelInterfaceProxy.Create<Excel.Range>((method, _) =>
            {
                Assert.Equal("get_Hidden", method.Name);
                return Probe(rows ? "RowMemberHidden" : "ColumnMemberHidden", memberHidden);
            });
        }

        private bool? ReadAggregate(bool? hidden)
        {
            AggregateReads++;
            return hidden;
        }

        private object? Probe(string member, object? value) => member == failedMember ? throw ProbeFailure : value;
    }

    private static T Properties<T>(params (string Name, object? Value)[] properties) where T : class =>
        ExcelInterfaceProxy.Create<T>((method, _) => properties.Single(property => property.Name == method.Name).Value);
}
