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
    [InlineData(true, null, false, 0, 0)]
    [InlineData(null, true, false, 0, 0)]
    [InlineData(false, false, true, 0, 0)]
    [InlineData(null, false, true, 2, 0)]
    [InlineData(false, null, true, 0, 3)]
    [InlineData(null, null, true, 2, 3)]
    public void GetSpecialCells_ChecksOnlyMixedDimensions(
        bool? rowsHidden, bool? columnsHidden, bool anyVisible, int rowReads, int columnReads)
    {
        var fixture = new VisibilityProbe(rowsHidden, columnsHidden);
        if (anyVisible)
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

        Assert.Equal(rowsHidden is true ? 0 : 1, fixture.ColumnReads);
        Assert.Equal(Enumerable.Range(1, rowReads), fixture.RowMemberIndexes);
        Assert.Equal(Enumerable.Range(1, columnReads), fixture.ColumnMemberIndexes);
    }

    [Theory]
    [InlineData("EntireRow")]
    [InlineData("RowsHidden")]
    [InlineData("EntireColumn")]
    [InlineData("ColumnsHidden")]
    [InlineData("RowMemberHidden")]
    [InlineData("ColumnMemberHidden")]
    public void GetSpecialCells_ReportsBothExcelErrorsWhenVisibilityCannotBeVerified(string failedMember)
    {
        var fixture = new VisibilityProbe(null, null, failedMember);
        var error = Assert.Throws<AggregateException>(fixture.Run);

        Assert.Collection(error.InnerExceptions,
            native => Assert.Same(fixture.NativeFailure, native),
            verification => Assert.Same(fixture.ProbeFailure, verification));
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2201",
        Justification = "Exercises the ambiguous native SpecialCells error without launching Excel.")]
    private sealed class VisibilityProbe(bool? rowsHidden, bool? columnsHidden, string? failedMember = null)
    {
        public COMException NativeFailure { get; } = new("Synthetic SpecialCells no-match code.", unchecked((int)0x800A03EC));
        public COMException ProbeFailure { get; } = new("Synthetic visibility read error.", unchecked((int)0x800A03EC));
        public List<int> RowMemberIndexes { get; } = [];
        public List<int> ColumnMemberIndexes { get; } = [];
        public int ColumnReads { get; private set; }

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
            var rows = Dimension(rows: true, rowsHidden, 2, RowMemberIndexes);
            var columns = Dimension(rows: false, columnsHidden, 3, ColumnMemberIndexes);
            scope = ExcelInterfaceProxy.Create<Excel.Range>((method, _) => method.Name switch
            {
                "get_Parent" => sheet,
                "get_Address" => address,
                "get_Areas" => areas,
                "get_CountLarge" => 6d,
                "get_EntireRow" => Probe("EntireRow", rows),
                "get_EntireColumn" => ReadColumns(columns),
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

        private Excel.Range Dimension(bool rows, bool? hidden, int count, List<int> memberIndexes)
        {
            var members = ExcelInterfaceProxy.Create<Excel.Range>((method, args) =>
            {
                if (method.Name == "get_Count")
                    return count;
                Assert.Equal("get__Default", method.Name);
                int index = Assert.IsType<int>(args![0]);
                memberIndexes.Add(index);
                return ExcelInterfaceProxy.Create<Excel.Range>((memberMethod, _) =>
                {
                    Assert.Equal("get_Hidden", memberMethod.Name);
                    return Probe(rows ? "RowMemberHidden" : "ColumnMemberHidden", index < count);
                });
            });
            return ExcelInterfaceProxy.Create<Excel.Range>((method, _) => method.Name switch
            {
                "get_Hidden" => Probe(rows ? "RowsHidden" : "ColumnsHidden", hidden),
                "get_Rows" when rows => members,
                "get_Columns" when !rows => members,
                _ => throw new NotSupportedException(method.Name)
            });
        }

        private object? ReadColumns(Excel.Range columns)
        {
            ColumnReads++;
            return Probe("EntireColumn", columns);
        }

        private object? Probe(string member, object? value) => member == failedMember ? throw ProbeFailure : value;
    }

    private static T Properties<T>(params (string Name, object? Value)[] properties) where T : class =>
        ExcelInterfaceProxy.Create<T>((method, _) => properties.Single(property => property.Name == method.Name).Value);
}
