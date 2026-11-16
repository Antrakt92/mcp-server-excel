using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "NamedRanges")]
[Trait("RequiresExcel", "false")]
public sealed class NamedRangeWorkbookOwnershipTests
{
    [Theory]
    [InlineData("InputParameter", true)]
    [InlineData("Scope!InputParameter", true)]
    [InlineData("InputParameter", false)]
    [InlineData("Scope!InputParameter", false)]
    public void ResolveRange_RequiresTheExactWorkbook(string requestedName, bool sameWorkbook)
    {
        var foreign = ExcelInterfaceProxy.Create<Excel.Workbook>((method, _) => throw new NotSupportedException(method.Name));
        Excel.Names? names = null;
        var selectedWorkbook = ExcelInterfaceProxy.Create<Excel.Workbook>((method, _) => method.Name == "get_Names"
            ? names : throw new NotSupportedException(method.Name));
        var worksheet = ExcelInterfaceProxy.Create<Excel.Worksheet>((method, _) => method.Name == "get_Parent"
            ? sameWorkbook ? selectedWorkbook : foreign : throw new NotSupportedException(method.Name));
        var range = ExcelInterfaceProxy.Create<Excel.Range>((method, _) => method.Name == "get_Worksheet"
            ? worksheet : throw new NotSupportedException(method.Name));
        var name = ExcelInterfaceProxy.Create<Excel.Name>((method, _) => method.Name == "get_RefersToRange"
            ? range : throw new NotSupportedException(method.Name));
        names = ExcelInterfaceProxy.Create<Excel.Names>((method, args) =>
        {
            Assert.Equal("Item", method.Name);
            Assert.Equal(requestedName, args![0]);
            return name;
        });

        if (sameWorkbook)
        {
            Assert.Same(range, RangeHelpers.ResolveRange(selectedWorkbook, string.Empty, requestedName));
        }
        else
        {
            var error = Assert.Throws<OperationFailureException>(() =>
                RangeHelpers.ResolveRange(selectedWorkbook, string.Empty, requestedName));
            Assert.Equal(OperationFailureCategory.InvalidInput, error.ErrorCategory);
            Assert.Contains("another workbook", error.Message, StringComparison.Ordinal);
        }
    }
}
