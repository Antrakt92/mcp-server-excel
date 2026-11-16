using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "Workbook")]
public sealed class SheetUnprotectPasswordTests
{
    [Theory]
    [InlineData(null, "")]
    [InlineData("", "")]
    [InlineData(" Sample password ", " Sample password ")]
    public void Unprotect_PassesExplicitPasswordAndKeepsItOutOfTheResult(string? password, string expected)
    {
        const string sheetName = "Target";
        var calls = new List<object?>();
        var sheet = ExcelInterfaceProxy.Create<Excel.Worksheet>((method, args) =>
        {
            if (method.Name == "get_Name")
                return sheetName;
            Assert.Equal("Unprotect", method.Name);
            calls.Add(Assert.Single(args!));
            return null;
        });
        var sheets = ExcelInterfaceProxy.Create<Excel.Sheets>((method, _) => method.Name switch
        {
            "get_Count" => 1,
            "get__Default" => sheet,
            _ => throw new NotSupportedException(method.Name)
        }, typeof(ExcelContext).Assembly);
        var workbook = ExcelInterfaceProxy.Create<Excel.Workbook>((method, _) => method.Name == "get_Worksheets"
            ? sheets : throw new NotSupportedException(method.Name), typeof(ExcelContext).Assembly);
        var app = ExcelInterfaceProxy.Create<Excel.Application>((method, _) => method.Name == "get_International"
            ? "General" : throw new NotSupportedException(method.Name), typeof(ExcelContext).Assembly);
        var context = new ExcelContext("fixture.xlsx", app, workbook);
        using var batch = ExcelInterfaceProxy.Create<IExcelBatch>((method, args) => method.Name switch
        {
            nameof(IDisposable.Dispose) => null,
            "get_WorkbookPath" => context.WorkbookPath,
            nameof(IExcelBatch.Execute) => Assert.IsType<Func<ExcelContext, CancellationToken, OperationResult>>(
                args![0])(context, Assert.IsType<CancellationToken>(args[1])),
            _ => throw new NotSupportedException(method.Name)
        });

        var result = new SheetCommands().SetProtection(batch, sheetName, isProtected: false, password);

        Assert.True(result.Success);
        Assert.Equal(expected, Assert.IsType<string>(Assert.Single(calls)));
        var serialized = JsonSerializer.Serialize(result);
        Assert.DoesNotContain("password", serialized, StringComparison.OrdinalIgnoreCase);
        if (expected.Length > 0)
            Assert.DoesNotContain(expected, serialized, StringComparison.Ordinal);
    }
}
