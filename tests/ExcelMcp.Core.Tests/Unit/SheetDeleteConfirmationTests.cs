using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "Workbook")]
[Trait("RequiresExcel", "false")]
public sealed class SheetDeleteConfirmationTests
{
    [Theory]
    [InlineData("no-op", false)]
    [InlineData("renamed-original", false)]
    [InlineData("same-name-replacement", true)]
    [InlineData("removed", true)]
    [InlineData("verification-fails", false)]
    public void Delete_ConfirmsOriginalWorksheetIdentity(string nativeOutcome, bool deleted)
    {
        const string name = "Target";
        string currentName = name;
        var worksheets = new List<Excel.Worksheet>();
        var deleteCalls = 0;
        var verificationFailure = new InvalidOperationException("Worksheet enumeration failed.");
        var target = ExcelInterfaceProxy.Create<Excel.Worksheet>((method, _) => method.Name switch
        {
            "get_Name" => currentName,
            "Delete" => DeleteTarget(),
            _ => throw new NotSupportedException(method.Name)
        });
        worksheets.Add(target);

        object? DeleteTarget()
        {
            deleteCalls++;
            if (nativeOutcome == "renamed-original")
                currentName = "Renamed Target";
            else if (nativeOutcome == "same-name-replacement")
                worksheets[0] = ExcelInterfaceProxy.Create<Excel.Worksheet>((method, _) => method.Name == "get_Name"
                    ? name : throw new NotSupportedException(method.Name));
            else if (nativeOutcome is "removed" or "verification-fails")
                worksheets.Clear();
            return null;
        }

        var sheets = ExcelInterfaceProxy.Create<Excel.Sheets>((method, args) => method.Name switch
        {
            "get_Count" => worksheets.Count,
            "get__Default" => worksheets[Assert.IsType<int>(args![0]) - 1],
            _ => throw new NotSupportedException(method.Name)
        }, typeof(ExcelContext).Assembly);
        var workbook = ExcelInterfaceProxy.Create<Excel.Workbook>((method, _) => method.Name switch
        {
            "get_Worksheets" when nativeOutcome == "verification-fails" && deleteCalls > 0 => throw verificationFailure,
            "get_Worksheets" => sheets,
            _ => throw new NotSupportedException(method.Name)
        }, typeof(ExcelContext).Assembly);
        var application = ExcelInterfaceProxy.Create<Excel.Application>((method, _) => method.Name == "get_International"
            ? "General" : throw new NotSupportedException(method.Name), typeof(ExcelContext).Assembly);
        var context = new ExcelContext("fixture.xlsx", application, workbook);
        using var batch = ExcelInterfaceProxy.Create<IExcelBatch>((method, args) => method.Name switch
        {
            nameof(IDisposable.Dispose) => null,
            "get_WorkbookPath" => context.WorkbookPath,
            nameof(IExcelBatch.Execute) => Assert.IsType<Func<ExcelContext, CancellationToken, OperationResult>>(
                args![0])(context, Assert.IsType<CancellationToken>(args[1])),
            _ => throw new NotSupportedException(method.Name)
        });

        if (deleted)
        {
            Assert.True(new SheetCommands().Delete(batch, name).Success);
            Assert.DoesNotContain(worksheets, sheet => ReferenceEquals(sheet, target));
        }
        else if (nativeOutcome == "verification-fails")
        {
            var failure = Assert.ThrowsAny<InvalidOperationException>(() => new SheetCommands().Delete(batch, name));
            Assert.Same(verificationFailure, failure.InnerException);
            Assert.Contains("could not confirm", failure.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Empty(worksheets);
        }
        else
        {
            var failure = Assert.Throws<InvalidOperationException>(() => new SheetCommands().Delete(batch, name));
            Assert.Contains("still present", failure.Message, StringComparison.Ordinal);
            Assert.Contains(worksheets, sheet => ReferenceEquals(sheet, target));
        }
        Assert.Equal(1, deleteCalls);
    }
}
