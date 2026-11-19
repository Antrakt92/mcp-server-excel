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
    [InlineData("same-name-replacement", false)]
    [InlineData("removed", false)]
    [InlineData("verification-fails", false)]
    [InlineData("removed", true)]
    [InlineData("same-name-replacement", true)]
    [InlineData("no-op", true)]
    [InlineData("renamed-original", true)]
    [InlineData("delete-fails", true)]
    [InlineData("show-fails-after-change", true)]
    [InlineData("cancelled-after-show", true)]
    [InlineData("rollback-fails", true)]
    [InlineData("verification-fails", true)]
    [InlineData("protected", true)]
    public void Delete_ConfirmsOriginalIdentityAndRestoresTemporaryVisibility(string nativeOutcome, bool veryHidden)
    {
        const string name = "Target";
        const Excel.XlSheetVisibility visible = Excel.XlSheetVisibility.xlSheetVisible;
        const Excel.XlSheetVisibility veryHiddenValue = Excel.XlSheetVisibility.xlSheetVeryHidden;
        var initialVisibility = veryHidden ? veryHiddenValue : visible;
        var currentVisibility = initialVisibility;
        string currentName = name;
        var worksheets = new List<Excel.Worksheet>();
        var visibilityWrites = new List<Excel.XlSheetVisibility>();
        var deleteCalls = 0;
        using var cancellation = new CancellationTokenSource();
        var verificationFailure = new InvalidOperationException("Worksheet enumeration failed.");
        var deleteFailure = new InvalidOperationException("Native worksheet Delete failed.");
        var showFailure = new InvalidOperationException("Native worksheet visibility write failed.");
        var rollbackFailure = new InvalidOperationException("Native visibility restoration failed.");
        var target = ExcelInterfaceProxy.Create<Excel.Worksheet>((method, args) => method.Name switch
        {
            "get_Name" => currentName,
            "get_Visible" => Enum.ToObject(method.ReturnType, (int)currentVisibility),
            "set_Visible" => SetVisibility((Excel.XlSheetVisibility)Convert.ToInt32(args![0], System.Globalization.CultureInfo.InvariantCulture)),
            "Delete" => DeleteTarget(),
            _ => throw new NotSupportedException(method.Name)
        });
        worksheets.Add(target);

        object? SetVisibility(Excel.XlSheetVisibility value)
        {
            visibilityWrites.Add(value);
            if (value == veryHiddenValue && nativeOutcome == "rollback-fails")
                throw rollbackFailure;

            currentVisibility = value;
            if (value == visible && nativeOutcome == "show-fails-after-change")
                throw showFailure;
            if (value == visible && nativeOutcome == "cancelled-after-show")
                cancellation.Cancel();
            return null;
        }

        object? DeleteTarget()
        {
            deleteCalls++;
            Assert.Equal(nativeOutcome == "protected" ? veryHiddenValue : visible, currentVisibility);
            if (nativeOutcome is "delete-fails" or "rollback-fails" or "protected")
                throw deleteFailure;
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
            "get_ProtectStructure" => nativeOutcome == "protected",
            _ => throw new NotSupportedException(method.Name)
        });
        var application = ExcelInterfaceProxy.Create<Excel.Application>((method, _) => method.Name == "get_International"
            ? "General" : throw new NotSupportedException(method.Name), typeof(ExcelContext).Assembly);
        var context = new ExcelContext("fixture.xlsx", application, workbook);
        using var batch = ExcelInterfaceProxy.Create<IExcelBatch>((method, args) => method.Name switch
        {
            nameof(IDisposable.Dispose) => null,
            "get_WorkbookPath" => context.WorkbookPath,
            nameof(IExcelBatch.Execute) => Assert.IsType<Func<ExcelContext, CancellationToken, OperationResult>>(
                args![0])(context, cancellation.Token),
            _ => throw new NotSupportedException(method.Name)
        });

        var commands = new SheetCommands();
        switch (nativeOutcome)
        {
            case "removed" or "same-name-replacement":
                Assert.True(commands.Delete(batch, name).Success);
                Assert.DoesNotContain(worksheets, sheet => ReferenceEquals(sheet, target));
                break;
            case "cancelled-after-show":
                var canceled = Assert.Throws<OperationCanceledException>(() => commands.Delete(batch, name));
                Assert.Equal(cancellation.Token, canceled.CancellationToken);
                break;
            case "delete-fails" or "protected":
                Assert.Same(deleteFailure, Assert.Throws<InvalidOperationException>(() => commands.Delete(batch, name)));
                break;
            case "show-fails-after-change":
                Assert.Same(showFailure, Assert.Throws<InvalidOperationException>(() => commands.Delete(batch, name)));
                break;
            case "rollback-fails":
                var rollback = Assert.Throws<AggregateException>(() => commands.Delete(batch, name));
                Assert.Equal(new[] { deleteFailure, rollbackFailure }, rollback.InnerExceptions);
                break;
            case "verification-fails":
                var failure = Record.Exception(() => commands.Delete(batch, name));
                if (veryHidden)
                {
                    var aggregate = Assert.IsType<AggregateException>(failure);
                    Assert.Equal(2, aggregate.InnerExceptions.Count);
                    Assert.Same(verificationFailure, aggregate.InnerExceptions[1]);
                    failure = aggregate.InnerExceptions[0];
                }
                Assert.IsAssignableFrom<InvalidOperationException>(failure);
                Assert.Same(verificationFailure, failure.InnerException);
                Assert.Contains("could not confirm", failure.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Empty(worksheets);
                break;
            default:
                var refused = Assert.Throws<InvalidOperationException>(() => commands.Delete(batch, name));
                Assert.Contains("still present", refused.Message, StringComparison.Ordinal);
                Assert.Contains(worksheets, sheet => ReferenceEquals(sheet, target));
                break;
        }

        Assert.Equal(nativeOutcome is "show-fails-after-change" or "cancelled-after-show" ? 0 : 1, deleteCalls);
        var expectedWrites = !veryHidden || nativeOutcome == "protected"
            ? Array.Empty<Excel.XlSheetVisibility>()
            : nativeOutcome is "removed" or "same-name-replacement" or "verification-fails"
                ? [visible]
                : [visible, veryHiddenValue];
        Assert.Equal(expectedWrites, visibilityWrites);
        if (worksheets.Contains(target) && nativeOutcome != "rollback-fails")
            Assert.Equal(initialVisibility, currentVisibility);
    }
}
