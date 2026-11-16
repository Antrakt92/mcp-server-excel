using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "false")]
public sealed class RangeMergeCleanupTests
{
    [Fact]
    public void MergeMarksCompletionUnconfirmedWhenRestoringAlertsFails()
    {
        var restoreFailure = new InvalidOperationException("Alerts could not be restored.");
        bool merged = false;
        var app = ExcelInterfaceProxy.Create<Excel.Application>((method, args) => method.Name switch
        {
            "get_DisplayAlerts" => true,
            "set_DisplayAlerts" when Assert.IsType<bool>(args![0]) => throw restoreFailure,
            "set_DisplayAlerts" => null,
            _ => throw new NotSupportedException(method.Name)
        });
        var range = ExcelInterfaceProxy.Create<Excel.Range>((method, _) =>
        {
            Assert.Equal("Merge", method.Name);
            merged = true;
            return null;
        });

        var failure = Assert.Throws<OperationCompletionUnconfirmedException>(() =>
            RangeCommands.MergeWithAuthorizedDiscard(app, range, OverwritePolicy.Allow));

        Assert.True(merged);
        Assert.Same(restoreFailure, failure.InnerException);
    }
}
