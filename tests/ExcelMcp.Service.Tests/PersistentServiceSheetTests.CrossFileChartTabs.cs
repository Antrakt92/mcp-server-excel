using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceSheetTests
{
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    [InlineData(null)]
    public void CopyToFile_WithChartTabBeforeAnchor_RenamesOnlyTheNewWorksheet(bool? before)
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(CopyToFile_WithChartTabBeforeAnchor_RenamesOnlyTheNewWorksheet),
            "TransferSheet", "Guard", targetDuplicateGuard: true);
        using (var batch = ExcelSession.BeginBatch(targetFile))
        {
            batch.Execute((context, _) =>
            {
                Excel.Worksheet? anchor = null;
                Excel.Sheets? charts = null;
                Excel.Chart? chart = null;
                try
                {
                    anchor = ComUtilities.FindSheet(context.Book, before.HasValue ? "Guard" : "Sheet1");
                    Assert.NotNull(anchor);
                    charts = context.Book.Charts;
                    chart = before.HasValue
                        ? (Excel.Chart)charts.Add(Before: anchor)
                        : (Excel.Chart)charts.Add(After: anchor);
                    chart.Name = "ChartTab";
                }
                finally
                {
                    ComUtilities.Release(ref chart);
                    ComUtilities.Release(ref charts);
                    ComUtilities.Release(ref anchor);
                }
            });
            batch.Save();
        }

        RequireSuccess(_sheetCommands.CopyToFile(sourceFile, "TransferSheet", targetFile, "Copied",
            beforeSheet: before == true ? "Guard" : null, afterSheet: before == false ? "Guard" : null));

        using var reopened = ExcelSession.BeginBatch(sourceFile, targetFile);
        reopened.Execute((_, _) =>
        {
            var source = reopened.GetWorkbook(sourceFile);
            var target = reopened.GetWorkbook(targetFile);
            AssertWorkbookState(source, ["TransferSheet", "Sheet1"], "TransferSheet");
            string[] expectedNames = before switch
            {
                true => ["Copied", "Guard", "Sheet1"],
                false => ["Guard", "Copied", "Sheet1"],
                null => ["Guard", "Sheet1", "Copied"]
            };
            AssertWorkbookState(target, expectedNames, "Copied");
            Excel.Sheets? worksheets = null;
            Excel.Sheets? charts = null;
            Excel.Sheets? tabs = null;
            Excel.Chart? chartTab = null;
            Excel.Worksheet? lastWorksheet = null;
            try
            {
                worksheets = target.Worksheets;
                AssertSheetCells(worksheets, "Guard", "duplicate guard", 12, 15, "=B1+3");
                charts = target.Charts;
                chartTab = (Excel.Chart)charts["ChartTab"];
                Assert.Equal("ChartTab", chartTab.Name);
                if (!before.HasValue)
                {
                    tabs = target.Sheets;
                    lastWorksheet = tabs[tabs.Count] as Excel.Worksheet;
                    Assert.True(lastWorksheet is not null, $"Last tab is not a worksheet. Chart tab index: {chartTab.Index}.");
                    Assert.Equal("Copied", lastWorksheet.Name);
                }
                else
                {
                    Assert.Equal(1, chartTab.Index);
                }
            }
            finally
            {
                ComUtilities.Release(ref lastWorksheet);
                ComUtilities.Release(ref chartTab);
                ComUtilities.Release(ref tabs);
                ComUtilities.Release(ref charts);
                ComUtilities.Release(ref worksheets);
            }
        });
    }
}
