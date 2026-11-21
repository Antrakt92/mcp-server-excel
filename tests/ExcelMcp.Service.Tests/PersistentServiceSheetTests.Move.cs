using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceSheetTests
{
    [Fact]
    public void Move_WithBeforeSheet_RepositionsSheet() => AssertMove("before");

    [Fact]
    public void Move_WithAfterSheet_RepositionsSheet() => AssertMove("after");

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Move_WithBlankBeforeAndValidAfter_RepositionsSheet(string blankBefore) => AssertMove("after", blankBefore);

    [Fact]
    public void Move_NoPositionSpecified_MovesToEnd() => AssertMove("end");

    [Fact]
    public void Move_NoPositionSpecified_MovesAfterTrailingChartTab()
    {
        var path = CreateMarkedWorkbook(
            nameof(Move_NoPositionSpecified_MovesAfterTrailingChartTab), "Workbook", "TransferSheet");
        using (var setup = ExcelSession.BeginBatch(path))
        {
            setup.Execute((context, _) =>
            {
                Excel.Worksheet? lastWorksheet = null;
                Excel.Sheets? charts = null;
                Excel.Chart? chart = null;
                Excel.Sheets? tabs = null;
                object? lastTab = null;
                try
                {
                    lastWorksheet = ComUtilities.FindSheet(context.Book, "Sheet1");
                    Assert.NotNull(lastWorksheet);
                    charts = context.Book.Charts;
                    chart = (Excel.Chart)charts.Add(Before: lastWorksheet);
                    chart.Name = "ChartTab";
                    chart.Move(After: lastWorksheet);
                    tabs = context.Book.Sheets;
                    Assert.Equal(3, tabs.Count);
                    lastTab = tabs[3];
                    Assert.Equal("ChartTab", ((Excel.Chart)lastTab).Name);
                }
                finally
                {
                    ComUtilities.Release(ref lastTab);
                    ComUtilities.Release(ref tabs);
                    ComUtilities.Release(ref chart);
                    ComUtilities.Release(ref charts);
                    ComUtilities.Release(ref lastWorksheet);
                }
            });
            setup.Save();
        }

        using (var service = new ExcelMcpService())
        {
            using var opened = ReadCrossFileServiceResult(service, "session.open", null, new { filePath = path });
            var sessionId = opened.RootElement.GetProperty("sessionId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(sessionId));
            using var moved = ReadCrossFileServiceResult(service, "sheet.move", sessionId, new { sheetName = "TransferSheet" });
            using var closed = ReadCrossFileServiceResult(service, "session.close", sessionId, new { save = true });
        }

        using var reopened = ExcelSession.BeginBatch(path);
        reopened.Execute((context, _) =>
        {
            AssertWorkbookState(context.Book, ["Sheet1", "TransferSheet"], "TransferSheet");
            Excel.Sheets? tabs = null;
            object? nextToLast = null;
            object? last = null;
            try
            {
                tabs = context.Book.Sheets;
                Assert.Equal(3, tabs.Count);
                nextToLast = tabs[2];
                last = tabs[3];
                Assert.Equal("TransferSheet", (last as Excel.Worksheet)?.Name);
                Assert.Equal("ChartTab", ((Excel.Chart)nextToLast).Name);
            }
            finally
            {
                ComUtilities.Release(ref last);
                ComUtilities.Release(ref nextToLast);
                ComUtilities.Release(ref tabs);
            }
        });
    }

    private void AssertMove(string position, string? blankBefore = null)
    {
        var source = CreateSeededSheet("Move");
        var target = CreateSeededSheet("Target");
        if (position == "after")
            RequireSuccess(_sheetCommands.Move(_fixture.BatchToken, source, beforeSheet: target));
        var before = ReadCurrentSheetNames();
        var sourceState = CaptureSheet(source);
        var targetState = CaptureSheet(target);
        var expected = before.Where(name => name != source).ToList();
        var index = position == "end" ? expected.Count : expected.IndexOf(target) + (position == "after" ? 1 : 0);
        Assert.InRange(index, 0, expected.Count);
        expected.Insert(index, source);
        if (position == "after")
            Assert.False(expected.SequenceEqual(before), "Arrange the sheets so that moving after the anchor changes their order.");
        RequireSuccess(_sheetCommands.Move(_fixture.BatchToken, source,
            beforeSheet: position == "before" ? target : blankBefore,
            afterSheet: position == "after" ? target : null));
        Assert.Equal(expected, ReadCurrentSheetNames());
        Assert.Equal(sourceState, CaptureSheet(source));
        Assert.Equal(targetState, CaptureSheet(target));
    }

    [Fact]
    public void Move_BothBeforeAndAfter_ThrowsException()
    {
        var source = CreateSeededSheet("Move");
        var target = CreateSeededSheet("Target");
        var names = ReadCurrentSheetNames();
        var sourceState = CaptureSheet(source);
        var targetState = CaptureSheet(target);
        var error = Assert.Throws<ArgumentException>(() =>
            _sheetCommands.Move(_fixture.BatchToken, source, beforeSheet: target, afterSheet: target));
        Assert.Contains("both beforeSheet and afterSheet", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(names, ReadCurrentSheetNames());
        Assert.Equal(sourceState, CaptureSheet(source));
        Assert.Equal(targetState, CaptureSheet(target));
    }

    [Fact]
    public void Move_NonExistentSheet_ThrowsException() => AssertMissingMove(missingSource: true);

    [Fact]
    public void Move_NonExistentTargetSheet_ThrowsException() => AssertMissingMove(missingSource: false);

    private void AssertMissingMove(bool missingSource)
    {
        var source = CreateSeededSheet("Move");
        var target = CreateSeededSheet("Target");
        var names = ReadCurrentSheetNames();
        var sourceState = CaptureSheet(source);
        var targetState = CaptureSheet(target);
        var error = Assert.Throws<InvalidOperationException>(() =>
            _sheetCommands.Move(_fixture.BatchToken, missingSource ? "MissingSheet" : source,
                beforeSheet: missingSource ? target : "MissingSheet"));
        Assert.Contains("not found", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(names, ReadCurrentSheetNames());
        Assert.Equal(sourceState, CaptureSheet(source));
        Assert.Equal(targetState, CaptureSheet(target));
    }
}
