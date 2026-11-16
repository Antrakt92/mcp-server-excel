using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceSheetTests
{
    [Fact]
    public void Copy_WithLastWorksheetHidden_RenamesOnlyTheNewWorksheet()
    {
        var batch = _fixture.BatchToken;
        var source = CreateSeededSheet("Source");
        var guard = CreateSeededSheet("Guard");
        RequireSuccess(_sheetCommands.Move(batch, guard, afterSheet: "Sheet1"));
        RequireSuccess(_sheetCommands.Hide(batch, guard));
        var before = ReadCurrentSheetNames();
        Assert.Equal(guard, before[^1]);
        var sourceContents = CaptureSheet(source);
        var guardContents = CaptureSheet(guard);
        var copied = $"Copy_{Guid.NewGuid():N}"[..31];

        try
        {
            RequireSuccess(_sheetCommands.Copy(batch, source, copied));
        }
        finally
        {
            foreach (var added in ReadCurrentSheetNames().Except(before, StringComparer.Ordinal))
                _fixture.RegisterSheetForCleanup(added);
        }

        var after = RequireSuccess(_sheetCommands.List(batch)).Worksheets;
        Assert.Equal(before[..^1].Append(copied).Append(guard), after.Select(sheet => sheet.Name));
        Assert.False(after[^1].Visible);
        Assert.Equal(sourceContents, CaptureSheet(source));
        Assert.Equal(sourceContents, CaptureSheet(copied));
        Assert.Equal(guardContents, CaptureSheet(guard));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CopyToFile_WithLastWorksheetHidden_RenamesOnlyTheNewWorksheet(bool explicitAfter)
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(CopyToFile_WithLastWorksheetHidden_RenamesOnlyTheNewWorksheet),
            "TransferSheet", "Guard", targetDuplicateGuard: true);
        using (var batch = ExcelSession.BeginBatch(targetFile))
        {
            batch.Execute((context, _) =>
            {
                Excel.Worksheet? hidden = null;
                try
                {
                    hidden = ComUtilities.FindSheet(context.Book, "Sheet1");
                    Assert.NotNull(hidden);
                    hidden.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                }
                finally { ComUtilities.Release(ref hidden); }
            });
            batch.Save();
        }

        RequireSuccess(_sheetCommands.CopyToFile(sourceFile, "TransferSheet", targetFile, "Copied",
            afterSheet: explicitAfter ? "Sheet1" : null));

        using var reopened = ExcelSession.BeginBatch(sourceFile, targetFile);
        reopened.Execute((_, _) =>
        {
            AssertWorkbookState(reopened.GetWorkbook(sourceFile), ["TransferSheet", "Sheet1"], "TransferSheet");
            var target = reopened.GetWorkbook(targetFile);
            AssertWorkbookState(target, ["Guard", "Copied", "Sheet1"], "Copied");
            Excel.Sheets? worksheets = null;
            Excel.Worksheet? hidden = null;
            try
            {
                worksheets = target.Worksheets;
                AssertSheetCells(worksheets, "Guard", "duplicate guard", 12, 15, "=B1+3");
                hidden = (Excel.Worksheet)worksheets["Sheet1"];
                Assert.Equal(Excel.XlSheetVisibility.xlSheetHidden, hidden.Visible);
            }
            finally
            {
                ComUtilities.Release(ref hidden);
                ComUtilities.Release(ref worksheets);
            }
        });
    }
}
