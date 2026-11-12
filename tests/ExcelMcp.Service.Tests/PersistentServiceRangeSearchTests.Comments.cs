using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceRangeSearchTests
{
    [Theory]
    [InlineData(false, 2)]
    [InlineData(true, 3)]
    public void Find_CommentsIncludesNotesAndThreadsWithoutDoubleCounting(bool searchValues, int count)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        const string marker = "audit-find-comment";
        Assert.True(_commands.SetValues(batch, sheetName, "A1:A3", [[marker], ["unrelated"], [marker]]).Success);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            Excel.Comment? note = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                cell = sheet.Range["A1"];
                note = cell.AddComment(marker);
            }
            finally
            {
                ComUtilities.Release(ref note);
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
        var links = _fixture.CreateCommands<IRangeLinkCommands>();
        Assert.True(links.AddThreadedComment(batch, sheetName, "A2", marker).Success);

        var result = _commands.Find(batch, sheetName, "A1:A3", marker, new FindOptions
        {
            SearchFormulas = false,
            SearchValues = searchValues,
            SearchComments = true
        }, maxMatches: 3);

        AssertFindCoverage(result, count, count);
        Assert.Equal(Enumerable.Range(1, count), result.MatchingCells.Select(cell => cell.Row).Order());
    }
}
