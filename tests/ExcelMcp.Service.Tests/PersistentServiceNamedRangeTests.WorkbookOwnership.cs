using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceNamedRangeTests
{
    [Fact]
    public async Task ExternalName_LeavesLinkedWorkbookUntouchedAndOmitsItsValues()
    {
        const string foreignValue = "content from the other workbook";
        var name = CreateUniqueNamedRangeName();
        Excel.Workbook? foreign = null;
        Excel.Range? foreignCell = null;
        string? foreignName = null;

        var failure = await Record.ExceptionAsync(async () =>
        {
            _fixture.RegisterNamedRangeForCleanup(name);
            _fixture.ExecuteRawVerification((context, _) =>
            {
                Excel.Workbooks? books = null;
                Excel.Sheets? sheets = null;
                Excel.Worksheet? sheet = null;
                Excel.Names? names = null;
                Excel.Name? nativeName = null;
                Excel.Range? linkedRange = null;
                Excel.Worksheet? linkedSheet = null;
                Excel.Workbook? linkedBook = null;
                try
                {
                    books = context.App.Workbooks;
                    foreign = books.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    foreignName = foreign.Name;
                    sheets = foreign.Worksheets;
                    sheet = (Excel.Worksheet)sheets[1];
                    sheet.Name = "LinkedOnly";
                    foreignCell = sheet.Range["B2"];
                    foreignCell.Value2 = foreignValue;
                    context.Book.Activate();
                    names = context.Book.Names;
                    nativeName = names.Add(Name: name, RefersTo: $"='[{foreignName}]LinkedOnly'!$B$2");
                    string definition = nativeName.RefersTo;
                    Assert.Contains(foreignName, definition, StringComparison.Ordinal);
                    linkedRange = nativeName.RefersToRange;
                    linkedSheet = linkedRange.Worksheet;
                    linkedBook = (Excel.Workbook)linkedSheet.Parent;
                    Assert.Same(foreign, linkedBook);
                    Assert.Equal(foreignValue, (object?)linkedRange.Value2);
                }
                finally
                {
                    ComUtilities.Release(ref linkedBook);
                    ComUtilities.Release(ref linkedSheet);
                    ComUtilities.Release(ref linkedRange);
                    ComUtilities.Release(ref nativeName);
                    ComUtilities.Release(ref names);
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                    ComUtilities.Release(ref books);
                }
            });

            string[][] attemptedValues = [["new value"]];
            var readsAndWrites = new (string Command, object Args)[]
            {
                ("range.get-values", new { sheetName = string.Empty, rangeAddress = name }),
                ("range.inspect-cells", new { sheetName = string.Empty, rangeAddress = name }),
                ("range.set-values", new
                {
                    sheetName = string.Empty, rangeAddress = name, values = attemptedValues, overwritePolicy = "allow"
                }),
                ("namedrange.read", new { name }),
                ("namedrange.write", new { name, value = "new value" })
            };
            foreach (var (command, args) in readsAndWrites)
            {
                var response = await _fixture.SendForFailureAsync(command, args);
                Assert.Equal("InvalidInput", response.ErrorCategory);
                Assert.Contains("another workbook", response.ErrorMessage, StringComparison.Ordinal);
            }

            var metadataResponse = _fixture.Send("namedrange.read", new { name, metadataOnly = true });
            Assert.True(metadataResponse.Success, metadataResponse.ErrorMessage);
            using var metadataDocument = JsonDocument.Parse(metadataResponse.Result!);
            var metadata = metadataDocument.RootElement;
            Assert.Equal("ExternalRange", metadata.GetProperty("resolutionStatus").GetString());
            Assert.Equal("NotRead", metadata.GetProperty("valueType").GetString());
            Assert.True(metadata.GetProperty("visible").GetBoolean());
            Assert.Contains(Assert.IsType<string>(foreignName),
                metadata.GetProperty("refersTo").GetString(), StringComparison.Ordinal);
            Assert.False(metadata.TryGetProperty("areaCount", out _));
            Assert.False(metadata.TryGetProperty("value", out _));

            var listed = RequireSuccess(_parameterCommands.List(_fixture.BatchToken));
            var linked = Assert.Single(listed.NamedRanges, item => item.Name == name);
            Assert.Contains(Assert.IsType<string>(foreignName), linked.RefersTo, StringComparison.Ordinal);
            Assert.Null(linked.Value);
            Assert.Equal("Unavailable", linked.ValueType);
            Assert.Contains("another workbook", linked.ValueOmittedReason, StringComparison.Ordinal);
            Assert.Equal(foreignValue, _fixture.ExecuteRawVerification((_, _) => (object?)foreignCell!.Value2));
        });
        var cleanup = Record.Exception(() => _fixture.ExecuteRawVerification((context, _) =>
            PersistentServiceCleanupFailures.Run(() => ComUtilities.Release(ref foreignCell), () =>
            {
                try
                {
                    foreign?.Close(SaveChanges: false);
                }
                finally
                {
                    ComUtilities.Release(ref foreign);
                }
            }, () => context.Book.Activate())));
        PersistentServiceCleanupFailures.ThrowIfFailed(cleanup is null
            ? failure
            : PersistentServiceCleanupFailures.Combine(failure, cleanup));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LocalName_StillSupportsDirectAndSharedRangeOperations(bool worksheetScoped)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var leafName = CreateUniqueNamedRangeName();
        var name = worksheetScoped ? $"{sheetName}!{leafName}" : leafName;
        RequireSuccess(_parameterCommands.Create(batch, name, $"'{sheetName}'!$B$2"));
        _fixture.RegisterNamedRangeForCleanup(name);

        RequireSuccess(_parameterCommands.Write(batch, name, "17"));
        RangeReadAssertions.Number(17, RequireSuccess(_commands.GetValues(batch, string.Empty, name)).Values[0][0]);
        RequireSuccess(_commands.SetValues(batch, string.Empty, name, [[28]],
            overwritePolicy: Sbroenne.ExcelMcp.Core.Commands.Range.OverwritePolicy.Allow));
        RangeReadAssertions.Number(28, _parameterCommands.Read(batch, name).Value);
        RangeReadAssertions.Number(28, RequireSuccess(_commands.GetValues(batch, sheetName, "B2")).Values[0][0]);
    }
}
