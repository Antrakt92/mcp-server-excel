using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceNamedRangeTests
{
    [Fact]
    public async Task Read_FullSheetNameRejectsOversizedDataBeforeReadingIt()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var name = CreateUniqueNamedRangeName();
        _fixture.RegisterNamedRangeForCleanup(name);
        AddName(_fixture, name, $"'{sheetName}'!$A$1:$XFD$1048576");
        var rejected = await _fixture.SendForFailureAsync("namedrange.read", new { name });

        Assert.Equal("InvalidInput", rejected.ErrorCategory);
        Assert.Contains("cell read limit", rejected.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("get-values-page", rejected.ErrorMessage, StringComparison.Ordinal);
        var metadata = _parameterCommands.Read(_fixture.BatchToken, name, metadataOnly: true);
        Assert.Equal("LocalRange", metadata.ResolutionStatus);
        Assert.Equal("NotRead", metadata.ValueType);
        Assert.Equal(1, metadata.AreaCount);
        Assert.Null(metadata.Value);
    }

    [Fact]
    public void ReadAndList_DistinguishNativeErrorsFromNumericErrorCodes()
    {
        const int number = -2146826281;
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var numericName = CreateUniqueNamedRangeName();
        var errorName = CreateUniqueNamedRangeName();
        var gridName = CreateUniqueNamedRangeName();
        _fixture.RegisterNamedRangeForCleanup(numericName);
        _fixture.RegisterNamedRangeForCleanup(errorName);
        _fixture.RegisterNamedRangeForCleanup(gridName);
        AddName(_fixture, numericName, $"'{sheetName}'!$D$8");
        AddName(_fixture, errorName, $"'{sheetName}'!$E$8");
        AddName(_fixture, gridName, $"'{sheetName}'!$D$8:$E$8");
        RequireSuccess(_commands.SetValues(batch, sheetName, "D8", [[number]]));
        RequireSuccess(_commands.SetFormulas(batch, sheetName, "E8", [["=1/0"]]));

        var numeric = _parameterCommands.Read(batch, numericName);
        var error = _parameterCommands.Read(batch, errorName);
        var grid = _parameterCommands.Read(batch, gridName);
        var listed = RequireSuccess(_parameterCommands.List(batch));

        Assert.Equal("Double", numeric.ValueType);
        RangeReadAssertions.Number(number, numeric.Value);
        Assert.Equal("ExcelError", error.ValueType);
        Assert.Equal("#DIV/0!", error.Value);
        Assert.Equal("Array", grid.ValueType);
        AssertGrid(grid.Value);

        var numericPreview = Assert.Single(listed.NamedRanges, item => item.Name == numericName);
        Assert.Equal("Double", numericPreview.ValueType);
        RangeReadAssertions.Number(number, numericPreview.Value);
        var errorPreview = Assert.Single(listed.NamedRanges, item => item.Name == errorName);
        Assert.Equal("ExcelError", errorPreview.ValueType);
        Assert.Equal("#DIV/0!", errorPreview.Value);
        var gridPreview = Assert.Single(listed.NamedRanges, item => item.Name == gridName);
        Assert.Equal("Array", gridPreview.ValueType);
        AssertGrid(gridPreview.Value);

        static void AssertGrid(object? value)
        {
            var rows = Assert.IsType<JsonElement>(value);
            var row = Assert.Single(rows.EnumerateArray());
            Assert.Equal(2, row.GetArrayLength());
            Assert.Equal((double)number, row[0].GetDouble());
            Assert.Equal("#DIV/0!", row[1].GetString());
        }
    }

    [Fact]
    public void Read_MultiTileNamePreservesArrayAndScalarTyping()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var name = CreateUniqueNamedRangeName();
        var scalar = CreateUniqueNamedRangeName();
        _fixture.RegisterNamedRangeForCleanup(name);
        _fixture.RegisterNamedRangeForCleanup(scalar);
        AddName(_fixture, name, $"'{sheetName}'!$A$1:$B$1001");
        AddName(_fixture, scalar, $"'{sheetName}'!$A$1");
        RequireSuccess(_commands.SetValues(batch, sheetName, "A1", [[123]]));
        RequireSuccess(_commands.SetValues(batch, sheetName, "B500", [[true]]));
        RequireSuccess(_commands.SetValues(batch, sheetName, "A501", [["next tile"]]));
        RequireSuccess(_commands.SetValues(batch, sheetName, "B1001", [["last"]]));

        var singleValue = _parameterCommands.Read(batch, scalar);
        var grid = _parameterCommands.Read(batch, name);

        Assert.Equal("Double", singleValue.ValueType);
        RangeReadAssertions.Number(123, singleValue.Value);
        Assert.Equal("Array", grid.ValueType);
        var rows = Assert.IsType<JsonElement>(grid.Value);
        Assert.Equal(1001, rows.GetArrayLength());
        Assert.All(rows.EnumerateArray(), row => Assert.Equal(2, row.GetArrayLength()));
        Assert.Equal(123d, rows[0][0].GetDouble());
        Assert.Equal(JsonValueKind.Null, rows[0][1].ValueKind);
        Assert.True(rows[499][1].GetBoolean());
        Assert.Equal("next tile", rows[500][0].GetString());
        Assert.Equal("last", rows[1000][1].GetString());
    }
}
