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
