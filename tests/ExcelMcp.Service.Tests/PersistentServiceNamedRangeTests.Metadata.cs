using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceNamedRangeTests
{
    [Fact]
    public void ReadMetadata_PartiallyBrokenUnion_RetainsNativeFormulaAndReportsUnavailableRange()
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        var name = CreateUniqueNamedRangeName();
        var reference = $"'{sheet}'!$A$1,'{sheet}'!$C$3";
        Assert.True(_parameterCommands.Create(batch, name, reference).Success);
        _fixture.RegisterNamedRangeForCleanup(name);

        var beforeResponse = _fixture.Send("namedrange.read", new { name, metadataOnly = true });
        Assert.True(beforeResponse.Success, beforeResponse.ErrorMessage);
        using var beforeDocument = JsonDocument.Parse(beforeResponse.Result!);
        var before = beforeDocument.RootElement;
        Assert.True(before.GetProperty("visible").GetBoolean());
        Assert.Equal("LocalRange", before.GetProperty("resolutionStatus").GetString());
        Assert.Equal(2, before.GetProperty("areaCount").GetInt32());
        Assert.DoesNotContain("#REF!", before.GetProperty("refersTo").GetString(), StringComparison.Ordinal);
        Assert.False(before.TryGetProperty("value", out _));
        Assert.False(before.TryGetProperty("resolutionHResult", out _));

        Assert.True(_commands.DeleteRows(batch, sheet, "1:1").Success);

        var afterResponse = _fixture.Send("namedrange.read", new { name, metadataOnly = true });
        Assert.True(afterResponse.Success, afterResponse.ErrorMessage);
        using var afterDocument = JsonDocument.Parse(afterResponse.Result!);
        var after = afterDocument.RootElement;
        var formula = after.GetProperty("refersTo").GetString();
        Assert.Contains("#REF!", formula, StringComparison.Ordinal);
        Assert.Contains("$C$2", formula, StringComparison.Ordinal);
        Assert.Equal("Unavailable", after.GetProperty("resolutionStatus").GetString());
        Assert.Equal(unchecked((int)0x800A03EC), after.GetProperty("resolutionHResult").GetInt32());
        Assert.False(after.TryGetProperty("areaCount", out _));
        Assert.False(after.TryGetProperty("value", out _));
    }
}
