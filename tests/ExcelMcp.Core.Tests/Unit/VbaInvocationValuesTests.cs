using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "VBA")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class VbaInvocationValuesTests
{
    [Fact]
    public void Arguments_PreserveExplicitTypesNullAndMissing()
    {
        var arguments = JsonSerializer.Deserialize<List<VbaArgument>>("""
            [{"Type":"string","Value":"=literal"},{"Type":"number","Value":-12.5},{"Type":"boolean","Value":false},{"Type":"null"},{"Type":"missing"}]
            """);
        var values = VbaInvocationValues.Arguments(arguments);
        Assert.Equal("=literal", values[0]);
        Assert.Equal(-12.5, values[1]);
        Assert.Equal(false, values[2]);
        Assert.Same(DBNull.Value, values[3]);
        Assert.Same(Type.Missing, values[4]);
        Assert.All(values.Skip(5), value => Assert.Same(Type.Missing, value));
    }

    [Theory]
    [InlineData("number", "12")]
    [InlineData("boolean", "true")]
    [InlineData("null", "not null")]
    [InlineData("missing", "not missing")]
    [InlineData("unknown", null)]
    public void Arguments_RejectCoercionAndUnknownKinds(string type, object? value) =>
        Assert.Throws<ArgumentException>(() => VbaInvocationValues.Arguments([new() { Type = type, Value = value }]));

    [Fact]
    public void Arguments_RejectTooManyAndNonfiniteNumbers()
    {
        Assert.Throws<ArgumentException>(() => VbaInvocationValues.Arguments(Enumerable.Repeat(new VbaArgument { Type = "missing" }, 31).ToList()));
        Assert.Throws<ArgumentException>(() => VbaInvocationValues.Arguments([new() { Type = "number", Value = double.NaN }]));
    }

    [Fact]
    public void Result_ArrayPreservesBoundsOrderNullAndFalse()
    {
        var input = Array.CreateInstance(typeof(object), [2, 2], [-1, 3]);
        input.SetValue("=literal", -1, 3);
        input.SetValue(DBNull.Value, -1, 4);
        input.SetValue(false, 0, 3);
        input.SetValue(-12.5, 0, 4);
        var result = VbaInvocationValues.Result(input);
        Assert.True(result.Success);
        Assert.Equal("array", result.ReturnKind);
        Assert.Equal("=literal", result.Values![0][0]);
        Assert.Null(result.Values[0][1]);
        Assert.Equal(false, result.Values[1][0]);
        Assert.Equal(-12.5, result.Values[1][1]);
    }

    [Fact]
    public void Result_NullAndScalarsRemainTyped()
    {
        Assert.Equal("null", VbaInvocationValues.Result(null).ReturnKind);
        Assert.Null(VbaInvocationValues.Result(DBNull.Value).Value);
        Assert.Equal("boolean", VbaInvocationValues.Result(false).ReturnKind);
        Assert.Equal("string", VbaInvocationValues.Result("=text").ReturnKind);
        Assert.Equal("number", VbaInvocationValues.Result(42d).ReturnKind);
        Assert.Throws<InvalidOperationException>(() => VbaInvocationValues.Result(new object()));
        Assert.Throws<InvalidOperationException>(() => VbaInvocationValues.Result(new object[2]));
    }

    [Fact]
    public void Result_RejectsOversizedArrayBeforeConvertingCells()
    {
        var returned = new object[101, 100];
        returned[0, 0] = new object();
        var error = Assert.Throws<InvalidOperationException>(() => VbaInvocationValues.Result(returned));
        Assert.Contains("10000", error.Message, StringComparison.Ordinal);
        Assert.Contains("Macro execution completed", error.Message, StringComparison.Ordinal);
        Assert.Contains("not rolled back", error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Result_RejectsScalarAndArrayTextBudgetWithoutReturningContent(bool array)
    {
        string text = "synthetic-private-marker" + new string('x', 300_000);
        object returned = array ? new object[,] { { text } } : text;
        var error = Assert.Throws<InvalidOperationException>(() => VbaInvocationValues.Result(returned));
        Assert.Contains("4 MiB", error.Message, StringComparison.Ordinal);
        Assert.Contains("Macro execution completed", error.Message, StringComparison.Ordinal);
        Assert.Contains("not rolled back", error.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("synthetic-private-marker", error.Message, StringComparison.Ordinal);
    }
}
