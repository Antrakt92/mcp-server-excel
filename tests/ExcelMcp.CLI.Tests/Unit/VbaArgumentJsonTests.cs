using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "VBA")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
[Collection("Sequential")]
public sealed class VbaArgumentJsonTests
{
    [Theory]
    [InlineData("type", "value")]
    [InlineData("Type", "Value")]
    public async Task InvokeArguments_PreserveNamesAndScalarKinds(string typeName, string valueName)
    {
        var arguments = $$"""
            [{"{{typeName}}":"number","{{valueName}}":2.5},
             {"{{typeName}}":"boolean","{{valueName}}":true},
             {"{{typeName}}":"string","{{valueName}}":"alpha\nbeta"},
             {"{{typeName}}":"null","{{valueName}}":null},
             {"{{typeName}}":"missing"}]
            """;
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
            ["vba", "invoke", "--session", "owned", "--procedure-name", "Probe.Check", "--arguments", arguments],
            request =>
            {
                captured = request;
                return new ServiceResponse { Success = true, Result = """{"success":true,"returnKind":"number","value":1}""" };
            });

        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("vba.invoke", captured.Command);
        Assert.Equal("owned", captured.SessionId);
        using var json = JsonDocument.Parse(captured.Args!);
        var root = json.RootElement;
        Assert.Equal("Probe.Check", root.GetProperty("procedureName").GetString());
        var values = root.GetProperty("arguments").EnumerateArray().ToArray();
        Assert.Equal(5, values.Length);
        Assert.Equal("number", values[0].GetProperty("type").GetString());
        Assert.Equal(JsonValueKind.Number, values[0].GetProperty("value").ValueKind);
        Assert.Equal(2.5, values[0].GetProperty("value").GetDouble());
        Assert.Equal("boolean", values[1].GetProperty("type").GetString());
        Assert.True(values[1].GetProperty("value").GetBoolean());
        Assert.Equal("string", values[2].GetProperty("type").GetString());
        Assert.Equal("alpha\nbeta", values[2].GetProperty("value").GetString());
        Assert.Equal("null", values[3].GetProperty("type").GetString());
        Assert.True(!values[3].TryGetProperty("value", out var nullValue) || nullValue.ValueKind == JsonValueKind.Null);
        Assert.Equal("missing", values[4].GetProperty("type").GetString());
        Assert.True(!values[4].TryGetProperty("value", out var missingValue) || missingValue.ValueKind == JsonValueKind.Null);
    }
}
