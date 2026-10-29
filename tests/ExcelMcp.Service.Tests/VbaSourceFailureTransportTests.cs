using System.Reflection;
using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "VBA")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class VbaSourceFailureTransportTests
{
    [Theory]
    [InlineData("direct")]
    [InlineData("wrapped")]
    [InlineData("aggregate")]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2201",
        Justification = "A synthetic fatal COM HRESULT exercises disconnected-response transport without starting or terminating Excel.")]
    public void FatalDisconnect_PreservesObservedVbaStateAndOriginalComCause(string wrapping)
    {
        var sourceError = new VbaSourceEditException("Source insertion failed.",
            new System.Runtime.InteropServices.COMException("original disconnect", unchecked((int)0x80010108)),
            new VbaSourceEditState { ModuleName = "Target", Phase = "inserting", ReadBackError = "Excel unavailable" });
        Exception error = wrapping switch
        {
            "wrapped" => new InvalidOperationException("wrapper", sourceError),
            "aggregate" => new AggregateException(new InvalidOperationException("independent cleanup failure"), sourceError),
            _ => sourceError
        };
        var create = typeof(ExcelMcpService).GetMethod("CreateExcelDisconnectedResponse", BindingFlags.Static | BindingFlags.NonPublic)!;
        var response = Assert.IsType<ServiceResponse>(create.Invoke(null, ["lost-session", error, "Excel process died."]));
        Assert.False(response.Success);
        Assert.Equal("ExcelProcessDied", response.ErrorCategory);
        Assert.Equal("0x80010108", response.HResult);
        Assert.Equal("lost-session", response.SessionId);
        Assert.Contains("original disconnect", response.InnerError);
        if (wrapping == "aggregate") Assert.Contains("independent cleanup failure", response.InnerError);
        Assert.NotNull(response.PartialState);
        using var state = JsonDocument.Parse(response.PartialState);
        Assert.Equal("Target", state.RootElement.GetProperty("moduleName").GetString());
        Assert.Equal("inserting", state.RootElement.GetProperty("phase").GetString());
        Assert.Equal("Excel unavailable", state.RootElement.GetProperty("readBackError").GetString());
    }

    [Theory]
    [InlineData("direct")]
    [InlineData("wrapped")]
    [InlineData("aggregate")]
    public void SourceEditFailure_PreservesStructuredObservedStateAndBothFailures(string wrapping)
    {
        using var service = new ExcelMcpService();
        var sourceError = new VbaSourceEditException("The edit failed; no rollback was attempted.",
            new InvalidOperationException("original insertion failure"),
            new VbaSourceEditState
            {
                ModuleName = "Target",
                Phase = "inserting",
                SourceHash = new string('A', 64),
                LineCount = 2,
                ReadBackError = "independent read-back failure"
            });
        Exception error = wrapping switch
        {
            "wrapped" => new InvalidOperationException("operation wrapper", sourceError),
            "aggregate" => new AggregateException(new InvalidOperationException("visibility snapshot failure"), sourceError),
            _ => sourceError
        };
        var create = typeof(ExcelMcpService).GetMethod("CreateErrorResponse", BindingFlags.Instance | BindingFlags.NonPublic)!;
        var response = Assert.IsType<ServiceResponse>(create.Invoke(service, [error, "vba.replace-module", null, null, null]));
        var roundTrip = JsonSerializer.Deserialize<ServiceResponse>(JsonSerializer.Serialize(response, ServiceProtocol.JsonOptions), ServiceProtocol.JsonOptions)!;
        Assert.False(roundTrip.Success);
        Assert.Equal("vba.replace-module", roundTrip.Command);
        Assert.Contains("original insertion failure", roundTrip.InnerError);
        using var state = JsonDocument.Parse(roundTrip.PartialState!);
        Assert.Equal("Target", state.RootElement.GetProperty("moduleName").GetString());
        Assert.Equal("inserting", state.RootElement.GetProperty("phase").GetString());
        Assert.Equal(new string('A', 64), state.RootElement.GetProperty("sourceHash").GetString());
        Assert.Equal(2, state.RootElement.GetProperty("lineCount").GetInt32());
        Assert.Equal("independent read-back failure", state.RootElement.GetProperty("readBackError").GetString());
    }
}
