using System.Diagnostics;
using System.Text;
using System.Text.Json;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "ToolResponse")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class ExcelToolsBaseResponseParityTests
{
    [Theory]
    [InlineData("{\"success\":true,\"values\":[[\"=literal\",3,null]]}", "{\"success\":true,\"values\":[[\"=literal\",3,null]]}", false)]
    [InlineData("[\"x\",2,null]", "{\"result\":[\"x\",2,null]}", false)]
    [InlineData("42", "{\"result\":42}", false)]
    [InlineData("\"hello\"", "{\"result\":\"hello\"}", false)]
    [InlineData("null", "{\"result\":null}", false)]
    [InlineData("{\"success\":false,\"errorCategory\":\"Timeout\"}", "{\"success\":false,\"errorCategory\":\"Timeout\"}", true)]
    [InlineData("{\"isError\":true,\"errorMessage\":\"failed\"}", "{\"isError\":true,\"errorMessage\":\"failed\"}", true)]
    public async Task PublicEntryPreservesTextStructureAndErrorFlag(string json, string expected, bool isError)
    {
        var result = await ExcelToolsBase.ExecuteToolActionAsync("range_read", "get-values",
            () => Task.FromResult(json), CancellationToken.None);
        await Task.Yield(); // Returned structured content must outlive the parsing document.
        Assert.Equal(isError, result.IsError);
        Assert.Equal(json, Assert.IsType<TextContentBlock>(Assert.Single(result.Content)).Text);
        Assert.NotNull(result.StructuredContent);
        using var expectedDocument = JsonDocument.Parse(expected);
        Assert.True(JsonElement.DeepEquals(expectedDocument.RootElement, result.StructuredContent.Value));
    }

    [Fact]
    public async Task DiagnosticNegativeRemainsExpectedAndPreservesOriginalPayload()
    {
        const string json = "{\"success\":false,\"reason\":\"not an Excel workbook\"}";
        var result = await ExcelToolsBase.ExecuteToolActionAsync("file_read", "test",
            () => Task.FromResult(json), CancellationToken.None);
        Assert.False(result.IsError);
        Assert.Equal(json, Assert.IsType<TextContentBlock>(Assert.Single(result.Content)).Text);
        Assert.False(result.StructuredContent!.Value.GetProperty("success").GetBoolean());
    }

    [Fact]
    public async Task CustomFactoryReceivesExactJsonAndRetainsItsOwnResult()
    {
        const string json = "{ \"success\": true, \"x\": 7 }";
        var calls = 0;
        var custom = new CallToolResult { IsError = true, Content = [new TextContentBlock { Text = "custom-rendered" }] };
        var result = await ExcelToolsBase.ExecuteToolActionAsync("screenshot", "range",
            () => Task.FromResult(json), CancellationToken.None, received =>
            {
                Assert.Equal(json, received);
                calls++;
                return custom;
            });
        Assert.Same(custom, result);
        Assert.Equal(1, calls);
        Assert.True(result.IsError);
    }

    [Theory]
    [InlineData("range_read", "get-values", true)]
    [InlineData("file_read", "test", false)]
    public async Task NullCustomFactoryResultRetainsDefaultFallbackClassification(string tool, string action, bool isError)
    {
        const string json = "{\"success\":false,\"reason\":\"expected fixture\"}";
        var result = await ExcelToolsBase.ExecuteToolActionAsync(tool, action,
            () => Task.FromResult(json), CancellationToken.None, _ => null!);
        Assert.Equal(isError, result.IsError);
        Assert.Equal(json, Assert.IsType<TextContentBlock>(Assert.Single(result.Content)).Text);
        Assert.False(result.StructuredContent!.Value.GetProperty("success").GetBoolean());
    }

    [Fact]
    public async Task InvalidJsonStillThrowsBeforeCustomFactory()
    {
        var calls = 0;
        await Assert.ThrowsAnyAsync<JsonException>(() => ExcelToolsBase.ExecuteToolActionAsync("range_read", "get-values",
            () => Task.FromResult("not-json"), CancellationToken.None, _ =>
            {
                calls++;
                return new CallToolResult { Content = [] };
            }));
        Assert.Equal(0, calls);
    }

    [Fact]
    public async Task ArgumentFailureStillReturnsStructuredInputError()
    {
        var result = await ExcelToolsBase.ExecuteToolActionAsync("range", "set-values",
            () => throw new ArgumentException("Invalid range fixture."), CancellationToken.None);
        Assert.True(result.IsError);
        Assert.NotNull(result.StructuredContent);
        Assert.Equal("InvalidInput", result.StructuredContent.Value.GetProperty("errorCategory").GetString());
        Assert.Contains("Invalid range fixture.", result.StructuredContent.Value.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task CancellationAfterOperationWinsOverInvalidJsonAndSkipsFactory()
    {
        using var cancel = new CancellationTokenSource();
        var calls = 0;
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ExcelToolsBase.ExecuteToolActionAsync("range_read", "get-values",
            () => { cancel.Cancel(); return Task.FromResult("not-json"); }, cancel.Token, _ =>
            {
                calls++;
                return new CallToolResult { Content = [] };
            }));
        Assert.Equal(0, calls);
    }
}

[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "ToolResponse")]
[Trait("Speed", "Medium")]
[Trait("RunType", "OnDemand")]
[Trait("RequiresExcel", "false")]
[Collection("Sequential")]
public sealed class ExcelToolsBaseResponseBenchmarks(ITestOutputHelper output)
{
    [Fact]
    public async Task PublicEntryResponseBenchmark()
    {
        foreach (var size in new[] { 1024, 1024 * 1024 })
        {
            var json = JsonSerializer.Serialize(new { success = true, values = new[] { new[] { new string('x', size) } } });
            Func<Task<string>> operation = () => Task.FromResult(json);
            for (var warm = 0; warm < 8; warm++)
                await ExcelToolsBase.ExecuteToolActionAsync("range_read", "get-values", operation, CancellationToken.None);
            var iterations = size < 2048 ? 1000 : 20;
            for (var round = 1; round <= 3; round++)
            {
                var thread = Environment.CurrentManagedThreadId;
                var before = GC.GetAllocatedBytesForCurrentThread();
                var stopwatch = Stopwatch.StartNew();
                CallToolResult? result = null;
                for (var index = 0; index < iterations; index++)
                    result = await ExcelToolsBase.ExecuteToolActionAsync("range_read", "get-values", operation, CancellationToken.None);
                stopwatch.Stop();
                var allocated = GC.GetAllocatedBytesForCurrentThread() - before;
                Assert.Equal(thread, Environment.CurrentManagedThreadId);
                Assert.NotNull(result);
                Assert.NotNull(result.StructuredContent);
                Assert.Equal(size, result.StructuredContent.Value.GetProperty("values")[0][0].GetString()!.Length);
                output.WriteLine(JsonSerializer.Serialize(new
                {
                    benchmark = "public-entry-response",
                    payloadBytes = Encoding.UTF8.GetByteCount(json),
                    iterations,
                    round,
                    elapsedMs = stopwatch.Elapsed.TotalMilliseconds,
                    allocatedBytes = allocated,
                    runtime = Environment.Version.ToString()
                }));
            }
        }
    }
}
