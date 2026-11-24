using System.IO.Compression;
using System.IO.Pipelines;
using System.Text.Json;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.McpServer.ServiceBridge;
using Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Collection("ProgramTransport")]
[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "ProgramTransport")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class ProgramTestTransportLifecycleTests(ITestOutputHelper output)
{
    [Fact]
    public async Task OperationResult_InvalidAndMissingReceiptsStayDistinct()
    {
        using var cancellation = new CancellationTokenSource();
        var input = new Pipe();
        var outputPipe = new Pipe();
        var host = await ProgramTransportTestHost.StartAsync(input, outputPipe, cancellation.Token, "OperationResultClient");
        try
        {
            var operationId = Guid.NewGuid();
            var args = new Dictionary<string, object?>
            {
                ["action"] = "operation-result",
                ["workbook_session_id"] = "unseen-session",
                ["operation_id"] = operationId.ToString()
            };
            var knownFormat = await host.Client.CallToolAsync("file_read", args);
            Assert.False(knownFormat.IsError is true);
            using var unknown = JsonDocument.Parse(Assert.Single(knownFormat.Content.OfType<TextContentBlock>()).Text);
            Assert.Equal(operationId, unknown.RootElement.GetProperty("operationId").GetGuid());
            Assert.Equal("unknown", unknown.RootElement.GetProperty("state").GetString());

            args["operation_id"] = "not-a-uuid";
            var malformedId = await host.Client.CallToolAsync("file_read", args);
            Assert.True(malformedId.IsError);
            using var malformed = JsonDocument.Parse(Assert.Single(malformedId.Content.OfType<TextContentBlock>()).Text);
            Assert.Equal("InvalidInput", malformed.RootElement.GetProperty("errorCategory").GetString());
        }
        finally
        {
            await ProgramTransportTestHost.StopAsync(host.Client, input, outputPipe, host.ServerTask, output, cancellation);
        }
    }

    [Fact]
    public async Task ComparePackages_ReturnsExactMismatchWithoutOpeningWorkbookSession()
    {
        using var cancellation = new CancellationTokenSource();
        var input = new Pipe();
        var outputPipe = new Pipe();
        var host = await ProgramTransportTestHost.StartAsync(input, outputPipe, cancellation.Token, "PackageComparisonClient");
        var directory = Directory.CreateTempSubdirectory("ExcelMcpPackageComparison-");
        try
        {
            var before = Path.Join(directory.FullName, "before.xlsx");
            var after = Path.Join(directory.FullName, "after.xlsx");
            CreateComparisonPackage(before, "<workbook/>");
            CreateComparisonPackage(after, "<workbook><sheets/></workbook>");

            string[] allowedParts = ["xl/vbaProject.bin"];
            var response = await host.Client.CallToolAsync("file_read", new Dictionary<string, object?>
            {
                ["action"] = "compare-packages",
                ["before_path"] = before,
                ["after_path"] = after,
                ["allowed_parts"] = allowedParts
            }, cancellationToken: cancellation.Token);

            Assert.False(response.IsError, response.StructuredContent?.GetRawText() ??
                string.Join(" ", response.Content.OfType<TextContentBlock>().Select(block => block.Text)));
            var payload = response.StructuredContent!.Value;
            Assert.True(payload.GetProperty("success").GetBoolean());
            Assert.False(payload.GetProperty("matchesAllowlist").GetBoolean());
            var change = Assert.Single(payload.GetProperty("changes").EnumerateArray());
            Assert.Equal("xl/workbook.xml", change.GetProperty("partName").GetString());
            Assert.Equal("changed", change.GetProperty("changeKind").GetString());
            Assert.NotEqual(change.GetProperty("beforeSha256").GetString(), change.GetProperty("afterSha256").GetString());
            Assert.False(change.GetProperty("allowed").GetBoolean());
            Assert.False(payload.TryGetProperty("structure", out _));

            var structural = Path.Join(directory.FullName, "structural.xlsx");
            CreateComparisonPackage(structural,
                """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>""");
            var structuralResponse = await host.Client.CallToolAsync("file_read", new Dictionary<string, object?>
            {
                ["action"] = "compare-packages",
                ["before_path"] = structural,
                ["after_path"] = structural,
                ["include_structure"] = true
            }, cancellationToken: cancellation.Token);
            Assert.False(structuralResponse.IsError, structuralResponse.StructuredContent?.GetRawText());
            var structure = structuralResponse.StructuredContent!.Value.GetProperty("structure");
            Assert.Equal("xl/workbook.xml", structure.GetProperty("before").GetProperty("workbookPart").GetString());
            Assert.Empty(structure.GetProperty("after").GetProperty("sheets").EnumerateArray());
            Assert.Single(structure.GetProperty("after").GetProperty("relationships").EnumerateArray());
            Assert.Empty(structure.GetProperty("parentStyleMappings").EnumerateArray());

            var malformedMetadata = Path.Join(directory.FullName, "malformed-metadata.xlsx");
            CreateComparisonPackage(malformedMetadata,
                """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets></workbook>""");
            var malformedResponse = await host.Client.CallToolAsync("file_read", new Dictionary<string, object?>
            {
                ["action"] = "compare-packages",
                ["before_path"] = structural,
                ["after_path"] = malformedMetadata,
                ["include_structure"] = true
            }, cancellationToken: cancellation.Token);
            Assert.True(malformedResponse.IsError);
            Assert.Equal("InvalidInput", malformedResponse.StructuredContent!.Value.GetProperty("errorCategory").GetString());

            var list = await host.Client.CallToolAsync("file_read", new Dictionary<string, object?> { ["action"] = "list" },
                cancellationToken: cancellation.Token);
            Assert.False(list.IsError);
            Assert.Empty(list.StructuredContent!.Value.GetProperty("sessions").EnumerateArray());

            var batch = await host.Client.CallToolAsync("batch", new Dictionary<string, object?>
            {
                ["action"] = "run",
                ["workbook_session_id"] = "unused",
                ["commands"] = new object[]
                {
                    new { command = "session.compare-packages", args = new { beforePath = before, afterPath = after } }
                }
            }, cancellationToken: cancellation.Token);
            Assert.False(batch.IsError, batch.StructuredContent?.GetRawText());
            var item = Assert.Single(batch.StructuredContent!.Value.GetProperty("results").EnumerateArray());
            Assert.Equal("session.compare-packages", item.GetProperty("command").GetString());

            var invalid = Path.Join(directory.FullName, "invalid.xlsx");
            await File.WriteAllTextAsync(invalid, "not an Excel package", cancellation.Token);
            var failure = await host.Client.CallToolAsync("file_read", new Dictionary<string, object?>
            {
                ["action"] = "compare-packages",
                ["before_path"] = before,
                ["after_path"] = invalid
            }, cancellationToken: cancellation.Token);
            Assert.True(failure.IsError);
            Assert.Equal("InvalidInput", failure.StructuredContent!.Value.GetProperty("errorCategory").GetString());
        }
        finally
        {
            try
            {
                await ProgramTransportTestHost.StopAsync(host.Client, input, outputPipe, host.ServerTask, output, cancellation);
            }
            finally
            {
                directory.Delete(recursive: true);
            }
        }
    }

    private static void CreateComparisonPackage(string path, string workbook)
    {
        using var package = ZipFile.Open(path, ZipArchiveMode.Create);
        var parts = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """,
            ["xl/workbook.xml"] = workbook
        };
        foreach (var part in parts)
        {
            using var writer = new StreamWriter(package.CreateEntry(part.Key).Open());
            writer.Write(part.Value);
        }
    }

    [Theory]
    [InlineData("exception", "list")]
    [InlineData("invalid-json", "list")]
    [InlineData("missing-result", "list")]
    [InlineData("missing-result", "operation-result")]
    public async Task BackendFailure_UsesToolErrorWithoutPrivateDetails(string failureMode, string action)
    {
        using var cancellation = new CancellationTokenSource();
        var input = new Pipe();
        var outputPipe = new Pipe();
        var backend = new ListBackend { FailureMode = failureMode };
        var host = await ProgramTransportTestHost.StartAsync(
            input, outputPipe, cancellation.Token, "ErrorPrivacyClient", () => backend);
        try
        {
            var args = new Dictionary<string, object?>
            {
                ["action"] = action
            };
            if (action == "operation-result")
            {
                args["workbook_session_id"] = "owned-session";
                args["operation_id"] = Guid.NewGuid().ToString();
            }
            var result = await host.Client.CallToolAsync("file_read", args);
            Assert.True(result.IsError);
            var text = Assert.Single(result.Content.OfType<TextContentBlock>()).Text;
            Assert.DoesNotContain("synthetic-private-detail", text, StringComparison.Ordinal);
            if (failureMode == "missing-result")
            {
                var payload = Assert.IsType<JsonElement>(result.StructuredContent);
                using var parsedText = JsonDocument.Parse(text);
                Assert.True(JsonElement.DeepEquals(parsedText.RootElement, payload));

                using var expected = JsonDocument.Parse($$"""
                    {
                        "success": false,
                        "errorMessage": "Read operation 'session.{{action}}' returned no result.",
                        "errorCategory": "InvalidResponse",
                        "command": "session.{{action}}"
                    }
                    """);
                Assert.True(JsonElement.DeepEquals(expected.RootElement, payload), payload.GetRawText());
            }
            else
            {
                Assert.Equal("An error occurred invoking 'file_read'.", text);
                Assert.Null(result.StructuredContent);
            }
        }
        finally
        {
            await ProgramTransportTestHost.StopAsync(host.Client, input, outputPipe, host.ServerTask, output, cancellation);
        }
    }

    [Fact]
    public async Task IndependentHosts_ShutdownDoesNotResetAnotherHost()
    {
        using var firstCancellation = new CancellationTokenSource();
        using var secondCancellation = new CancellationTokenSource();
        var firstInput = new Pipe();
        var firstOutput = new Pipe();
        var secondInput = new Pipe();
        var secondOutput = new Pipe();
        var firstBackend = new ListBackend();
        var secondBackend = new ListBackend();
        var first = await ProgramTransportTestHost.StartAsync(
            firstInput, firstOutput, firstCancellation.Token, "FirstHost", () => firstBackend);
        var second = await ProgramTransportTestHost.StartAsync(
            secondInput, secondOutput, secondCancellation.Token, "SecondHost", () => secondBackend);
        var firstStopped = false;
        try
        {
            var args = new Dictionary<string, object?> { ["action"] = "list" };
            Assert.False((await first.Client.CallToolAsync("file_read", args)).IsError is true);
            Assert.False((await second.Client.CallToolAsync("file_read", args)).IsError is true);
            await ProgramTransportTestHost.StopAsync(first.Client, firstInput, firstOutput, first.ServerTask, output, firstCancellation);
            firstStopped = true;
            Assert.True(firstBackend.Disposed);
            Assert.False(secondBackend.Disposed);
            Assert.False((await second.Client.CallToolAsync("file_read", args)).IsError is true);
        }
        finally
        {
            if (!firstStopped)
                await ProgramTransportTestHost.StopAsync(first.Client, firstInput, firstOutput, first.ServerTask, output, firstCancellation);
            await ProgramTransportTestHost.StopAsync(second.Client, secondInput, secondOutput, second.ServerTask, output, secondCancellation);
        }
        Assert.True(secondBackend.Disposed);
    }

    private sealed class ListBackend : IServiceBridgeBackend
    {
        internal bool Disposed { get; private set; }
        internal string? FailureMode { get; init; }
        public Task<ServiceResponse> ProcessAsync(ServiceRequest request)
        {
            ObjectDisposedException.ThrowIf(Disposed, this);
            if (FailureMode == "exception")
                throw new InvalidOperationException("synthetic-private-detail");
            if (FailureMode is "invalid-json" or "missing-result")
                return Task.FromResult(new ServiceResponse { Success = true, Result = FailureMode == "invalid-json" ? "synthetic-private-detail" : null });
            return Task.FromResult(new ServiceResponse { Success = true, Result = """{"success":true,"sessions":[],"count":0}""" });
        }
        public void RevealSession(string sessionId) => throw new InvalidOperationException("Unexpected reveal.");
        public void Dispose() => Disposed = true;
    }
}
