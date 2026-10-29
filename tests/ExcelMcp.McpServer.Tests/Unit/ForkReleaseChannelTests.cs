using System.Net;
using Sbroenne.ExcelMcp.McpServer.Infrastructure;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Feature", "ReleaseChannel")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ForkReleaseChannelTests
{
    [Theory]
    [InlineData("{\"tag_name\":\"v2.5.4\"}", "2.5.4")]
    [InlineData("{\"tag_name\":\"v2.5.5\"}", "2.5.5")]
    [InlineData("{\"tag_name\":null}", null)]
    [InlineData("{}", null)]
    public async Task LatestVersion_UsesForkReleaseMetadata(string json, string? expected)
    {
        using var handler = new RecordingHandler(HttpStatusCode.OK, json);
        using var client = new HttpClient(handler);
        var version = await NuGetVersionChecker.GetLatestVersionAsync(client);
        Assert.Equal(expected, version);
        Assert.Equal("https://api.github.com/repos/Antrakt92/mcp-server-excel/releases/latest", handler.RequestUrl);
        Assert.Equal("https://github.com/Antrakt92/mcp-server-excel/releases/latest", DistributionInfo.ReleasesUrl);
        Assert.Equal("https://github.com/Antrakt92/mcp-server-excel", DistributionInfo.RepositoryUrl);
    }

    [Fact]
    public async Task UnavailableForkRelease_DoesNotFallBackToUpstream()
    {
        using var handler = new RecordingHandler(HttpStatusCode.NotFound, "{}");
        using var client = new HttpClient(handler);
        Assert.Null(await NuGetVersionChecker.GetLatestVersionAsync(client));
        Assert.Equal(1, handler.RequestCount);
        Assert.Equal("https://api.github.com/repos/Antrakt92/mcp-server-excel/releases/latest", handler.RequestUrl);
    }

    private sealed class RecordingHandler(HttpStatusCode status, string json) : HttpMessageHandler
    {
        internal string? RequestUrl { get; private set; }
        internal int RequestCount { get; private set; }
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            RequestUrl = request.RequestUri?.AbsoluteUri;
            RequestCount++;
            return Task.FromResult(new HttpResponseMessage(status) { Content = new StringContent(json) });
        }
    }
}
