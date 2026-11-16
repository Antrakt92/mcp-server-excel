using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Fact]
    public async Task NamedRange_CreateHidden_PersistsAcrossReopenAndStaysOutOfList()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            const string name = "FixtureMarker";
            const string reference = "Sheet1!$A$1";
            var path = Path.Join(directory, "hidden-name.xlsm");
            var session = await CreateSessionAsync(service, path);
            sessions[session] = 0;
            await WriteMarkerAsync(service, session, "Persisted marker");

            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "namedrange.create",
                SessionId = session,
                Args = JsonSerializer.Serialize(new { name, reference, visible = false }, ServiceProtocol.JsonOptions)
            }));
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "namedrange.create",
                SessionId = session,
                Args = JsonSerializer.Serialize(new { name = "VisibleMarker", reference }, ServiceProtocol.JsonOptions)
            }));

            await CloseSessionAsync(service, session, save: true);
            sessions.TryRemove(session, out _);
            session = await OpenSessionAsync(service, path);
            sessions[session] = 0;

            var list = await service.ProcessAsync(new ServiceRequest
            {
                Command = "namedrange.list",
                SessionId = session
            });
            RequireSuccess(list);
            using var names = JsonDocument.Parse(list.Result!);
            var visible = Assert.Single(names.RootElement.GetProperty("namedRanges").EnumerateArray());
            Assert.Equal("VisibleMarker", visible.GetProperty("name").GetString());

            var metadata = await service.ProcessAsync(new ServiceRequest
            {
                Command = "namedrange.read",
                SessionId = session,
                Args = JsonSerializer.Serialize(new { name, metadataOnly = true }, ServiceProtocol.JsonOptions)
            });
            RequireSuccess(metadata);
            using var result = JsonDocument.Parse(metadata.Result!);
            Assert.Equal(name, result.RootElement.GetProperty("name").GetString());
            Assert.Equal("=" + reference, result.RootElement.GetProperty("refersTo").GetString());
            Assert.False(result.RootElement.GetProperty("visible").GetBoolean());
            Assert.Equal("LocalRange", result.RootElement.GetProperty("resolutionStatus").GetString());
            Assert.Equal(1, result.RootElement.GetProperty("areaCount").GetInt32());
            Assert.Equal("NotRead", result.RootElement.GetProperty("valueType").GetString());
            Assert.False(result.RootElement.TryGetProperty("value", out _));

            var read = await service.ProcessAsync(new ServiceRequest
            {
                Command = "namedrange.read",
                SessionId = session,
                Args = JsonSerializer.Serialize(new { name }, ServiceProtocol.JsonOptions)
            });
            RequireSuccess(read);
            using var value = JsonDocument.Parse(read.Result!);
            Assert.Equal("Persisted marker", value.RootElement.GetProperty("value").GetString());
            Assert.False(value.RootElement.TryGetProperty("visible", out _));
            Assert.False(value.RootElement.TryGetProperty("resolutionStatus", out _));
        });
    }
}
