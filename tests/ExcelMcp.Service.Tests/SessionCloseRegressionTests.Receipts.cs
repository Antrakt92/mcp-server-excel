using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class SessionCloseRegressionTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SessionClose_OnlyCompletedSaveCanBeAcknowledgedByReplay(bool saveFirst)
    {
        using var service = new ExcelMcpService();
        var saves = 0;
        var batch = new FakeBatch { WorkbookPath = CreateFakeWorkbookPath(), BeforeSave = () => saves++ };
        const string sessionId = "closed-session-save-replay";
        RegisterSession(service, sessionId, batch, addKnownSessionId: true);

        var original = await CloseSessionAsync(service, sessionId, saveFirst);
        Assert.True(original.Success, original.ErrorMessage);
        if (saveFirst)
        {
            using var receipt = JsonDocument.Parse(original.Result!);
            Assert.True(receipt.RootElement.GetProperty("saved").GetBoolean());
        }
        var replay = await CloseSessionAsync(service, sessionId, save: true);
        Assert.Equal(saveFirst, replay.Success);
        if (saveFirst)
        {
            using var receipt = JsonDocument.Parse(replay.Result!);
            Assert.True(receipt.RootElement.GetProperty("saved").GetBoolean());
        }
        else
            Assert.Equal("SessionUnavailable", replay.ErrorCategory);
        Assert.Equal(saveFirst ? 1 : 0, saves);
        Assert.True((await CloseSessionAsync(service, sessionId)).Success);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SessionClose_RecoveredTeardownCannotAcknowledgeARequestedSave(bool saveBeforeFailure)
    {
        using var service = new ExcelMcpService();
        var saves = 0;
        var batch = new FakeBatch
        {
            WorkbookPath = CreateFakeWorkbookPath(),
            BeforeSave = () => saves++,
            DisposeException = new IOException("Fixture teardown did not complete.")
        };
        const string sessionId = "recovered-close-without-save";
        RegisterSession(service, sessionId, batch, addKnownSessionId: true);

        try
        {
            var failed = await CloseSessionAsync(service, sessionId, saveBeforeFailure);
            Assert.False(failed.Success);
            batch.IsAlive = false;
            batch.TeardownConfirmed = true;
            var retry = await CloseSessionAsync(service, sessionId, save: true);
            Assert.False(retry.Success);
            Assert.Equal("SessionUnavailable", retry.ErrorCategory);
            Assert.Equal(saveBeforeFailure ? 1 : 0, saves);
            Assert.False((await CloseSessionAsync(service, sessionId, save: true)).Success);
            Assert.True((await CloseSessionAsync(service, sessionId)).Success);
        }
        finally
        {
            batch.DisposeException = null;
        }
    }
}
