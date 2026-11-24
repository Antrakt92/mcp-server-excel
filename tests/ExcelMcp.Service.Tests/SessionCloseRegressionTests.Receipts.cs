using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
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
    public async Task SessionClose_AfterQuarantineCleanup_DoesNotSynthesizeSuccess(bool passiveStatus)
    {
        using var service = new ExcelMcpService();
        var identity = new ExcelProcessIdentity(42, 1234);
        var batch = new FakeBatch
        {
            WorkbookPath = CreateFakeWorkbookPath(),
            DisposeException = new IOException("Fixture teardown did not complete.")
        };
        const string sessionId = "recovered-close-status";
        RegisterSession(service, sessionId, batch, addKnownSessionId: true);
        Assert.False((await CloseSessionAsync(service, sessionId)).Success);

        batch.IsAlive = false;
        batch.TeardownConfirmed = true;
        batch.ProcessObservation = new(identity, ExcelNativeProcessState.Exited, 0, null);
        if (passiveStatus)
        {
            var receipts = new List<ExcelNativeProcessObservation>();
            Assert.Empty(service.SessionManager.GetActiveSessionStatus(receipts.Add));
            Assert.Equal(batch.ProcessObservation, Assert.Single(receipts));
        }
        else
        {
            Assert.Null(service.SessionManager.GetSession(sessionId));
            Assert.False(service.SessionManager.TryGetFilePath(sessionId, out _));
        }

        foreach (var save in new[] { false, true })
        {
            var retry = await CloseSessionAsync(service, sessionId, save);
            Assert.False(retry.Success);
            Assert.Equal("SessionUnavailable", retry.ErrorCategory);
            using var state = JsonDocument.Parse(retry.PartialState!);
            Assert.Equal("unknown", state.RootElement.GetProperty("outcome").GetString());
        }
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
            var closeOnly = await CloseSessionAsync(service, sessionId);
            Assert.False(closeOnly.Success);
            Assert.Equal("SessionUnavailable", closeOnly.ErrorCategory);
            using var state = JsonDocument.Parse(closeOnly.PartialState!);
            Assert.Equal("unknown", state.RootElement.GetProperty("outcome").GetString());
        }
        finally
        {
            batch.DisposeException = null;
        }
    }
}
