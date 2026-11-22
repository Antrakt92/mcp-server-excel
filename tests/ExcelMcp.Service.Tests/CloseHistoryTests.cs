using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "ServiceDaemon")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class CloseHistoryTests
{
    [Fact]
    public void CombinedDispatchAndRecordingFailureKeepsPrimaryAndBothRecordingStages()
    {
        var history = new CloseHistory();
        var target = new CloseHistory.Target("same-session", new ExcelProcessIdentity(42, 1234));
        var request = new ServiceRequest { Command = "session.close", SessionId = target.SessionId };
        var original = new ServiceResponse
        {
            ErrorCategory = "Busy",
            ExceptionType = "ExcelBusyException",
            HResult = "0x800706B5",
            PartialState = """{"outcome":"unknown"}"""
        };
        var failure = original.WithCompletionFailure(request, new IOException("worker publication"))
            .WithCompletionFailure(request, new InvalidOperationException("controller publication"));
        history.Record(target, failure, new JsonObject { ["stage"] = "snapshotRefresh", ["hresult"] = "0x800706BE" });

        var entry = Assert.IsType<JsonObject>(Assert.Single(history.Snapshot([], _ => false)));
        Assert.False(entry["closeConfirmed"]!.GetValue<bool>());
        Assert.False(entry["saveConfirmed"]!.GetValue<bool>());
        Assert.Equal("Busy", entry["lastFailure"]!["category"]!.GetValue<string>());
        Assert.Equal("0x800706B5", entry["lastFailure"]!["hresult"]!.GetValue<string>());
        Assert.Equal("0x800706BE", entry["priorStatusDiagnostic"]!["hresult"]!.GetValue<string>());
        var recording = Assert.IsType<JsonArray>(entry["recordingFailures"]);
        Assert.Equal(2, recording.Count);
        Assert.Equal("InvalidOperationException", recording[0]!["exceptionType"]!.GetValue<string>());
        Assert.Equal("IOException", recording[1]!["exceptionType"]!.GetValue<string>());
    }

    [Fact]
    public void SaveAndCloseConfirmationSurvivesARecordingFailureWithoutAttestingRelease()
    {
        var history = new CloseHistory();
        var target = new CloseHistory.Target("same-session", new ExcelProcessIdentity(42, 1234));
        var request = new ServiceRequest { Command = "session.close", SessionId = target.SessionId };
        var original = new ServiceResponse { Success = true, Result = """{"saved":true}""" };
        history.Record(target, original.WithCompletionFailure(request, new IOException("recording failed")));

        var entry = Assert.IsType<JsonObject>(Assert.Single(history.Snapshot([], _ => false)));
        Assert.True(entry["closeConfirmed"]!.GetValue<bool>());
        Assert.True(entry["saveConfirmed"]!.GetValue<bool>());
        Assert.False(entry["sessionReleased"]!.GetValue<bool>());
        Assert.Null(entry["lastFailure"]);
        Assert.Equal("IOException", entry["recordingFailures"]![0]!["exceptionType"]!.GetValue<string>());
    }

    [Fact]
    public void FailureAndLateExitDoNotConfirmSaveOrRelease()
    {
        var history = new CloseHistory();
        var identity = new ExcelProcessIdentity(42, 1234);
        var target = new CloseHistory.Target("close-session", identity);
        history.Record(target, new ServiceResponse
        {
            ErrorCategory = "Busy",
            ExceptionType = "ExcelBusyException",
            ErrorMessage = "Private workbook path",
            HResult = "0x800706B5"
        });

        var unconfirmed = Assert.IsType<JsonObject>(Assert.Single(history.Snapshot(
            [new(new ExcelProcessIdentity(42, 1235), ExcelNativeProcessState.Exited, 0, null)], _ => false)));
        Assert.Null(unconfirmed["nativeExit"]);
        Assert.False(unconfirmed["sessionReleased"]!.GetValue<bool>());

        var exit = new ExcelNativeProcessObservation(identity, ExcelNativeProcessState.Exited, 0, null);
        var settled = Assert.IsType<JsonObject>(Assert.Single(history.Snapshot([exit], session => session == target.SessionId)));
        Assert.Equal(target.SessionId, settled["sessionId"]!.GetValue<string>());
        Assert.Equal("1234", settled["startedAtUtcFileTime"]!.GetValue<string>());
        Assert.Equal(0u, settled["nativeExit"]!["exitCode"]!.GetValue<uint>());
        Assert.True(settled["sessionReleased"]!.GetValue<bool>());
        Assert.False(settled["closeConfirmed"]!.GetValue<bool>());
        Assert.False(settled["saveConfirmed"]!.GetValue<bool>());
        Assert.Equal("Busy", settled["lastFailure"]!["category"]!.GetValue<string>());
        Assert.Equal("0x800706B5", settled["lastFailure"]!["hresult"]!.GetValue<string>());
        Assert.DoesNotContain("Private workbook path", settled.ToJsonString());
    }

    [Fact]
    public void LaterSuccessPreservesFailureAndRequiresAnExplicitSaveReceipt()
    {
        var history = new CloseHistory();
        var target = new CloseHistory.Target("same-session", new ExcelProcessIdentity(42, 1234));
        history.Record(target, new ServiceResponse { ErrorCategory = "ExcelDisconnected" });
        history.Record(target, new ServiceResponse { Success = true });
        var closed = Assert.IsType<JsonObject>(Assert.Single(history.Snapshot([], _ => true)));
        Assert.True(closed["closeConfirmed"]!.GetValue<bool>());
        Assert.False(closed["saveConfirmed"]!.GetValue<bool>());
        Assert.Equal("ExcelDisconnected", closed["lastFailure"]!["category"]!.GetValue<string>());

        history.Record(target, new ServiceResponse { Success = true, Result = """{"saved":true}""" });
        var saved = Assert.IsType<JsonObject>(Assert.Single(history.Snapshot([], _ => true)));
        Assert.True(saved["saveConfirmed"]!.GetValue<bool>());
        Assert.Equal("ExcelDisconnected", saved["lastFailure"]!["category"]!.GetValue<string>());
    }

    [Fact]
    public void HistoryIsBoundedAndDoesNotMixWorkbookProcessGenerations()
    {
        var history = new CloseHistory();
        for (int index = 0; index < 18; ++index)
        {
            var target = new CloseHistory.Target("same-session", new ExcelProcessIdentity(42, 1234 + index));
            history.Record(target, new ServiceResponse { Success = index == 17 });
        }

        var entries = history.Snapshot([], _ => false);
        Assert.Equal(16, entries.Count);
        Assert.Equal("1236", entries[0]!["startedAtUtcFileTime"]!.GetValue<string>());
        Assert.False(entries[14]!["closeConfirmed"]!.GetValue<bool>());
        Assert.True(entries[15]!["closeConfirmed"]!.GetValue<bool>());
    }
}
