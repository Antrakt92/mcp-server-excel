using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service.Rpc;
using Sbroenne.ExcelMcp.Service.Workers;
using StreamJsonRpc;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ExcelControllerTests
{
    [Theory(Timeout = 15000)]
    [InlineData(false, false, true, null)]
    [InlineData(false, false, false, null)]
    [InlineData(true, false, true, null)]
    [InlineData(false, true, true, null)]
    [InlineData(false, false, true, "workbookLocationChanged")]
    [InlineData(false, false, true, "Excel is busy")]
    public async Task Controller_WorkerStatusFailureSeparatesUnknownReadinessFromBlockedClose(
        bool pending, bool releasing, bool matchingObservation, string? lastBlocker)
    {
        const string sessionId = "readiness-to-confirm";
        var directory = Directory.CreateTempSubdirectory("excelmcp-close-readiness-");
        var registry = new WorkerRegistry(directory.FullName);
        try
        {
            using var controller = new ExcelController(launcher: Launcher, registry: registry);
            await StopNativeObservationPollingAsync(controller);
            var generation = Assert.IsType<string>(typeof(ExcelController)
                .GetField("_generation", NativeObservationHidden)!.GetValue(controller));
            var current = WorkerProcess.Current;
            var native = new ExcelProcessIdentity(123, 134_381_234_567_890_123);
            var observed = matchingObservation ? native : native with { StartedAtUtcFileTime = 456 };
            var registration = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", current,
                current, generation, WorkerRegistry.Build, native, OwnsExcel: true,
                LastExcelObservation: new(observed, ExcelNativeProcessState.Alive, null, null));
            var empty = new ServiceResponse { Success = true, Result = """{"success":true,"sessions":[],"count":0}""" };
            var released = new WorkerSnapshot(registration with { ReleaseConfirmed = true }, [], empty,
                null, null, new Dictionary<string, IReadOnlyList<string>>(), Revision: 2);
            var target = new CloseReceiptWorker(new ServiceResponse { Success = true, Result = """{"saved":false}""" }, released);
            using var host = new DaemonHost(registration.PipeName, pipe => new JsonRpc(pipe, pipe, target),
                maximumConnections: 8);
            var run = host.RunAsync(registration.PipeName);
            var workers = Assert.IsType<ConcurrentDictionary<string, ExcelWorkerClient>>(
                typeof(ExcelController).GetField("_workers", NativeObservationHidden)!.GetValue(controller));
            try
            {
                var worker = new ExcelWorkerClient(registration, current, generation);
                worker.RecordSnapshot(released with
                {
                    Registration = registration,
                    Sessions = [new SessionDescriptor(sessionId, null)],
                    SessionList = new ServiceResponse
                    {
                        Success = true,
                        Result = JsonSerializer.Serialize(new
                        {
                            sessions = new[] { new { sessionId, canClose = lastBlocker is null, operationPending = false, blockingReason = lastBlocker } }
                        }, ServiceProtocol.JsonOptions)
                    },
                    PendingOperation = pending ? "range-write" : null,
                    OwnershipKeys = new Dictionary<string, IReadOnlyList<string>> { [sessionId] = [] },
                    Releasing = releasing,
                    Revision = 1
                });
                Assert.True(workers.TryAdd(registration.PipeName, worker));
                var sessions = Assert.IsType<ConcurrentDictionary<string, ExcelWorkerClient>>(
                    typeof(ExcelController).GetField("_sessions", NativeObservationHidden)!.GetValue(controller));
                Assert.True(sessions.TryAdd(sessionId, worker));
                typeof(ExcelController).GetMethod("ApplyWorkerSnapshot", NativeObservationHidden)!
                    .Invoke(controller, [worker]);
#pragma warning disable CA2201 // Synthetic COM failure exercises controller status reporting without Excel.
                worker.ReportStatusFailure("snapshotTransport", "Worker status failed.",
                    new COMException("The interface is unknown.", unchecked((int)0x800706B5)));
#pragma warning restore CA2201

                var list = await controller.ProcessAsync(new ServiceRequest { Command = "session.list" });
                Assert.True(list.Success, list.ErrorMessage);
                using var document = JsonDocument.Parse(list.Result!);
                var row = Assert.Single(document.RootElement.GetProperty("sessions").EnumerateArray());
                Assert.False(row.GetProperty("canClose").GetBoolean());
                Assert.Equal(pending, row.GetProperty("operationPending").GetBoolean());
                Assert.Equal("WorkerUnavailable", row.GetProperty("blockingReason").GetString());
                Assert.Equal("Worker status failed.", row.GetProperty("blockingMessage").GetString());
                Assert.Equal(pending || releasing || lastBlocker is not null ? "blocked" : "unconfirmed",
                    row.GetProperty("closeReadiness").GetString());
                Assert.Equal(lastBlocker is not null, row.TryGetProperty("lastKnownBlockingReason", out var knownBlocker));
                if (lastBlocker is not null)
                    Assert.Equal(lastBlocker, knownBlocker.GetString());
                var diagnostic = row.GetProperty("statusDiagnostic");
                Assert.Equal("snapshotTransport", diagnostic.GetProperty("stage").GetString());
                Assert.Equal("COMException", diagnostic.GetProperty("exceptionType").GetString());
                Assert.Equal("0x800706B5", diagnostic.GetProperty("hresult").GetString());
                Assert.Equal(matchingObservation, diagnostic.TryGetProperty("nativeProcess", out var process));
                if (matchingObservation)
                {
                    Assert.Equal(123, process.GetProperty("processId").GetInt32());
                    Assert.Equal("134381234567890123", process.GetProperty("startedAtUtcFileTime").GetString());
                    Assert.Equal("Alive", process.GetProperty("state").GetString());
                    Assert.Equal(JsonValueKind.Null, process.GetProperty("exitCode").ValueKind);
                    Assert.Equal(JsonValueKind.Null, process.GetProperty("nativeError").ValueKind);
                }

                if (!pending && !releasing && lastBlocker is null)
                {
                    var close = await controller.ProcessAsync(new ServiceRequest
                    {
                        Command = "session.close",
                        SessionId = sessionId,
                        Args = """{"save":false}"""
                    });
                    Assert.True(close.Success, close.ErrorMessage);
                    Assert.Equal(1, target.Calls);
                    var after = await controller.ProcessAsync(new ServiceRequest { Command = "session.list" });
                    using var settled = JsonDocument.Parse(after.Result!);
                    Assert.Empty(settled.RootElement.GetProperty("sessions").EnumerateArray());
                    var closeHistory = Assert.Single(settled.RootElement.GetProperty("recentCloseAttempts").EnumerateArray());
                    Assert.Equal(sessionId, closeHistory.GetProperty("sessionId").GetString());
                    Assert.True(closeHistory.GetProperty("closeConfirmed").GetBoolean());
                    Assert.True(closeHistory.GetProperty("sessionReleased").GetBoolean());
                    Assert.Equal("0x800706B5", closeHistory.GetProperty("priorStatusDiagnostic").GetProperty("hresult").GetString());
                }
            }
            finally
            {
                workers.TryRemove(registration.PipeName, out _);
                host.RequestShutdown();
                await run.WaitAsync(TimeSpan.FromSeconds(4));
            }
        }
        finally
        {
            directory.Delete(recursive: true);
        }
    }
}
