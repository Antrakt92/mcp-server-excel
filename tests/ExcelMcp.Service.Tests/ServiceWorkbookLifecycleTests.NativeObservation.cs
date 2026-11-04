using System.Collections.Concurrent;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Fact]
    public async Task NativeObservation_AlreadyExitedNormalWorkbookDoesNotReportSuccessfulSave()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var sessionId = await CreateSessionAsync(service, Path.Join(directory, "observed-dead-save.xlsx"));
            sessions[sessionId] = 0;
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(sessionId));
            batch.Close();
            Assert.Equal(ExcelNativeProcessState.Exited,
                Assert.IsType<ExcelNativeProcessObservation>(batch.ProcessObservation).State);

            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = sessionId,
                Args = """{"save":true}"""
            });

            Assert.False(response.Success);
            Assert.Contains(sessionId, service.SessionManager.ActiveSessionIds);
            await CloseSessionAsync(service, sessionId, save: false);
            sessions.TryRemove(sessionId, out _);
        });
    }

    [Fact]
    public async Task NativeObservation_DeadBeforeServiceCloseCannotHidePersistenceFailure()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var sessionId = await CreateSessionAsync(service, Path.Join(directory, "observed-dead-before-close.xlsx"));
            sessions[sessionId] = 0;
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(sessionId));
            var identity = Assert.IsType<ExcelProcessIdentity>(batch.ProcessIdentity);
            void RejectObservation(ExcelNativeProcessObservation observation)
            {
                if (observation.Identity == identity)
                    throw new IOException("Synthetic native observation persistence failure.");
            }
            SessionManager.ExcelProcessObservationRecorded += RejectObservation;
            try
            {
                // Native exit precedes the first Service close; its persistence has not succeeded.
                Assert.Throws<IOException>(() => batch.Close());
                Assert.Equal(ExcelNativeProcessState.Exited,
                    Assert.IsType<ExcelNativeProcessObservation>(batch.ProcessObservation).State);

                var response = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.close",
                    SessionId = sessionId,
                    Args = """{"save":false}"""
                });

                Assert.False(response.Success);
                Assert.Contains("Synthetic native observation persistence failure.",
                    response.ErrorMessage + response.InnerError, StringComparison.Ordinal);
                Assert.Contains(sessionId, service.SessionManager.ActiveSessionIds);
                Assert.Contains(identity, SessionManager.GetTrackedExcelProcesses());
            }
            finally
            {
                SessionManager.ExcelProcessObservationRecorded -= RejectObservation;
                batch.Dispose();
            }
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NativeObservation_CloseRecordsExitBeforeReleaseAndRetainsItAfterDisposal(bool rejectFirstRecord)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var sessionId = await CreateSessionAsync(service, Path.Join(directory, "observed-close.xlsx"));
            sessions[sessionId] = 0;
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(sessionId));
            var identity = Assert.IsType<ExcelProcessIdentity>(batch.ProcessIdentity);
            var observed = new ConcurrentQueue<ExcelNativeProcessObservation>();
            var order = new ConcurrentQueue<string>();
            void RecordObservation(ExcelNativeProcessObservation observation)
            {
                if (observation.Identity != identity)
                    return;
                observed.Enqueue(observation);
                order.Enqueue("observation");
            }
            void RejectObservation(ExcelNativeProcessObservation observation)
            {
                if (observation.Identity == identity)
                    throw new IOException("Synthetic native observation persistence failure.");
            }
            void RecordRelease(ExcelProcessIdentity released)
            {
                if (released == identity)
                    order.Enqueue("release");
            }

            SessionManager.ExcelProcessObservationRecorded += RecordObservation;
            if (rejectFirstRecord)
                SessionManager.ExcelProcessObservationRecorded += RejectObservation;
            SessionManager.ExcelProcessIdentityReleased += RecordRelease;
            try
            {
                if (rejectFirstRecord)
                {
                    var failed = await service.ProcessAsync(new ServiceRequest
                    {
                        Command = "session.close",
                        SessionId = sessionId,
                        Args = """{"save":false}"""
                    });
                    Assert.False(failed.Success);
                    Assert.Contains("Synthetic native observation persistence failure.",
                        failed.ErrorMessage + failed.InnerError, StringComparison.Ordinal);
                    Assert.Contains(sessionId, service.SessionManager.ActiveSessionIds);
                    Assert.Contains(identity, SessionManager.GetTrackedExcelProcesses());
                    Assert.DoesNotContain("release", order);
                    var retained = Assert.IsType<ExcelNativeProcessObservation>(batch.ProcessObservation);
                    Assert.Equal(identity, retained.Identity);
                    Assert.Equal(ExcelNativeProcessState.Exited, retained.State);
                    Assert.Equal(0u, retained.ExitCode);
                    Assert.Null(retained.NativeError);
                    SessionManager.ExcelProcessObservationRecorded -= RejectObservation;
                }

                await CloseSessionAsync(service, sessionId, save: false);
                sessions.TryRemove(sessionId, out _);

                Assert.NotEmpty(observed);
                Assert.All(observed, observation =>
                {
                    Assert.Equal(identity, observation.Identity);
                    Assert.Equal(ExcelNativeProcessState.Exited, observation.State);
                    Assert.Equal(0u, observation.ExitCode);
                    Assert.Null(observation.NativeError);
                });
                Assert.Equal("observation", order.First());
                // Repeated finalization may republish the same diagnostic after release.
                var events = order.ToArray();
                var releaseIndex = Array.IndexOf(events, "release");
                Assert.True(releaseIndex > 0);
                Assert.Equal("observation", events[releaseIndex - 1]);
                Assert.Single(order, item => item == "release");
                Assert.DoesNotContain(identity, SessionManager.GetTrackedExcelProcesses());
                Assert.DoesNotContain(sessionId, service.SessionManager.ActiveSessionIds);
                Assert.True(OwnedProcessGuard.TryConfirmExited(identity));
                var terminal = Assert.IsType<ExcelNativeProcessObservation>(batch.ProcessObservation);
                Assert.Equal(identity, terminal.Identity);
                Assert.Equal(ExcelNativeProcessState.Exited, terminal.State);
                Assert.Equal(0u, terminal.ExitCode);
                Assert.Null(terminal.NativeError);
                batch.Dispose();
                Assert.Equal(terminal, batch.ProcessObservation);
            }
            finally
            {
                SessionManager.ExcelProcessObservationRecorded -= RecordObservation;
                SessionManager.ExcelProcessObservationRecorded -= RejectObservation;
                SessionManager.ExcelProcessIdentityReleased -= RecordRelease;
            }
        });
    }

    [Fact]
    public async Task NativeObservation_FailedStartupFinalizationCanRepeatWithoutLosingExitOrRelease()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "observed-startup.xlsx");
            var seed = await CreateSessionAsync(service, path);
            sessions[seed] = 0;
            await CloseSessionAsync(service, seed, save: true);
            sessions.TryRemove(seed, out _);

            ExcelProcessIdentity? original = null;
            var observed = new ConcurrentQueue<ExcelNativeProcessObservation>();
            var released = new ConcurrentQueue<ExcelProcessIdentity>();
            void Track(ExcelProcessIdentity identity) => original = identity;
            void Observe(ExcelNativeProcessObservation observation) => observed.Enqueue(observation);
            void Release(ExcelProcessIdentity identity) => released.Enqueue(identity);
            SessionManager.ExcelProcessIdentityTracked += Track;
            SessionManager.ExcelProcessObservationRecorded += Observe;
            SessionManager.ExcelProcessIdentityReleased += Release;
            ExcelBatch.BeforeWorkbookOpenHook = (opening, _) =>
            {
                if (string.Equals(opening, path, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Synthetic failure before workbook open.");
            };
            try
            {
                var failed = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.open",
                    Args = System.Text.Json.JsonSerializer.Serialize(new { filePath = path })
                });

                Assert.False(failed.Success);
                Assert.Contains("Synthetic failure before workbook open.",
                    failed.ErrorMessage + failed.InnerError, StringComparison.Ordinal);
                var identity = Assert.IsType<ExcelProcessIdentity>(original);
                Assert.Empty(service.SessionManager.ActiveSessionIds);
                Assert.DoesNotContain(identity, SessionManager.GetTrackedExcelProcesses());
                Assert.Equal(identity, Assert.Single(released));
                Assert.NotEmpty(observed);
                Assert.All(observed, observation =>
                {
                    Assert.Equal(identity, observation.Identity);
                    Assert.Equal(ExcelNativeProcessState.Exited, observation.State);
                    Assert.Equal(0u, observation.ExitCode);
                    Assert.Null(observation.NativeError);
                });
                Assert.True(OwnedProcessGuard.TryConfirmExited(identity));
            }
            finally
            {
                ExcelBatch.BeforeWorkbookOpenHook = null;
                SessionManager.ExcelProcessIdentityTracked -= Track;
                SessionManager.ExcelProcessObservationRecorded -= Observe;
                SessionManager.ExcelProcessIdentityReleased -= Release;
            }
        });
    }

    [Fact]
    public async Task NativeObservation_AttachedDetachDoesNotDisposeSiblingOwnerObservation()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "observed-attachment.xlsx");
            var ownedId = await CreateSessionAsync(service, path);
            sessions[ownedId] = 0;
            await WriteMarkerAsync(service, ownedId, "Owner remains usable after attachment detaches");
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var identity = Assert.IsType<ExcelProcessIdentity>(owned.ProcessIdentity);
            await RunWithCleanupAsync(async (observer, observerDirectory, attachedSessions) =>
            {
                var attachedId = await AttachHandoverAsync(observer, path, identity.ProcessId);
                attachedSessions[attachedId] = 0;
                var attached = Assert.IsType<AttachedWorkbookBatch>(observer.SessionManager.GetSession(attachedId));
                Assert.NotSame(owned, attached.Owner);
                Assert.Equal(identity, attached.Owner.ProcessIdentity);
                Assert.Equal(ExcelNativeProcessState.Alive,
                    Assert.IsType<ExcelNativeProcessObservation>(attached.Owner.ProcessObservation).State);

                observer.SessionManager.Dispose();
                attachedSessions.TryRemove(attachedId, out _);

                Assert.Equal(ExcelNativeProcessState.Indeterminate,
                    Assert.IsType<ExcelNativeProcessObservation>(attached.Owner.ProcessObservation).State);
                var remaining = Assert.IsType<ExcelNativeProcessObservation>(owned.ProcessObservation);
                Assert.Equal(identity, remaining.Identity);
                Assert.Equal(ExcelNativeProcessState.Alive, remaining.State);
                Assert.Null(remaining.ExitCode);
                Assert.Null(remaining.NativeError);
                Assert.True(OwnedProcessGuard.MatchesLiveProcess(identity));
                Assert.Equal("Owner remains usable after attachment detaches", await ReadMarkerAsync(service, ownedId));
            });
        });
    }

    [Fact]
    public async Task NativeObservation_PreservingNeighborRecordsLiveHandoverWithoutClaimingExit()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "observed-handover.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var identity = Assert.IsType<ExcelProcessIdentity>(owned.ProcessIdentity);
            var neighborName = owned.Execute((context, _) => AddReviewWorkbook(context.App, "Preserved native observation marker"));
            await RunWithCleanupAsync(async (observer, observerDirectory, attachedSessions) =>
            {
                var attachedId = await AttachReviewWorkbookAsync(observer,
                    await DiscoverReviewWorkbookAsync(observer, neighborName));
                attachedSessions[attachedId] = 0;
                var observed = new ConcurrentQueue<ExcelNativeProcessObservation>();
                void RecordObservation(ExcelNativeProcessObservation observation)
                {
                    if (observation.Identity == identity)
                        observed.Enqueue(observation);
                }
                SessionManager.ExcelProcessObservationRecorded += RecordObservation;
                try
                {
                    await CloseSessionAsync(service, ownedId, save: false);
                    sessions.TryRemove(ownedId, out _);

                    Assert.True(SessionManager.IsExcelProcessPreserved(identity));
                    Assert.True(OwnedProcessGuard.MatchesLiveProcess(identity));
                    Assert.NotEmpty(observed);
                    Assert.All(observed, observation =>
                    {
                        Assert.Equal(identity, observation.Identity);
                        Assert.Equal(ExcelNativeProcessState.Alive, observation.State);
                        Assert.Null(observation.ExitCode);
                        Assert.Null(observation.NativeError);
                    });
                    Assert.Equal(ExcelNativeProcessState.Indeterminate,
                        Assert.IsType<ExcelNativeProcessObservation>(owned.ProcessObservation).State);
                    Assert.Equal("Preserved native observation marker", await ReadMarkerAsync(observer, attachedId));
                }
                finally
                {
                    SessionManager.ExcelProcessObservationRecorded -= RecordObservation;
                }
            });
            Assert.True(OwnedProcessGuard.TryConfirmExited(identity));
        });
    }
}
