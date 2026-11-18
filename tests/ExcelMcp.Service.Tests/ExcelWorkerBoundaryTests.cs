using System.ComponentModel;
using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service.Workers;
using Sbroenne.ExcelMcp.Service.Rpc;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("Sequential")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceDaemon")]
[Trait("Speed", "Medium")]
[Trait("RequiresExcel", "false")]
public sealed class ExcelWorkerBoundaryTests
{
    [Fact(Timeout = 20000)]
    public async Task WorkerMonitor_RetriesUnverifiableControllerAndStopsAfterConfirmedExit()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-controller-monitor-");
        var registry = new WorkerRegistry(directory.FullName);
        using var controller = Process.Start(new ProcessStartInfo("pwsh")
        {
            ArgumentList = { "-NoProfile", "-Command", "[Console]::ReadLine() | Out-Null" },
            RedirectStandardInput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        })!;
        var owner = WorkerProcess.From(controller);
        var generation = Guid.NewGuid().ToString("N");
        var registration = new WorkerRegistration("excelmcp-worker-" + generation, WorkerProcess.Current, owner,
            generation, WorkerRegistry.Build, null);
        using var host = new ExcelWorkerHost(registration, registry: registry);
        var denied = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var retried = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var probes = 0;
        WorkerProcess.BeforeIdentityReadForTests = process =>
        {
            if (process.Id != owner.ProcessId)
                return;
            if (Interlocked.Increment(ref probes) == 1)
            {
                denied.TrySetResult();
                throw new Win32Exception(WorkerProcess.ErrorAccessDenied);
            }
            retried.TrySetResult();
        };
        var run = host.RunAsync();
        var primary = await Record.ExceptionAsync(async () =>
        {
            await denied.Task.WaitAsync(TimeSpan.FromSeconds(5));
            await retried.Task.WaitAsync(TimeSpan.FromSeconds(6));
            Assert.False(run.IsCompleted);
            Assert.Equal(registration, registry.Read(registration.PipeName));

            controller.StandardInput.Close();
            await controller.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(4));
            await run.WaitAsync(TimeSpan.FromSeconds(5));
            Assert.Null(registry.Read(registration.PipeName));
        });
        WorkerProcess.BeforeIdentityReadForTests = null;
        var cleanup = await Record.ExceptionAsync(async () =>
        {
            controller.StandardInput.Close();
            await controller.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(4));
            if (!run.IsCompleted)
            {
                var server = Assert.IsType<DaemonHost>(typeof(ExcelWorkerHost)
                    .GetField("_server", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
                    .GetValue(host));
                server.RequestShutdown();
            }
            try
            {
                await run.WaitAsync(TimeSpan.FromSeconds(4));
            }
            finally
            {
                registry.Remove(registration.PipeName);
                directory.Delete();
            }
        });
        if (primary is not null && cleanup is not null) throw new AggregateException(primary, cleanup);
        if (primary is not null || cleanup is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture((primary ?? cleanup)!).Throw();
    }

    [Fact(Timeout = 15000)]
    public async Task WorkerSnapshot_RemainsPendingUntilCompletionStatusIsRecorded()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-completion-status-");
        var registry = new WorkerRegistry(directory.FullName);
        var process = WorkerProcess.Current;
        var generation = Guid.NewGuid().ToString("N");
        var registration = new WorkerRegistration("excelmcp-worker-" + generation, process, process,
            generation, WorkerRegistry.Build, null);
        using var host = new ExcelWorkerHost(registration, registry: registry);
        var run = host.RunAsync();
        var worker = new ExcelWorkerClient(registration, process, generation);
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        using var resume = new ManualResetEventSlim();
        await worker.InitializeAsync();
        host.AfterOperationStatusCaptureForTests = () =>
        {
            entered.TrySetResult();
            if (!resume.Wait(TimeSpan.FromSeconds(5)))
                throw new TimeoutException("The test did not finish reading the in-flight worker status.");
        };
        var execution = worker.ExecuteWithAdmissionAsync(new ServiceRequest
        {
            Command = "diag.echo",
            Args = """{"message":"write complete, status pending"}"""
        }, CancellationToken.None);
        try
        {
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(3));
            var snapshot = await worker.ReadSnapshotAsync();
            Assert.NotNull(snapshot.PendingOperation);
            Assert.Equal("diag.echo", snapshot.PendingCommand);
            Assert.False(execution.IsCompleted);
        }
        finally
        {
            resume.Set();
            var completed = await execution.WaitAsync(TimeSpan.FromSeconds(3));
            Assert.True(completed.Response.Success, completed.Response.ErrorMessage);
            Assert.Null(completed.Snapshot.PendingOperation);
            Assert.Null(completed.Snapshot.PendingCommand);
            host.AfterOperationStatusCaptureForTests = null;
            var released = await worker.ReleaseAsync(CancellationToken.None);
            Assert.True(released.Success, released.ErrorMessage);
            await run.WaitAsync(TimeSpan.FromSeconds(3));
            registry.Remove(registration.PipeName);
            directory.Delete();
        }
    }

    [Theory(Timeout = 30000)]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PrivateWorker_ExecutesAndReportsStatusWithoutStartingExcel(bool extendedLibraryDirectory)
    {
        var executable = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            @"..\..\..\..\..\src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe"));
        Assert.True(File.Exists(executable), "Build the CLI product before running private-worker process tests.");
        var originalDirectory = AppContext.GetData("APP_CONTEXT_BASE_DIRECTORY");
        var libraryDirectory = extendedLibraryDirectory
            ? Path.Combine(Path.GetTempPath(), $"excelmcp-launcher-{Guid.NewGuid():N}", new string('x', 220))
            : null;
        ExcelWorkerClient worker;
        try
        {
            if (libraryDirectory is not null)
            {
                Directory.CreateDirectory(libraryDirectory);
                var extended = @"\\?\" + libraryDirectory + Path.DirectorySeparatorChar;
                AppContext.SetData("APP_CONTEXT_BASE_DIRECTORY", extended);
                Assert.Equal(extended, AppContext.BaseDirectory);
            }
            worker = ExcelWorkerClient.Launch(new ProductProcessLauncher(executable, []),
                WorkerProcess.Current, Guid.NewGuid().ToString("N"), null);
        }
        finally
        {
            AppContext.SetData("APP_CONTEXT_BASE_DIRECTORY", originalDirectory);
            if (libraryDirectory is not null)
            {
                Directory.Delete(libraryDirectory);
                Directory.Delete(Path.GetDirectoryName(libraryDirectory)!);
            }
        }
        try
        {
            await worker.InitializeAsync();
            var initialRevision = worker.Snapshot!.Revision;
            var message = new string('x', 1024 * 1024);
            var response = await worker.ExecuteAsync(new ServiceRequest
            {
                Command = "diag.echo",
                Args = System.Text.Json.JsonSerializer.Serialize(new { message })
            }, CancellationToken.None);
            Assert.True(response.Success, response.ErrorMessage);
            Assert.Equal(message, System.Text.Json.JsonDocument.Parse(
                response.Result!).RootElement.GetProperty("message").GetString());
            Assert.True(worker.Snapshot!.Revision > initialRevision,
                "The operation reply must update recorded status without another status request.");
            Assert.Null(worker.Snapshot.PendingOperation);
            var snapshot = await worker.ReadSnapshotAsync();
            Assert.Empty(snapshot.Sessions);
            Assert.Null(snapshot.Registration.Excel);
            Assert.Null(snapshot.PendingOperation);
            Assert.True(System.Text.Json.JsonSerializer.SerializeToUtf8Bytes(snapshot).Length < 65536,
                "Status must not repeat the last command's large result.");
            Assert.NotEqual(Environment.ProcessId, snapshot.Registration.Worker.ProcessId);
            Assert.Equal(WorkerProcess.Current, snapshot.Registration.Controller);
        }
        finally
        {
            using var process = Process.GetProcessById(worker.Registration.Worker.ProcessId);
            Assert.Equal(worker.Registration.Worker, WorkerProcess.From(process));
            _ = process.Handle;
            var released = await worker.ReleaseAsync(CancellationToken.None);
            Assert.True(released.Success, released.ErrorMessage);
            await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
            Assert.Equal(0, process.ExitCode);
            new WorkerRegistry().Remove(worker.Registration.PipeName);
        }
    }

    [Fact]
    public void WorkerSnapshot_DelayedOlderReplyDoesNotReplaceNewerState()
    {
        var registration = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}",
            WorkerProcess.Current, WorkerProcess.Current, "generation", WorkerRegistry.Build, null);
        var worker = new ExcelWorkerClient(registration, WorkerProcess.Current, "generation");
        var newer = new WorkerSnapshot(registration, [],
            new ServiceResponse { Success = true, Result = """{"sessions":[]}""" },
            "pending-operation", "range.set-values",
            new Dictionary<string, IReadOnlyList<string>>(), Revision: 2);
        worker.RecordSnapshot(newer);
        var retained = worker.RecordSnapshot(newer with { PendingOperation = null, Revision = 1 });
        Assert.Same(newer, retained);
        Assert.Same(newer, worker.Snapshot);
        Assert.Equal("pending-operation", worker.Snapshot!.PendingOperation);
    }

    [Fact(Timeout = 30000)]
    public async Task WorkerExit_IsReportedAsTransportFailureNotInvalidInput()
    {
        var executable = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            @"..\..\..\..\..\src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe"));
        var worker = ExcelWorkerClient.Launch(new ProductProcessLauncher(executable, []),
            WorkerProcess.Current, Guid.NewGuid().ToString("N"), null);
        await worker.InitializeAsync();
        using var process = Process.GetProcessById(worker.Registration.Worker.ProcessId);
        Assert.Equal(worker.Registration.Worker, WorkerProcess.From(process));
        Assert.Null(worker.Snapshot!.Registration.Excel);
        Assert.Empty(worker.Snapshot.Sessions);
        _ = process.Handle;
        process.Kill();
        await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(3));
        var elapsed = Stopwatch.StartNew();
        var error = await Assert.ThrowsAsync<IOException>(() => worker.ReadSnapshotAsync());
        Assert.Contains("worker", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(elapsed.Elapsed < TimeSpan.FromSeconds(2));
    }

    [Fact(Timeout = 15000)]
    public async Task WorkerCancellation_RetainsTheExactLateResponseWithoutRepeatingWork()
    {
        var pipeName = $"excelmcp-worker-{Guid.NewGuid():N}";
        var target = new DeferredWorker(pipeName);
        using var host = new DaemonHost(pipeName,
            pipe => new StreamJsonRpc.JsonRpc(pipe, pipe, target), maximumConnections: 8);
        var run = host.RunAsync(pipeName);
        using var cancellation = new CancellationTokenSource();
        var worker = new ExcelWorkerClient(new(pipeName, WorkerProcess.Current, WorkerProcess.Current,
            "generation", WorkerRegistry.Build, null), WorkerProcess.Current, "generation");
        var execution = worker.ExecuteAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = """{"filePath":"C:\\missing.xlsx"}"""
        }, cancellation.Token);
        try
        {
            await target.Started.Task.WaitAsync(TimeSpan.FromSeconds(3));
            cancellation.Cancel();
            await target.Cancelled.Task.WaitAsync(TimeSpan.FromSeconds(3));
            await Task.Delay(100);
            Assert.False(execution.IsCompleted);
            target.Finish.TrySetResult();
            var response = await execution.WaitAsync(TimeSpan.FromSeconds(3));
            Assert.False(response.Success);
            Assert.Equal("FileNotFoundException", response.ExceptionType);
            Assert.Equal("The requested workbook does not exist.", response.ErrorMessage);
            Assert.Null(response.PartialState);
            Assert.Equal(1, target.Calls);
        }
        finally
        {
            target.Finish.TrySetResult();
            host.RequestShutdown();
            await run.WaitAsync(TimeSpan.FromSeconds(3));
        }
    }

    [Fact]
    public void ProcessIdentity_RejectsReusedProcessNumber()
    {
        var current = WorkerProcess.Current;
        Assert.True(current.IsAlive());
        Assert.False((current with { StartedAtUtcFileTime = current.StartedAtUtcFileTime - 1 }).IsAlive());
    }

    [Theory(Timeout = 15000)]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task ProcessIdentity_ExitDuringReadDoesNotHideLiveProbeFailures(bool remainAlive, bool accessDenied)
    {
        using var process = Process.Start(new ProcessStartInfo("pwsh")
        {
            ArgumentList = { "-NoProfile", "-Command", "[Console]::ReadLine() | Out-Null" },
            RedirectStandardInput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        })!;
        var identity = WorkerProcess.From(process);
        _ = process.Handle;
        try
        {
            WorkerProcess.BeforeIdentityReadForTests = probe =>
            {
                if (probe.Id != identity.ProcessId)
                    return;
                if (!remainAlive)
                {
                    process.StandardInput.Close();
                    Assert.True(process.WaitForExit(5000));
                    probe.Refresh();
                }
                if (accessDenied)
                    throw new Win32Exception(5);
                if (remainAlive)
                    throw new InvalidOperationException("Live process identity could not be read.");
            };
            if (remainAlive && accessDenied)
                Assert.Equal(5, Assert.Throws<Win32Exception>(() => identity.IsAlive()).NativeErrorCode);
            else if (remainAlive)
                Assert.Equal("Live process identity could not be read.",
                    Assert.Throws<InvalidOperationException>(() => identity.IsAlive()).Message);
            else
                Assert.False(identity.IsAlive());
        }
        finally
        {
            WorkerProcess.BeforeIdentityReadForTests = null;
            process.StandardInput.Close();
            await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
        }
        Assert.False(identity.IsAlive());
    }

    [Fact]
    public void ProductLauncher_DoesNotRelaunchTestHost()
    {
        var error = Assert.Throws<InvalidOperationException>(ProductProcessLauncher.Current);
        Assert.Contains("built product launcher", error.Message);
    }

    [Fact]
    public void Registry_RecordRemovedAfterEnumerationDoesNotFailDiscovery()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"excelmcp-workers-{Guid.NewGuid():N}");
        var registry = new WorkerRegistry(directory);
        var current = WorkerProcess.Current;
        var record = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}",
            current, current, "generation", WorkerRegistry.Build, null);
        try
        {
            registry.Write(record);
            registry.BeforeRecordReadForTests = File.Delete;
            Assert.Empty(registry.Read());
            registry.BeforeRecordReadForTests = null;
            Assert.Null(registry.Read(record.PipeName));
        }
        finally
        {
            registry.Remove(record.PipeName);
            Directory.Delete(directory);
        }
    }

    [Theory]
    [InlineData("syntax")]
    [InlineData("size")]
    [InlineData("name")]
    [InlineData("identity")]
    [InlineData("worker-pid")]
    [InlineData("controller-start")]
    [InlineData("excel-pid")]
    [InlineData("released-start")]
    public void Registry_UnconfirmedRecordNamesItsFileAndPreservesIt(string defect)
    {
        var directory = Path.Combine(Path.GetTempPath(), $"excelmcp-workers-{Guid.NewGuid():N}");
        var registry = new WorkerRegistry(directory);
        var current = WorkerProcess.Current;
        var record = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}",
            current, current, "generation", WorkerRegistry.Build, null);
        var filenamePipe = defect == "name" ? $"excelmcp-worker-{Guid.NewGuid():N}" : record.PipeName;
        string basename = filenamePipe + ".json";
        string path = Path.Combine(directory, basename);
        try
        {
            switch (defect)
            {
                case "syntax":
                    File.WriteAllText(path, "{");
                    break;
                case "size":
                    File.WriteAllText(path, new string(' ', 65537));
                    break;
                case "name":
                    registry.Write(record);
                    File.Move(Path.Combine(directory, record.PipeName + ".json"), path);
                    break;
                case "identity":
                    registry.Write(record with { Worker = null! });
                    break;
                case "worker-pid":
                    registry.Write(record with { Worker = current with { ProcessId = 0 } });
                    break;
                case "controller-start":
                    registry.Write(record with { Controller = current with { StartedAtUtcFileTime = 0 } });
                    break;
                case "excel-pid":
                    registry.Write(record with { Excel = new ExcelProcessIdentity(0, current.StartedAtUtcFileTime) });
                    break;
                case "released-start":
                    registry.Write(record with { ReleasedExcel = new ExcelProcessIdentity(current.ProcessId, 0) });
                    break;
            }

            byte[] before = File.ReadAllBytes(path);
            var error = Assert.Throws<InvalidDataException>(() => registry.Read());
            Assert.Contains(basename, error.Message, StringComparison.Ordinal);
            Assert.Contains("ownership is unconfirmed", error.Message, StringComparison.Ordinal);
            Assert.Contains("record was left untouched", error.Message, StringComparison.Ordinal);
            var directError = Assert.Throws<InvalidDataException>(() => registry.Read(filenamePipe));
            Assert.Equal(error.Message, directError.Message);
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
            registry.Remove(record.PipeName);
            Directory.Delete(directory);
        }
    }

    [Fact]
    public void Registry_AtomicReplacementCanCompleteWhileRecordIsOpen()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"excelmcp-workers-{Guid.NewGuid():N}");
        var registry = new WorkerRegistry(directory);
        var record = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}",
            WorkerProcess.Current, WorkerProcess.Current, "original", WorkerRegistry.Build, null);
        var replacement = record with { Generation = "replacement" };
        try
        {
            registry.Write(record);
            registry.AfterRecordOpenForTests = () => registry.Write(replacement);
            Assert.Equal(record, Assert.Single(registry.Read()));
            registry.AfterRecordOpenForTests = null;
            Assert.Equal(replacement, Assert.Single(registry.Read()));
        }
        finally
        {
            registry.Remove(record.PipeName);
            File.Delete(Path.Combine(directory, record.PipeName + ".json.tmp"));
            Directory.Delete(directory);
        }
    }

    [Fact(Timeout = 15000)]
    public async Task Registry_AccessDeniedPreservesUnconfirmedRecordsAndContinuesReading()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-denied-worker-");
        var registry = new WorkerRegistry(directory.FullName);
        var current = WorkerProcess.Current;
        var exited = current with { StartedAtUtcFileTime = current.StartedAtUtcFileTime - 1 };
        using var process = Process.Start(new ProcessStartInfo("pwsh")
        {
            ArgumentList = { "-NoProfile", "-Command", "[Console]::ReadLine() | Out-Null" },
            RedirectStandardInput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        })!;
        var denied = WorkerProcess.From(process);
        WorkerRegistration CreateRecord(WorkerProcess worker, WorkerProcess controller, bool released = false) =>
            new($"excelmcp-worker-{Guid.NewGuid():N}", worker, controller, "generation", WorkerRegistry.Build, null,
                ReleaseConfirmed: released);
        var inaccessibleWorker = CreateRecord(denied, exited);
        var inaccessibleController = CreateRecord(exited, denied);
        var inaccessibleReleaseOwner = CreateRecord(exited, denied, released: true);
        var live = CreateRecord(current, current);
        var stale = CreateRecord(exited, exited);
        var records = new[] { inaccessibleWorker, inaccessibleController, inaccessibleReleaseOwner, live, stale };
        try
        {
            foreach (var record in records)
                registry.Write(record);
            WorkerProcess.BeforeIdentityReadForTests = probe =>
            {
                if (probe.Id == denied.ProcessId)
                    throw new Win32Exception(5);
            };

            var retained = registry.Read();
            Assert.Equal(3, retained.Count);
            Assert.Contains(inaccessibleWorker, retained);
            Assert.Contains(inaccessibleReleaseOwner, retained);
            Assert.Contains(live, retained);
            Assert.Equal(inaccessibleWorker, registry.Read(inaccessibleWorker.PipeName));
            Assert.Equal(inaccessibleController, registry.Read(inaccessibleController.PipeName));
            Assert.Equal(inaccessibleReleaseOwner, registry.Read(inaccessibleReleaseOwner.PipeName));
            Assert.Null(registry.Read(stale.PipeName));
        }
        finally
        {
            WorkerProcess.BeforeIdentityReadForTests = null;
            process.StandardInput.Close();
            await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
            foreach (var record in records)
                registry.Remove(record.PipeName);
            directory.Delete();
        }
    }

    [Theory(Timeout = 10000)]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Registry_ReadReportsPersistentLocksAndRecoversAfterTemporaryLocks(bool releaseLock)
    {
        var directory = Path.Combine(Path.GetTempPath(), $"excelmcp-locked-worker-{Guid.NewGuid():N}");
        var registry = new WorkerRegistry(directory);
        var record = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}",
            WorkerProcess.Current, WorkerProcess.Current, "generation", WorkerRegistry.Build, null,
            ReleaseConfirmed: true);
        var path = Path.Combine(directory, record.PipeName + ".json");
        try
        {
            registry.Write(record);
            using (var locked = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
                registry.BeforeRecordReadForTests = _ => started.TrySetResult();
                var reading = Task.Run(() => registry.Read(record.PipeName));
                await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
                if (releaseLock)
                {
                    await Task.Delay(50);
                    locked.Dispose();
                    Assert.Equal(record, await reading.WaitAsync(TimeSpan.FromSeconds(5)));
                }
                else
                {
                    var error = await Assert.ThrowsAsync<IOException>(
                        () => reading.WaitAsync(TimeSpan.FromSeconds(5)));
                    Assert.True((error.HResult & 0xFFFF) is 32 or 33);
                    Assert.True(File.Exists(path));
                }
            }
            Assert.Equal(record, registry.Read(record.PipeName));
        }
        finally
        {
            registry.Remove(record.PipeName);
            Directory.Delete(directory);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Registry_CompletedReleaseRemainsAvailableUntilAcknowledged(bool controllerAlive)
    {
        var directory = Path.Combine(Path.GetTempPath(), $"excelmcp-completed-worker-{Guid.NewGuid():N}");
        var registry = new WorkerRegistry(directory);
        var current = WorkerProcess.Current;
        var stopped = current with { StartedAtUtcFileTime = current.StartedAtUtcFileTime - 1 };
        var released = new ExcelProcessIdentity(current.ProcessId, current.StartedAtUtcFileTime);
        var record = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}",
            stopped, controllerAlive ? current : stopped, "generation", WorkerRegistry.Build, null,
            ReleaseConfirmed: true, ReleasedExcel: released);
        try
        {
            registry.Write(record);
            Assert.Equal(record, registry.Read(record.PipeName));
            if (controllerAlive)
            {
                Assert.Equal(record, Assert.Single(registry.Read()));
                Assert.False(record.OwnsExcel);
                Assert.Null(record.Excel);
                registry.Remove(record.PipeName);
            }
            else
                Assert.Empty(registry.Read());
            Assert.Null(registry.Read(record.PipeName));
        }
        finally
        {
            registry.Remove(record.PipeName);
            Directory.Delete(directory);
        }
    }

    [Fact]
    public void Registry_DeadWorkerRetainsLiveApplicationIdentity()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"excelmcp-workers-{Guid.NewGuid():N}");
        var registry = new WorkerRegistry(directory);
        var current = WorkerProcess.Current;
        var record = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}",
            current with { StartedAtUtcFileTime = current.StartedAtUtcFileTime - 1 },
            current, "generation", WorkerRegistry.Build,
            new ExcelProcessIdentity(current.ProcessId, current.StartedAtUtcFileTime));
        try
        {
            registry.Write(record);
            Assert.Equal(record, Assert.Single(registry.Read()));
            Assert.True(File.Exists(Path.Combine(directory, record.PipeName + ".json")));
        }
        finally
        {
            registry.Remove(record.PipeName);
            Directory.Delete(directory);
        }
    }

    [Fact(Timeout = 10000)]
    public async Task PipeInput_ClosesOversizedIncompleteMessageWithoutExecutingIt()
    {
        var pipeName = $"excelmcp-input-limit-{Guid.NewGuid():N}";
        var calls = 0;
        using var host = new DaemonHost(_ =>
        {
            Interlocked.Increment(ref calls);
            return Task.FromResult(new ServiceResponse { Success = true });
        });
        var run = host.RunAsync(pipeName);
        try
        {
            await using var pipe = ServiceSecurity.CreateClient(pipeName);
            await pipe.ConnectAsync(2000);
            const int limit = 64 * 1024 * 1024;
            await pipe.WriteAsync(System.Text.Encoding.ASCII.GetBytes($"Content-Length: {limit + 1048576}\r\n\r\n"));
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            try
            {
                var block = new byte[65536];
                for (var sent = 0; sent <= limit; sent += block.Length)
                    await pipe.WriteAsync(block, timeout.Token);
                Assert.Equal(0, await pipe.ReadAsync(new byte[1], timeout.Token));
            }
            catch (IOException)
            {
                // Closing an oversized pipe can fail its pending write instead of returning EOF.
            }
            Assert.Equal(0, Volatile.Read(ref calls));
            using var healthy = new ServiceClient(pipeName);
            Assert.True((await healthy.SendAsync(new ServiceRequest { Command = "service.ping" },
                TimeSpan.FromSeconds(2))).Success);
            Assert.Equal(1, Volatile.Read(ref calls));
        }
        finally
        {
            host.RequestShutdown();
            await run.WaitAsync(TimeSpan.FromSeconds(3));
        }
    }

    [Fact(Timeout = 10000)]
    public async Task WorkerStatus_NoResponseReportsTimeoutNotCallerCancellation()
    {
        var pipeName = $"excelmcp-worker-{Guid.NewGuid():N}";
        await using var server = ServiceSecurity.CreateSecureServer(pipeName);
        var connected = server.WaitForConnectionAsync();
        var worker = new ExcelWorkerClient(new(pipeName, WorkerProcess.Current, WorkerProcess.Current,
            "generation", WorkerRegistry.Build, null), WorkerProcess.Current, "generation");
        var read = worker.ReadSnapshotAsync();
        await connected.WaitAsync(TimeSpan.FromSeconds(3));
        await Assert.ThrowsAsync<TimeoutException>(() => read);
    }

    private sealed class DeferredWorker(string pipeName) : IExcelWorkerRpc
    {
        internal TaskCompletionSource Started { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource Cancelled { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource Finish { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal int Calls { get; private set; }

        public async Task<WorkerExecution> ExecuteAsync(WorkerCall call, CancellationToken cancellationToken = default)
        {
            Calls++;
            using var registration = cancellationToken.Register(() => Cancelled.TrySetResult());
            Started.TrySetResult();
            await Finish.Task;
            return new(new ServiceResponse
            {
                Success = false,
                ExceptionType = "FileNotFoundException",
                ErrorMessage = "The requested workbook does not exist."
            }, Accepted: true, new WorkerSnapshot(
                new(pipeName, WorkerProcess.Current, WorkerProcess.Current, call.Generation, WorkerRegistry.Build, null),
                [], new ServiceResponse { Success = true, Result = """{"sessions":[]}""" },
                null, null, new Dictionary<string, IReadOnlyList<string>>(), Revision: 1));
        }

        public Task<WorkerSnapshot> ClaimAsync(WorkerClaim claim, CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();

        public Task<WorkerSnapshot> SnapshotAsync(string generation, CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();

        public Task<ServiceResponse?> ObserveLateCompletionAsync(string generation, string operationId,
            CancellationToken cancellationToken = default) => Task.FromResult<ServiceResponse?>(null);

        public Task<WorkerRelease> ReleaseAsync(string generation, CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();
    }
}
