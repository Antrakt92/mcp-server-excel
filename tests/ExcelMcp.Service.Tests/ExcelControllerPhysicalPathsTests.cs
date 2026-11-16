using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service.Workers;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("Sequential")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceDaemon")]
[Trait("Speed", "Medium")]
[Trait("RequiresExcel", "false")]
public sealed class ExcelControllerPhysicalPathsTests
{
    [Fact]
    public void Controller_PublishesKnownWorkerReservationThroughHardLink()
    {
        var root = Path.Join(Path.GetTempPath(), "excelmcp-controller-publish-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var primary = Record.Exception(() =>
        {
            var original = Path.Join(root, "original.xlsx");
            var reported = Path.Join(root, "reported.xlsx");
            File.WriteAllText(original, "synthetic");
            if (!CreateHardLink(reported, original, IntPtr.Zero))
                throw new Win32Exception(Marshal.GetLastWin32Error());

            using var controller = new ExcelController(launcher: new("unused.exe", []));
            var process = WorkerProcess.Current;
            var generation = Guid.NewGuid().ToString("N");
            var pipe = "excelmcp-worker-" + generation;
            var worker = new ExcelWorkerClient(
                new WorkerRegistration(pipe, process, process, generation, WorkerRegistry.Build, null), process, generation);
            const BindingFlags hidden = BindingFlags.Instance | BindingFlags.NonPublic;
            var type = typeof(ExcelController);
            var claims = type.GetField("_paths", hidden)!.GetValue(controller)!;
            var reservation = Activator.CreateInstance(type.GetNestedType("PathReservation", BindingFlags.NonPublic)!, pipe, null)!;
            Assert.True((bool)claims.GetType().GetMethod("TryAdd", hidden)!.Invoke(claims, [original, reservation])!);

            type.GetMethod("RecordPath", hidden)!.Invoke(controller, [worker, reported, "session"]);

            var lookup = claims.GetType().GetMethod("TryGetValue", hidden)!;
            foreach (var path in new[] { original, reported })
            {
                object?[] args = [path, null];
                Assert.True((bool)lookup.Invoke(claims, args)!);
                Assert.Equal("session", args[1]!.GetType().GetProperty("SessionId")!.GetValue(args[1]));
            }
            Assert.Equal("synthetic", File.ReadAllText(reported));
        });
        var cleanup = Record.Exception(() => Directory.Delete(root, recursive: true));
        if (primary is not null && cleanup is not null) throw new AggregateException(primary, cleanup);
        if (primary is not null || cleanup is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture((primary ?? cleanup)!).Throw();
    }

    [Theory(Timeout = 30000)]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Controller_RejectsHardLinkWhileOriginalOpenOrCreateIsPending(bool create)
    {
        var root = Path.Join(Path.GetTempPath(), "excelmcp-controller-identity-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var path = Path.Join(root, "original.xlsx");
        var alias = Path.Join(root, "alias.xlsx");
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var starts = 0;
        var launcher = new ProductProcessLauncher(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            @"..\..\..\..\..\src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe")), []);
        var primary = await Record.ExceptionAsync(async () =>
        {
            if (!create) File.WriteAllText(path, "synthetic");
            using var controller = new ExcelController(launcher: launcher, initializeWorker: async worker =>
            {
                await worker.InitializeAsync();
                if (Interlocked.Increment(ref starts) == 1)
                {
                    entered.TrySetResult();
                    await release.Task;
                }
            });
            // The deliberately invalid timeout is rejected inside the worker before any Excel startup.
            var pending = controller.ProcessAsync(new ServiceRequest
            {
                Command = create ? "session.create" : "session.open",
                Args = JsonSerializer.Serialize(new { filePath = path, timeoutSeconds = 1 })
            });
            try
            {
                await entered.Task.WaitAsync(TimeSpan.FromSeconds(10));
                if (create) File.WriteAllText(path, "synthetic");
                if (!CreateHardLink(alias, path, IntPtr.Zero)) throw new Win32Exception(Marshal.GetLastWin32Error());
                var duplicate = await controller.ProcessAsync(new ServiceRequest
                {
                    Command = "session.open",
                    Args = JsonSerializer.Serialize(new { filePath = alias, timeoutSeconds = 1 })
                });
                Assert.False(duplicate.Success);
                Assert.Equal("Conflict", duplicate.ErrorCategory);
                Assert.Equal(1, starts);
            }
            finally
            {
                release.TrySetResult();
                var response = await pending.WaitAsync(TimeSpan.FromSeconds(10));
                Assert.False(response.Success);
                Assert.Equal("ArgumentOutOfRangeException", response.ExceptionType);
            }
            Assert.Equal("synthetic", File.ReadAllText(path));
        });
        var cleanup = Record.Exception(() => Directory.Delete(root, recursive: true));
        if (primary is not null && cleanup is not null) throw new AggregateException(primary, cleanup);
        if (primary is not null || cleanup is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture((primary ?? cleanup)!).Throw();
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CreateHardLink(string alias, string original, IntPtr security);
}
