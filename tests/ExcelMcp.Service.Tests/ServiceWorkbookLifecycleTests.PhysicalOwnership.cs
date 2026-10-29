using System.Collections.Concurrent;
using System.ComponentModel;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.Win32.SafeHandles;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service.Workers;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Controller_PhysicalAliasOpenRejectsConflictWithoutAnotherExcelOrChangingSavedBytes(bool hardLink)
    {
        await WithPhysicalWorkbookAsync(async (controller, path, junction, workerStarts) =>
        {
            var opened = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = path })
            });
            RequireSuccess(opened);
            var id = Assert.IsType<string>(opened.SessionId);
            var identity = Assert.Single(controller.ExcelProcesses);
            Assert.Equal(1, workerStarts());
            byte[] before = await HashPhysicalWorkbookAsync(path);
            string alternate = Path.Combine(junction, Path.GetFileName(path));
            if (hardLink)
            {
                alternate = Path.Combine(Path.GetDirectoryName(junction)!, "hard-link.xlsx");
                if (!PhysicalCreateHardLink(alternate, path, IntPtr.Zero)) throw new Win32Exception(Marshal.GetLastWin32Error());
            }
            var refused = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = alternate })
            });
            Assert.False(refused.Success);
            Assert.Equal("Conflict", refused.ErrorCategory);
            Assert.Equal(1, workerStarts());
            Assert.Equal(identity, Assert.Single(controller.ExcelProcesses));
            Assert.Equal(id, Assert.Single(controller.Sessions).SessionId);
            Assert.Equal("physical ownership marker", await ReadPhysicalMarkerAsync(controller, id));
            Assert.Equal(before, await HashPhysicalWorkbookAsync(path));
            RequireSuccess(await controller.ProcessAsync(new ServiceRequest { Command = "session.close", SessionId = id, Args = "{\"save\":false}" }));
            Assert.Empty(controller.Sessions);
            Assert.Equal(before, await HashPhysicalWorkbookAsync(path));
        });
    }

    [Fact]
    public async Task Controller_SaveAsThroughJunctionRetainsDestinationAndReleasesOldPath()
    {
        await WithPhysicalWorkbookAsync(async (controller, original, junction, workerStarts) =>
        {
            byte[] originalHash = await HashPhysicalWorkbookAsync(original);
            var opened = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = original })
            });
            RequireSuccess(opened);
            var id = Assert.IsType<string>(opened.SessionId);
            var target = Path.Combine(junction, "saved-as.xlsx");
            var physicalTarget = Path.Combine(Path.GetDirectoryName(original)!, "saved-as.xlsx");
            RequireSuccess(await controller.ProcessAsync(new ServiceRequest
            {
                Command = "workbook.save-as",
                SessionId = id,
                Args = JsonSerializer.Serialize(new { targetPath = target })
            }));
            Assert.True(File.Exists(physicalTarget));
            Assert.True(WorkbookLocation.MatchesReportedLocation(Assert.Single(controller.Sessions).FilePath!, physicalTarget));
            Assert.Equal("physical ownership marker", await ReadPhysicalMarkerAsync(controller, id));
            var refused = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = physicalTarget })
            });
            Assert.False(refused.Success);
            Assert.Equal("Conflict", refused.ErrorCategory);
            Assert.Equal(1, workerStarts());
            var old = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = original })
            });
            RequireSuccess(old);
            Assert.NotEqual(id, old.SessionId);
            Assert.Equal(2, workerStarts());
            Assert.Equal(2, controller.Sessions.Count);
            Assert.Equal("physical ownership marker", await ReadPhysicalMarkerAsync(controller, old.SessionId!));
            Assert.Equal(originalHash, await HashPhysicalWorkbookAsync(original));
            byte[] targetHash = await HashPhysicalWorkbookAsync(physicalTarget);
            RequireSuccess(await controller.ProcessAsync(new ServiceRequest { Command = "session.close", SessionId = old.SessionId, Args = "{\"save\":false}" }));
            RequireSuccess(await controller.ProcessAsync(new ServiceRequest { Command = "session.close", SessionId = id, Args = "{\"save\":false}" }));
            Assert.Empty(controller.Sessions);
            Assert.Equal(targetHash, await HashPhysicalWorkbookAsync(physicalTarget));
        });
    }

    private static async Task<byte[]> HashPhysicalWorkbookAsync(string path)
    {
        await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete, 4096, FileOptions.Asynchronous | FileOptions.SequentialScan);
        return await SHA256.HashDataAsync(stream);
    }
    private static async Task WithPhysicalWorkbookAsync(Func<ExcelController, string, string, Func<int>, Task> body)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var actual = Path.Combine(directory, "actual");
            var junction = Path.Combine(directory, "junction");
            Directory.CreateDirectory(actual);
            var path = Path.Combine(actual, "original.xlsx");
            var seed = await CreateSessionAsync(service, path);
            sessions[seed] = 0;
            await WriteMarkerAsync(service, seed, "physical ownership marker");
            await CloseSessionAsync(service, seed, save: true);
            sessions.TryRemove(seed, out _);
            CreatePhysicalJunction(junction, actual);
            var starts = 0;
            var observed = new ConcurrentDictionary<ExcelProcessIdentity, byte>();
            var executable = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
                @"..\..\..\..\..\src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe"));
            var controller = new ExcelController(launcher: new ProductProcessLauncher(executable, []), initializeWorker: async worker =>
            {
                Interlocked.Increment(ref starts);
                await worker.InitializeAsync();
            });
            controller.ExcelProcessObserved += identity => observed.TryAdd(identity, 0);
            Exception? failure = await Record.ExceptionAsync(() => body(controller, path, junction, () => Volatile.Read(ref starts)));
            foreach (var session in controller.Sessions)
            {
                var cleanup = await Record.ExceptionAsync(async () => RequireSuccess(await controller.ProcessAsync(new ServiceRequest
                {
                    Command = "session.close",
                    SessionId = session.SessionId,
                    Args = "{\"save\":false}"
                })));
                if (cleanup is not null) failure = PersistentServiceCleanupFailures.Combine(failure, cleanup);
            }
            var shutdown = await Record.ExceptionAsync(async () => Assert.Null(await controller.RequestShutdownAsync()));
            if (shutdown is not null) failure = PersistentServiceCleanupFailures.Combine(failure, shutdown);
            var disposal = Record.Exception(controller.Dispose);
            if (disposal is not null) failure = PersistentServiceCleanupFailures.Combine(failure, disposal);
            if (TestExcelProcessScope.FindLeftovers(observed.Keys.ToArray()) is { } leftover)
                failure = PersistentServiceCleanupFailures.Combine(failure, leftover);
            var unlink = Record.Exception(() => Directory.Delete(junction));
            if (unlink is not null) failure = PersistentServiceCleanupFailures.Combine(failure, unlink);
            if (failure is not null) ExceptionDispatchInfo.Capture(failure).Throw();
        });
    }

    private static async Task<string?> ReadPhysicalMarkerAsync(ExcelController controller, string id)
    {
        var response = await controller.ProcessAsync(new ServiceRequest
        {
            Command = "range.get-values",
            SessionId = id,
            Args = "{\"sheetName\":\"Sheet1\",\"rangeAddress\":\"A1\"}"
        });
        RequireSuccess(response);
        using var result = JsonDocument.Parse(response.Result!);
        return result.RootElement.GetProperty("values")[0][0].GetString();
    }

    private static void CreatePhysicalJunction(string alias, string target)
    {
        Directory.CreateDirectory(alias);
        using var handle = PhysicalCreateFile(alias, 0x40000000, 7, IntPtr.Zero, 3, 0x02200000, IntPtr.Zero);
        if (handle.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
        byte[] substitute = Encoding.Unicode.GetBytes(@"\??\" + target);
        byte[] print = Encoding.Unicode.GetBytes(target);
        var data = new byte[16 + substitute.Length + 2 + print.Length + 2];
        BitConverter.GetBytes(0xA0000003u).CopyTo(data, 0);
        BitConverter.GetBytes((ushort)(data.Length - 8)).CopyTo(data, 4);
        BitConverter.GetBytes((ushort)substitute.Length).CopyTo(data, 10);
        BitConverter.GetBytes((ushort)(substitute.Length + 2)).CopyTo(data, 12);
        BitConverter.GetBytes((ushort)print.Length).CopyTo(data, 14);
        substitute.CopyTo(data, 16);
        print.CopyTo(data, 18 + substitute.Length);
        if (!PhysicalDeviceIoControl(handle, 0x000900A4, data, data.Length, IntPtr.Zero, 0, out _, IntPtr.Zero))
            throw new Win32Exception(Marshal.GetLastWin32Error());
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "CreateHardLinkW")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PhysicalCreateHardLink(string alias, string original, IntPtr security);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "CreateFileW")]
    private static extern SafeFileHandle PhysicalCreateFile(string path, uint access, uint share, IntPtr security, uint mode, uint flags, IntPtr template);
    [DllImport("kernel32.dll", SetLastError = true, EntryPoint = "DeviceIoControl")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PhysicalDeviceIoControl(SafeFileHandle handle, uint code, byte[] input, int inputLength, IntPtr output, int outputLength, out int returned, IntPtr overlapped);
}
