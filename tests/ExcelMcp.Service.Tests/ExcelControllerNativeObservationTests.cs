using System.Collections.Concurrent;
using System.Reflection;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service.Workers;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ExcelControllerTests
{
    [Fact]
    public void WorkerRegistration_OlderRecordDefaultsObservationToNull()
    {
        var process = new WorkerProcess(123, 456);
        var registration = new WorkerRegistration("old-record", process, process, "generation", "build",
            new ExcelProcessIdentity(789, 101112), OwnsExcel: true);
        var json = JsonSerializer.Serialize(registration, ServiceProtocol.JsonOptions);
        using var document = JsonDocument.Parse(json);
        Assert.False(document.RootElement.TryGetProperty("lastExcelObservation", out _));

        var restored = JsonSerializer.Deserialize<WorkerRegistration>(json, ServiceProtocol.JsonOptions);

        Assert.NotNull(restored);
        Assert.Equal(registration, restored);
        Assert.Null(restored.LastExcelObservation);
    }

    [Fact]
    public async Task Controller_ExitedHistoryIsVisibleBeforeReleaseAndBoundedByExactIdentity()
    {
        using var controller = new ExcelController(launcher: Launcher);
        await StopNativeObservationPollingAsync(controller);
        var first = new ExcelProcessIdentity(1000, 2000);
        var releaseObserved = false;
        void OnRelease(ExcelProcessIdentity identity)
        {
            using var output = NativeObservationSessionList(controller);
            var exit = Assert.Single(output.RootElement.GetProperty("recentExcelExits").EnumerateArray());
            Assert.Equal(identity.ProcessId, exit.GetProperty("processId").GetInt32());
            Assert.Equal(identity.StartedAtUtcFileTime.ToString(System.Globalization.CultureInfo.InvariantCulture),
                exit.GetProperty("startedAtUtcFileTime").GetString());
            releaseObserved = true;
        }
        controller.ExcelProcessReleased += OnRelease;
        try
        {
            PublishNativeObservation(controller, first, ExcelNativeProcessState.Exited, 0, null);
        }
        finally
        {
            controller.ExcelProcessReleased -= OnRelease;
        }
        Assert.True(releaseObserved);
        for (var index = 1; index < 18; index++)
            PublishNativeObservation(controller, new(1000 + index, 2000 + index),
                ExcelNativeProcessState.Exited, index == 17 ? uint.MaxValue : (uint)index, null);
        // A duplicate cannot consume capacity or replace a known code with an unknown one.
        PublishNativeObservation(controller, new(1017, 2017), ExcelNativeProcessState.Exited, null, 5);
        // Reusing a PID with a different creation time is a separate process identity.
        PublishNativeObservation(controller, new(1017, 9999), ExcelNativeProcessState.Exited, 259, null);

        using var list = NativeObservationSessionList(controller);
        var root = list.RootElement;
        var exits = root.GetProperty("recentExcelExits").EnumerateArray().ToArray();
        Assert.True(root.GetProperty("success").GetBoolean());
        Assert.Empty(root.GetProperty("sessions").EnumerateArray());
        Assert.Equal(0, root.GetProperty("count").GetInt32());
        Assert.Equal(16, exits.Length);
        Assert.Equal(1003, exits[0].GetProperty("processId").GetInt32());
        Assert.Equal("2003", exits[0].GetProperty("startedAtUtcFileTime").GetString());
        Assert.Equal(uint.MaxValue, exits[^2].GetProperty("exitCode").GetUInt32());
        Assert.Equal("2017", exits[^2].GetProperty("startedAtUtcFileTime").GetString());
        Assert.Equal("9999", exits[^1].GetProperty("startedAtUtcFileTime").GetString());
        Assert.Equal(259u, exits[^1].GetProperty("exitCode").GetUInt32());
        Assert.All(exits, exit => Assert.Equal("Exited", exit.GetProperty("state").GetString()));
        Assert.All(exits, exit => Assert.Equal(JsonValueKind.Null, exit.GetProperty("nativeError").ValueKind));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Controller_UnknownExitCodeRemainsExplicitNullInCompleteAndPartialLists(bool partial)
    {
        using var controller = new ExcelController(launcher: Launcher);
        await StopNativeObservationPollingAsync(controller);
        const long creationTime = 134_381_234_567_890_123;
        var identity = new ExcelProcessIdentity(123, creationTime);
        PublishNativeObservation(controller, identity, ExcelNativeProcessState.Exited, null, 5);
        var current = WorkerProcess.Current;
        var registration = new WorkerRegistration($"observation-{Guid.NewGuid():N}", current, current,
            "generation", WorkerRegistry.Build, null);
        var worker = new ExcelWorkerClient(registration, current, "generation")
        {
            StatusError = "Test-only unavailable worker."
        };
        var workers = Assert.IsType<ConcurrentDictionary<string, ExcelWorkerClient>>(
            typeof(ExcelController).GetField("_workers", NativeObservationHidden)!.GetValue(controller));
        if (partial)
            Assert.True(workers.TryAdd(registration.PipeName, worker));
        try
        {
            var response = await controller.ProcessAsync(new ServiceRequest { Command = "session.list" });
            Assert.Equal(!partial, response.Success);
            if (partial)
                Assert.Equal("ServiceUnavailable", response.ErrorCategory);
            using var document = JsonDocument.Parse(partial ? response.PartialState! : response.Result!);
            var root = document.RootElement;
            Assert.Equal(!partial, root.GetProperty(partial ? "complete" : "success").GetBoolean());
            Assert.Empty(root.GetProperty("sessions").EnumerateArray());
            Assert.Equal(0, root.GetProperty("count").GetInt32());
            var exit = Assert.Single(root.GetProperty("recentExcelExits").EnumerateArray());
            Assert.Equal(123, exit.GetProperty("processId").GetInt32());
            Assert.Equal("134381234567890123", exit.GetProperty("startedAtUtcFileTime").GetString());
            Assert.Equal(creationTime, long.Parse(exit.GetProperty("startedAtUtcFileTime").GetString()!,
                System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal("Exited", exit.GetProperty("state").GetString());
            Assert.Equal(JsonValueKind.Null, exit.GetProperty("exitCode").ValueKind);
            Assert.Equal(5, exit.GetProperty("nativeError").GetInt32());
        }
        finally
        {
            workers.TryRemove(registration.PipeName, out _);
        }
    }

    [Theory]
    [InlineData(0, true)]
    [InlineData(0, false)]
    [InlineData(2, true)]
    public async Task Controller_NonterminalObservationsDoNotBecomeRecentExits(int stateValue, bool owned)
    {
        using var controller = new ExcelController(launcher: Launcher);
        await StopNativeObservationPollingAsync(controller);
        var current = WorkerProcess.Current;
        var identity = new ExcelProcessIdentity(123, 456);
        var state = (ExcelNativeProcessState)stateValue;
        var registration = new WorkerRegistration("observation", current, current, "generation", "build", identity,
            OwnsExcel: owned, ExcelPreserved: owned,
            LastExcelObservation: new(identity, state, null, state == ExcelNativeProcessState.Indeterminate ? 5 : null));
        PublishNativeRegistration(controller, registration);

        using var list = NativeObservationSessionList(controller);
        Assert.Empty(list.RootElement.GetProperty("recentExcelExits").EnumerateArray());
    }

    private const BindingFlags NativeObservationHidden = BindingFlags.Instance | BindingFlags.NonPublic;

    private static async Task StopNativeObservationPollingAsync(ExcelController controller)
    {
        var stop = Assert.IsType<CancellationTokenSource>(
            typeof(ExcelController).GetField("_stop", NativeObservationHidden)!.GetValue(controller));
        await stop.CancelAsync();
        await ((Task)typeof(ExcelController).GetField("_poll", NativeObservationHidden)!.GetValue(controller)!)
            .WaitAsync(TimeSpan.FromSeconds(5));
    }

    private static void PublishNativeObservation(ExcelController controller, ExcelProcessIdentity identity,
        ExcelNativeProcessState state, uint? exitCode, int? nativeError)
    {
        var current = WorkerProcess.Current;
        PublishNativeRegistration(controller, new WorkerRegistration("observation", current, current,
            "generation", "build", null, ReleasedExcel: identity,
            LastExcelObservation: new(identity, state, exitCode, nativeError)));
    }

    private static void PublishNativeRegistration(ExcelController controller, WorkerRegistration registration) =>
        typeof(ExcelController).GetMethod("PublishProcessIdentity", NativeObservationHidden)!
            .Invoke(controller, [registration]);

    private static JsonDocument NativeObservationSessionList(ExcelController controller)
    {
        var response = Assert.IsType<ServiceResponse>(typeof(ExcelController)
            .GetMethod("SessionList", NativeObservationHidden)!.Invoke(controller, null));
        Assert.True(response.Success);
        return JsonDocument.Parse(response.Result!);
    }
}
