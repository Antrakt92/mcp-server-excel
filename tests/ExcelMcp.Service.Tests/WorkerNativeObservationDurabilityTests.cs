using System.Reflection;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service.Workers;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "ServiceDaemon")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class WorkerNativeObservationDurabilityTests
{
    [Fact]
    public void RecordObservation_DurableDuplicateSurvivesWriteFailureButRicherStateStillRequiresPersistence()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"excelmcp-native-observation-{Guid.NewGuid():N}");
        try
        {
            var registry = new WorkerRegistry(directory);
            var initial = Registration(released: true);
            registry.Write(initial);
            using var host = new ExcelWorkerHost(initial, registry: registry);
            var unknown = new ExcelNativeProcessObservation(initial.ReleasedExcel!.Value,
                ExcelNativeProcessState.Exited, null, 5);
            Record(host, unknown);
            var durable = CurrentRegistration(host);
            var temporary = Path.Combine(directory, initial.PipeName + ".json.tmp");
            Directory.CreateDirectory(temporary);

            Record(host, unknown);

            Assert.Same(durable, CurrentRegistration(host));
            Assert.Equal(durable, registry.Read(initial.PipeName));
            var richer = unknown with { ExitCode = 0, NativeError = null };
            var failure = Assert.Throws<TargetInvocationException>(() => Record(host, richer));
            Assert.True(failure.InnerException is IOException or UnauthorizedAccessException);
            Assert.Same(durable, CurrentRegistration(host));
            Assert.Equal(durable, registry.Read(initial.PipeName));
            Directory.Delete(temporary);
            Record(host, richer);
            Assert.Equal(richer, CurrentRegistration(host).LastExcelObservation);
            Assert.Equal(CurrentRegistration(host), registry.Read(initial.PipeName));
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void RecordObservation_WriteFailureRetainsRegistrationAndCanRetry()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"excelmcp-native-observation-{Guid.NewGuid():N}");
        try
        {
            var registry = new WorkerRegistry(directory);
            var initial = Registration(released: false);
            registry.Write(initial);
            using var host = new ExcelWorkerHost(initial, registry: registry);
            var observation = new ExcelNativeProcessObservation(initial.Excel!.Value, ExcelNativeProcessState.Exited, 17, null);
            var temporary = Path.Combine(directory, initial.PipeName + ".json.tmp");
            // A directory at the staging-file path makes the actual registry write fail.
            Directory.CreateDirectory(temporary);

            var failure = Assert.Throws<TargetInvocationException>(() => Record(host, observation));

            Assert.True(failure.InnerException is IOException or UnauthorizedAccessException);
            Assert.Same(initial, CurrentRegistration(host));
            Assert.Equal(initial, registry.Read(initial.PipeName));
            Directory.Delete(temporary);
            Record(host, observation);
            var updated = CurrentRegistration(host);
            Assert.Equal(observation, updated.LastExcelObservation);
            Assert.Equal(initial.Excel, updated.Excel);
            Assert.True(updated.OwnsExcel);
            Assert.Equal(updated, registry.Read(initial.PipeName));
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void RecordObservation_ExactReleasedIdentityIsIdempotentWithoutRestoringOwnership()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"excelmcp-native-observation-{Guid.NewGuid():N}");
        try
        {
            var registry = new WorkerRegistry(directory);
            var initial = Registration(released: true);
            registry.Write(initial);
            using var host = new ExcelWorkerHost(initial, registry: registry);
            var observation = new ExcelNativeProcessObservation(initial.ReleasedExcel!.Value, ExcelNativeProcessState.Exited, 0, null);

            Record(host, observation);
            Record(host, observation);

            var updated = CurrentRegistration(host);
            Assert.Null(updated.Excel);
            Assert.False(updated.OwnsExcel);
            Assert.False(updated.ExcelPreserved);
            Assert.Equal(initial.ReleasedExcel, updated.ReleasedExcel);
            Assert.Equal(observation, updated.LastExcelObservation);
            Assert.Equal(updated, registry.Read(initial.PipeName));
            var unrelated = observation with { Identity = observation.Identity with { StartedAtUtcFileTime = 999 } };
            Assert.IsType<InvalidOperationException>(Assert.Throws<TargetInvocationException>(
                () => Record(host, unrelated)).InnerException);
            Assert.Same(updated, CurrentRegistration(host));
            Assert.Equal(updated, registry.Read(initial.PipeName));
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, recursive: true);
        }
    }

    private const BindingFlags Hidden = BindingFlags.Instance | BindingFlags.NonPublic;

    private static WorkerRegistration Registration(bool released)
    {
        var process = new WorkerProcess(123, 456);
        var identity = new ExcelProcessIdentity(789, 101112);
        return new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", process, process, "generation", "build",
            released ? null : identity, OwnsExcel: !released, ReleasedExcel: released ? identity : null);
    }

    private static WorkerRegistration CurrentRegistration(ExcelWorkerHost host) =>
        Assert.IsType<WorkerRegistration>(typeof(ExcelWorkerHost).GetField("_registration", Hidden)!.GetValue(host));

    private static void Record(ExcelWorkerHost host, ExcelNativeProcessObservation observation) =>
        typeof(ExcelWorkerHost).GetMethod("RecordObservation", Hidden)!.Invoke(host, [observation]);
}
