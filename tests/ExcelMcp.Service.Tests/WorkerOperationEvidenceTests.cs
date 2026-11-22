using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service.Workers;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("Sequential")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceDaemon")]
[Trait("Speed", "Medium")]
[Trait("RequiresExcel", "false")]
public sealed class WorkerOperationEvidenceTests
{
    [Theory]
    [InlineData("42")]
    [InlineData("""{"phase":42}""")]
    public void Registry_MalformedOptionalCheckpointDoesNotDropOwnership(string checkpoint)
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-malformed-checkpoint-");
        var registry = new WorkerRegistry(directory.FullName);
        var owner = WorkerProcess.Current;
        var record = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", owner, owner,
            "generation", WorkerRegistry.Build, new(1234, 5678), OwnsExcel: true);
        var json = JsonNode.Parse(ServiceProtocol.Serialize(record))!.AsObject();
        json["operationEvidence"] = JsonNode.Parse(checkpoint);
        var path = Path.Combine(directory.FullName, record.PipeName + ".json");
        try
        {
            File.WriteAllText(path, json.ToJsonString());
            Assert.Equal(record, registry.Read(record.PipeName));
            Assert.Equal(json.ToJsonString(), File.ReadAllText(path));
        }
        finally
        {
            registry.Remove(record.PipeName);
            directory.Delete();
        }
    }

    [Fact]
    public void Client_ObservesOriginalRegisteredReplyEvenAfterANewerPoll()
    {
        var process = WorkerProcess.Current;
        const string generation = "test-generation";
        var excel = new ExcelProcessIdentity(1234, 5678);
        var registration = new WorkerRegistration("test-pipe", process, process, generation, WorkerRegistry.Build, excel);
        var client = new ExcelWorkerClient(registration, process, generation);
        var earlier = new WorkerOperationEvidence(Guid.NewGuid().ToString("N"), Guid.NewGuid().ToString("N"),
            "vba.invoke", Guid.NewGuid().ToString("N"),
            excel, WorkerOperationEvidence.ResponsePhase);
        var later = earlier with { OperationId = Guid.NewGuid().ToString("N") };
        var empty = new ServiceResponse { Success = true, Result = """{"sessions":[]}""" };
        var newest = new WorkerSnapshot(registration with { OperationEvidence = later }, [], empty,
            null, null, new Dictionary<string, IReadOnlyList<string>>(), Revision: 2);
        client.RecordSnapshot(newest);
        WorkerOperationEvidence? observed = null;
        var effective = client.RecordSnapshot(newest with
        {
            Registration = registration with { OperationEvidence = earlier },
            Revision = 1
        }, source => observed = source.OperationEvidence);
        Assert.Equal(earlier, observed);
        Assert.Same(newest, effective);
        Assert.Equal(later, client.Registration.OperationEvidence);
    }

    [Theory]
    [InlineData("table.append")]
    [InlineData("session.save")]
    [InlineData("vba.invoke")]
    public void Checkpoint_StoresOnlyRecognizedFieldsFromTrackedRequestAndResponse(string command)
    {
        var id = Guid.NewGuid();
        var dispatchId = Guid.NewGuid();
        var sessionId = Guid.NewGuid().ToString("N");
        var excel = new ExcelProcessIdentity(1234, 5678);
        var call = new WorkerCall(dispatchId.ToString("N"), "generation", new ServiceRequest
        {
            OperationId = id.ToString("D"),
            Command = command,
            SessionId = sessionId,
            Args = """{"value":"PrivateWorkbookArgs"}"""
        });
        var admitted = Assert.IsType<WorkerOperationEvidence>(WorkerOperationEvidence.TryAdmit(call, excel));
        Assert.Equal(WorkerOperationEvidence.AdmittedPhase, admitted.Phase);
        Assert.Equal(excel, admitted.Excel);
        Assert.Equal(sessionId, admitted.SessionId);
        Assert.Equal(id.ToString("N"), admitted.OperationId);
        Assert.Equal(dispatchId.ToString("N"), admitted.DispatchId);
        Assert.Equal(command, admitted.Command);
        Assert.Null(admitted.ErrorCategory);

        var recorded = admitted.WithResponse(new ServiceResponse
        {
            ErrorCategory = "ExcelDisconnected",
            ExceptionType = "COMException",
            HResult = "0x800706BE",
            ErrorMessage = @"PrivateWorkbookMessage C:\fixture\book.xlsm",
            Result = """{"value":"PrivateWorkbookResult"}""",
            InnerError = "PrivateWorkbookInner",
            PartialState = """{"value":"PrivateWorkbookPartialState"}"""
        });

        Assert.Equal(WorkerOperationEvidence.ResponsePhase, recorded.Phase);
        Assert.Equal("ExcelDisconnected", recorded.ErrorCategory);
        Assert.Equal("COMException", recorded.ExceptionType);
        Assert.Equal("0x800706BE", recorded.HResult);
        string json = ServiceProtocol.Serialize(recorded);
        Assert.DoesNotContain("PrivateWorkbook", json, StringComparison.Ordinal);
        using var document = JsonDocument.Parse(json);
        Assert.Equal("0x800706BE", document.RootElement.GetProperty("hresult").GetString());
    }

    [Theory]
    [InlineData("no-id")]
    [InlineData("invalid-dispatch")]
    [InlineData("invalid-id-shape")]
    [InlineData("unsupported")]
    [InlineData("non-generated-session")]
    public void Checkpoint_RejectsUntrackedOrUnrecognizedIdentifiers(string scenario)
    {
        var id = Guid.NewGuid();
        string? requested = scenario switch
        {
            "no-id" => null,
            "invalid-id-shape" => id.ToString("B"),
            _ => id.ToString("N")
        };
        var call = new WorkerCall(scenario == "invalid-dispatch" ? Guid.Empty.ToString("N") : id.ToString("N"), "generation", new ServiceRequest
        {
            OperationId = requested,
            Command = scenario == "unsupported" ? "diag.echo" : "vba.invoke",
            SessionId = scenario == "non-generated-session" ? "private-path.xlsx" : Guid.NewGuid().ToString("N")
        });
        Assert.Null(WorkerOperationEvidence.TryAdmit(call, null));
    }

    [Fact]
    public void Registry_DropsMisidentifiedOptionalCheckpointWithoutDroppingWorkbookOwnership()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-evidence-mismatch-");
        var registry = new WorkerRegistry(directory.FullName);
        var owner = WorkerProcess.Current;
        var native = new ExcelProcessIdentity(1234, 5678);
        var record = new WorkerRegistration($"excelmcp-worker-{Guid.NewGuid():N}", owner, owner,
            "generation", WorkerRegistry.Build, native, OwnsExcel: true, OperationEvidence: new(
                Guid.NewGuid().ToString("N"), Guid.NewGuid().ToString("N"), "vba.invoke", Guid.NewGuid().ToString("N"),
                native with { StartedAtUtcFileTime = 9876 }, WorkerOperationEvidence.ResponsePhase, "ExcelDisconnected"));
        try
        {
            registry.Write(record);
            Assert.Equal(record with { OperationEvidence = null }, registry.Read(record.PipeName));
        }
        finally
        {
            registry.Remove(record.PipeName);
            directory.Delete();
        }
    }

    [Theory(Timeout = 15000)]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PrivateWorker_CheckpointWritesDoNotChangeTheWorkbookResponse(bool failWrite)
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-operation-checkpoint-");
        var registry = new WorkerRegistry(directory.FullName);
        var process = WorkerProcess.Current;
        var generation = Guid.NewGuid().ToString("N");
        var registration = new WorkerRegistration("excelmcp-worker-" + generation, process, process,
            generation, WorkerRegistry.Build, null);
        using var host = new ExcelWorkerHost(registration, registry: registry);
        var run = host.RunAsync();
        var worker = new ExcelWorkerClient(registration, process, generation);
        var operation = Guid.NewGuid();
        var sessionId = Guid.NewGuid().ToString("N");
        var tempPath = Path.Combine(directory.FullName, registration.PipeName + ".json.tmp");
        await worker.InitializeAsync();
        try
        {
            if (failWrite)
                Directory.CreateDirectory(tempPath);
            var result = await worker.ExecuteWithAdmissionAsync(new ServiceRequest
            {
                Command = "session.save",
                SessionId = sessionId,
                OperationId = operation.ToString("D")
            }, CancellationToken.None);
            Assert.True(result.Accepted);
            Assert.False(result.Response.Success);
            Assert.Equal("SessionNotFound", result.Response.ErrorCategory);

            var evidence = registry.Read(registration.PipeName)!.OperationEvidence;
            Assert.Equal(evidence, result.Snapshot.Registration.OperationEvidence);
            if (failWrite)
                Assert.Null(evidence);
            else
            {
                Assert.NotNull(evidence);
                Assert.Equal(WorkerOperationEvidence.ResponsePhase, evidence!.Phase);
                Assert.Equal(operation.ToString("N"), evidence.OperationId);
                Assert.Equal(sessionId, evidence.SessionId);
                Assert.Equal("session.save", evidence.Command);
                Assert.Equal("SessionNotFound", evidence.ErrorCategory);
            }
        }
        finally
        {
            if (failWrite)
                Directory.Delete(tempPath);
            var released = await worker.ReleaseAsync(CancellationToken.None);
            Assert.True(released.Success, released.ErrorMessage);
            await run.WaitAsync(TimeSpan.FromSeconds(3));
            registry.Remove(registration.PipeName);
            directory.Delete();
        }
    }
}
