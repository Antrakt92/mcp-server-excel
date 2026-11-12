using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "VBA")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
public sealed class VbaRunValidationCliTests
{
    [Fact]
    public async Task VbaRun_WhitespaceProcedureName_IsRejectedBeforeServiceCall()
    {
        var result = await CliProcessHelper.RunAsync(
            ["vba", "run", "--session", "unused-session", "--procedure-name", "   "],
            timeoutMs: 60000,
            diagnosticLabel: "vba run with whitespace procedure name");

        Assert.Equal(1, result.ExitCode);

        using var json = JsonDocument.Parse(result.Stdout);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains(
            "procedureName is required for run action",
            json.RootElement.GetProperty("error").GetProperty("message").GetString());
    }

    [Fact]
    public async Task VbaImportAndRun_MapFileAndProcedureArguments()
    {
        ServiceRequest? import = null;
        var importResult = await InProcessCliHelper.RunAsync(
            ["vba", "import", "--session", "session-1", "--module-name", "TransportProof", "--vba-code-file", "module.bas"],
            request =>
            {
                import = request;
                return new ServiceResponse { Success = true, Result = """{"success":true}""" };
            });
        ServiceRequest? run = null;
        var runResult = await InProcessCliHelper.RunAsync(
            ["vba", "run", "--session", "session-1", "--procedure-name", "TransportProof.WriteTransportProof"],
            request =>
            {
                run = request;
                return new ServiceResponse { Success = true, Result = """{"success":true}""" };
            });

        Assert.True(importResult.ExitCode == 0, importResult.Stdout + importResult.Stderr);
        Assert.True(runResult.ExitCode == 0, runResult.Stdout + runResult.Stderr);
        Assert.NotNull(import);
        Assert.Equal("vba.import", import.Command);
        Assert.Equal("session-1", import.SessionId);
        using (var args = JsonDocument.Parse(import.Args!))
        {
            Assert.Equal("TransportProof", args.RootElement.GetProperty("moduleName").GetString());
            Assert.Equal(Path.GetFullPath("module.bas"), args.RootElement.GetProperty("vbaCodeFile").GetString());
            Assert.False(args.RootElement.TryGetProperty("vbaCode", out _));
        }
        Assert.NotNull(run);
        Assert.Equal("vba.run", run.Command);
        using (var args = JsonDocument.Parse(run.Args!))
            Assert.Equal("TransportProof.WriteTransportProof", args.RootElement.GetProperty("procedureName").GetString());
    }

    [Fact]
    public async Task Batch_ValidatesFileInputAtItsTurnAndResolvesRelativeGridPath()
    {
        string fullPath = Path.Combine(Path.GetTempPath(), $"excelmcp-test-{Guid.NewGuid():N}.m");
        var requests = new List<ServiceRequest>();
        var input = JsonSerializer.Serialize(new object[]
        {
            new { command = "vba.run", sessionId = "target", args = new { procedureName = "Module1.WriteFixture" } },
            new { command = "powerquery.evaluate", sessionId = "target", args = new { mCodeFile = fullPath } },
            new { command = "range.set-values", sessionId = "target", args = new { sheetName = "Data", rangeAddress = "A1", valuesFile = "relative.json" } }
        });
        try
        {
            var result = await InProcessCliHelper.RunAsync(["batch"], request =>
            {
                requests.Add(request);
                if (request.Command == "vba.run")
                    File.WriteAllText(fullPath, "let Test = 1 in Test");
                if (request.Command == "powerquery.evaluate")
                    Assert.Equal("let Test = 1 in Test", File.ReadAllText(fullPath));
                return new ServiceResponse { Success = true, Result = """{"success":true}""" };
            }, input);

            Assert.Equal(0, result.ExitCode);
            Assert.Collection(requests,
                request => Assert.Equal("vba.run", request.Command),
                request => Assert.Equal("powerquery.evaluate", request.Command),
                request => Assert.Equal("range.set-values", request.Command));
            using var query = JsonDocument.Parse(requests[1].Args!);
            Assert.Equal(fullPath, query.RootElement.GetProperty("mCodeFile").GetString());
            using var range = JsonDocument.Parse(requests[2].Args!);
            Assert.Equal(Path.GetFullPath("relative.json"), range.RootElement.GetProperty("valuesFile").GetString());
        }
        finally
        {
            File.Delete(fullPath);
        }
    }
}
