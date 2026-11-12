using System.Collections.Concurrent;
using System.Reflection;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "SessionLifecycle")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class AuditSessionPolicyTests
{
    [Fact]
    public async Task ForbiddenFileInput_ReturnsAuditRefusalBeforeFileExistenceCheck()
    {
        using var service = new ExcelMcpService();
        const string session = "audit-fixture";
        var field = typeof(SessionManager).GetField("_auditSessions", BindingFlags.Instance | BindingFlags.NonPublic);
        var auditSessions = Assert.IsType<ConcurrentDictionary<string, byte>>(field?.GetValue(service.SessionManager));
        Assert.True(auditSessions.TryAdd(session, 0));
        var missingFile = Path.Combine(Path.GetTempPath(), $"missing-{Guid.NewGuid():N}.bas");

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "vba.import",
            SessionId = session,
            Args = JsonSerializer.Serialize(new { moduleName = "TestModule", vbaCodeFile = missingFile })
        });

        Assert.False(response.Success);
        Assert.Equal("Permissions", response.ErrorCategory);
        Assert.Contains("Audit session", response.ErrorMessage, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("range.get-values")]
    [InlineData("range.get-used-range")]
    [InlineData("range.get-used-range-info")]
    [InlineData("range.get-values-page")]
    [InlineData("rangeformat.get-format-summary")]
    [InlineData("chartconfig.get-text-format")]
    [InlineData("vba.view")]
    [InlineData("workbook.get-info")]
    [InlineData("calculationmode.get-settings")]
    [InlineData("session.notices")]
    public void InspectionCommands_RemainAvailable(string command) =>
        AuditSessionPolicy.EnsureAllowed(command, null);

    [Theory]
    [InlineData("range.set-values")]
    [InlineData("chartconfig.set-text-format")]
    [InlineData("pivottable.share-cache")]
    [InlineData("vba.run")]
    [InlineData("vba.invoke")]
    [InlineData("vba.unlock-project")]
    [InlineData("vba.verify-password")]
    [InlineData("vba.replace-module")]
    [InlineData("vba.update")]
    [InlineData("workbook.save-as")]
    [InlineData("session.save")]
    [InlineData("session.enable-editing")]
    [InlineData("session.enable-macros")]
    [InlineData("calculationmode.calculate")]
    [InlineData("unknown.command")]
    public void MutationsAndUnknownCommands_AreRefused(string command)
    {
        var error = Assert.Throws<Sbroenne.ExcelMcp.Core.Models.OperationFailureException>(() =>
            AuditSessionPolicy.EnsureAllowed(command, null));
        Assert.Contains(command, error.Message, StringComparison.Ordinal);
        Assert.Contains("Audit session", error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("{}")]
    [InlineData("{\"save\":false}")]
    public void ExplicitDiscardClose_IsAllowed(string? args) =>
        AuditSessionPolicy.EnsureAllowed("session.close", args);

    [Theory]
    [InlineData("{\"save\":true}")]
    [InlineData("{\"save\":null}")]
    [InlineData("{\"save\":\"false\"}")]
    public void SavingOrAmbiguousClose_IsRefused(string args) =>
        Assert.Throws<Sbroenne.ExcelMcp.Core.Models.OperationFailureException>(() =>
            AuditSessionPolicy.EnsureAllowed("session.close", args));

    [Fact]
    public void ReadOnlyParameter_IsAcceptedOnlyByOpen()
    {
        ExcelMcpService.ValidateSessionActionArguments("open", """{"filePath":"C:\\fixture.xlsx","readOnly":true}""");
        var error = Assert.Throws<ArgumentException>(() =>
            ExcelMcpService.ValidateSessionActionArguments("create", """{"readOnly":true}"""));
        Assert.Contains("readOnly", error.Message, StringComparison.Ordinal);
    }
}
