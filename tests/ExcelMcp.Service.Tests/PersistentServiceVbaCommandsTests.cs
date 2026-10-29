using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for VBA operations on desktop Excel with VBA project access enabled.
/// </summary>
internal interface IPersistentVbaCommands : IVbaCommands;

[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "VBA")]
public sealed class PersistentServiceVbaCommandsTests :
    PersistentServiceWorkbookTestBase,
    IClassFixture<PersistentServiceVbaFixture>
{
    private readonly IPersistentVbaCommands _scriptCommands;

    public PersistentServiceVbaCommandsTests(
        PersistentServiceVbaFixture fixture) :
        base(fixture)
    {
        _scriptCommands = fixture.CreateCommands<IPersistentVbaCommands>();
    }
    [Fact]
    public void ScriptCommands_Import_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        const string vbaCode = "Sub TestImport()\n    MsgBox \"Hello\"\nEnd Sub";

        // Act
        var batch = _fixture.BatchToken;
        Import(batch, "TestModule", vbaCode);

        // Assert - verify module exists via list
        var listResult = RequireSuccess(_scriptCommands.List(batch));
        Assert.Single(listResult.Scripts, s => s.Name == "TestModule");
        var code = RequireSuccess(_scriptCommands.View(batch, "TestModule"));
        AssertStoredCode(vbaCode, code.Code);
    }
    [Fact]
    public void ScriptCommands_Run_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        // Import a test macro first
        string vbaCode = @"Sub TestProcedure()
    ThisWorkbook.Sheets(1).Range(""A1"").Value = ""macro-ran""
End Sub";

        var batch = _fixture.BatchToken;
        Import(batch, "TestModule", vbaCode);

        // Act - Run the macro
        RequireSuccess(_scriptCommands.Run(batch, "TestModule.TestProcedure", null));

        // Assert - No exception thrown; to be thorough, ensure module still exists
        var listResult = _scriptCommands.List(batch);
        Assert.True(listResult.Success, listResult.ErrorMessage);
        Assert.Contains(listResult.Scripts, s => s.Name == "TestModule");
        var values = _commands.GetValues(batch, "Sheet1", "A1");
        Assert.True(values.Success, values.ErrorMessage);
        Assert.Equal("macro-ran", Assert.Single(Assert.Single(values.Values)));
    }

    [Fact]
    public async Task ScriptCommands_Run_OnReopenedMacroWorkbook_AfterList_WorksCorrectly()
    {
        // Arrange
        var rangeCommands = _commands;
        string vbaCode = @"Sub WriteMarker()
    ThisWorkbook.Sheets(1).Range(""A1"").Value = ""reopened-run-ok""
End Sub";

        var batch = _fixture.BatchToken;
        Import(batch, "ReopenTestModule", vbaCode);
        await _fixture.SaveAndReopenAsync();
        _fixture.Send("session.enable-macros", new { approved = true, allowReopen = true });

        // Act - reopen the existing .xlsm, prove VBA project access still works, then run
        var reopenedBatch = _fixture.BatchToken;
        var listResult = _scriptCommands.List(reopenedBatch);
        Assert.True(listResult.Success, $"List should succeed after reopen. Error: {listResult.ErrorMessage}");
        Assert.Contains(listResult.Scripts, s => s.Name == "ReopenTestModule");

        RequireSuccess(_scriptCommands.Run(reopenedBatch, "ReopenTestModule.WriteMarker", null));

        // Assert - macro execution against the reopened workbook should have real side effects
        var cellResult = rangeCommands.GetValues(reopenedBatch, "Sheet1", "A1");
        Assert.True(cellResult.Success, $"GetValues should succeed after reopened run. Error: {cellResult.ErrorMessage}");
        Assert.NotNull(cellResult.Values);
        Assert.Single(cellResult.Values);
        Assert.Equal("reopened-run-ok", cellResult.Values[0][0]?.ToString());
    }

    [Fact]
    public async Task ScriptCommands_Update_OnReopenedMacroWorkbook_ThenRun_UsesUpdatedCode()
    {
        // Arrange
        var rangeCommands = _commands;
        const string moduleName = "ReopenUpdateModule";
        string initialCode = @"Sub WriteMarker()
    ThisWorkbook.Sheets(1).Range(""A1"").Value = ""original-run""
End Sub";
        string updatedCode = @"Sub WriteMarker()
    ThisWorkbook.Sheets(1).Range(""A1"").Value = ""updated-run-ok""
End Sub";

        var batch = _fixture.BatchToken;
        Import(batch, moduleName, initialCode);
        await _fixture.SaveAndReopenAsync();
        _fixture.Send("session.enable-macros", new { approved = true, allowReopen = true });

        // Act
        var reopenedBatch = _fixture.BatchToken;
        RequireSuccess(_scriptCommands.Update(reopenedBatch, moduleName, updatedCode));
        var viewResult = _scriptCommands.View(reopenedBatch, moduleName);
        RequireSuccess(_scriptCommands.Run(reopenedBatch, $"{moduleName}.WriteMarker", null));

        // Assert
        Assert.True(viewResult.Success, $"View should succeed after reopened update. Error: {viewResult.ErrorMessage}");
        Assert.Contains("updated-run-ok", viewResult.Code);
        Assert.DoesNotContain("original-run", viewResult.Code);
        AssertStoredCode(updatedCode, viewResult.Code);

        var cellResult = rangeCommands.GetValues(reopenedBatch, "Sheet1", "A1");
        Assert.True(cellResult.Success, $"GetValues should succeed after reopened update run. Error: {cellResult.ErrorMessage}");
        Assert.NotNull(cellResult.Values);
        Assert.Single(cellResult.Values);
        Assert.Equal("updated-run-ok", cellResult.Values[0][0]?.ToString());
    }

    [Fact]
    public async Task ScriptCommands_DeleteThenImport_OnReopenedMacroWorkbook_ThenRun_UsesReplacementModule()
    {
        // Arrange
        var rangeCommands = _commands;
        const string moduleName = "ReopenReplaceModule";
        string originalCode = @"Sub WriteMarker()
    ThisWorkbook.Sheets(1).Range(""A1"").Value = ""original-module""
End Sub";
        string replacementCode = @"Sub WriteMarker()
    ThisWorkbook.Sheets(1).Range(""A1"").Value = ""replacement-run-ok""
End Sub";

        var batch = _fixture.BatchToken;
        Import(batch, moduleName, originalCode);
        await _fixture.SaveAndReopenAsync();
        _fixture.Send("session.enable-macros", new { approved = true, allowReopen = true });

        // Act
        var reopenedBatch = _fixture.BatchToken;
        Delete(reopenedBatch, moduleName);
        var afterDeleteList = _scriptCommands.List(reopenedBatch);
        Import(reopenedBatch, moduleName, replacementCode);
        var replacementView = _scriptCommands.View(reopenedBatch, moduleName);
        RequireSuccess(_scriptCommands.Run(reopenedBatch, $"{moduleName}.WriteMarker", null));

        // Assert
        Assert.True(afterDeleteList.Success, $"List should succeed after delete. Error: {afterDeleteList.ErrorMessage}");
        Assert.DoesNotContain(afterDeleteList.Scripts, s => s.Name == moduleName);

        Assert.True(replacementView.Success, $"View should succeed after replacement import. Error: {replacementView.ErrorMessage}");
        Assert.Contains("replacement-run-ok", replacementView.Code);
        Assert.DoesNotContain("original-module", replacementView.Code);

        var cellResult = rangeCommands.GetValues(reopenedBatch, "Sheet1", "A1");
        Assert.True(cellResult.Success, $"GetValues should succeed after replacement run. Error: {cellResult.ErrorMessage}");
        Assert.NotNull(cellResult.Values);
        Assert.Single(cellResult.Values);
        Assert.Equal("replacement-run-ok", cellResult.Values[0][0]?.ToString());
    }

    [Fact]
    public void ScriptCommands_Delete_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        // Import a module first
        string vbaCode = "Sub TestCode()\nEnd Sub";

        var batch = _fixture.BatchToken;
        Import(batch, "TestModule", vbaCode);
        Import(batch, "RetainedModule", "Sub Retained()\nEnd Sub");
        var retained = RequireSuccess(_scriptCommands.View(batch, "RetainedModule")).Code;

        // Act - Delete the module
        Delete(batch, "TestModule");

        // Verify module is gone
        var listResult = RequireSuccess(_scriptCommands.List(batch));
        Assert.DoesNotContain(listResult.Scripts, s => s.Name == "TestModule");
        Assert.Single(listResult.Scripts, s => s.Name == "RetainedModule");
        Assert.Equal(retained, RequireSuccess(_scriptCommands.View(batch, "RetainedModule")).Code);
    }
    [Fact]
    public void ScriptCommands_Update_WithTrustEnabled_WorksCorrectly()
    {
        // Arrange
        // Import initial module
        string initialCode = "Sub OriginalCode()\nEnd Sub";

        var batch = _fixture.BatchToken;
        Import(batch, "UpdateTestModule", initialCode);

        // Prepare updated code
        string updatedCode = "Sub UpdatedCode()\n    MsgBox \"Updated\"\nEnd Sub";

        // Act - Update the module with new code
        RequireSuccess(_scriptCommands.Update(batch, "UpdateTestModule", updatedCode));

        // Verify the code was updated
        var viewResult = _scriptCommands.View(batch, "UpdateTestModule");
        Assert.True(viewResult.Success);
        Assert.Contains("UpdatedCode", viewResult.Code);
        Assert.Contains("Updated", viewResult.Code);
        Assert.DoesNotContain("OriginalCode", viewResult.Code);
        AssertStoredCode(updatedCode, viewResult.Code);
    }

    [Fact]
    public void ScriptCommands_ReadAndReplaceProcedure_PreservesOtherModuleSource()
    {
        const string moduleName = "ProcedureEditModule";
        const string initialCode = """
            Option Explicit

            Public Function Target( _
                ByVal value As String) As String
                Target = value
            End Function

            Private Sub Retained()
            End Sub
            """;
        const string replacementCode = """
            Public Function Target( _
                ByVal value As String) As String
                Target = "changed-" & value
            End Function
            """;

        var batch = _fixture.BatchToken;
        Import(batch, moduleName, initialCode);

        var listed = RequireSuccess(_scriptCommands.List(batch));
        var info = Assert.Single(listed.Scripts, script => script.Name == moduleName);
        var targetInfo = Assert.Single(info.ProcedureDetails, procedure => procedure.Name == "Target");
        Assert.Equal("Function", targetInfo.Kind);
        Assert.True(targetInfo.StartLine > 1);
        Assert.True(targetInfo.LineCount >= 4);
        Assert.Contains("Target", info.Procedures);

        var procedureRead = RequireSuccess(_scriptCommands.Read(
            batch,
            moduleName,
            "Target",
            "Function",
            null,
            null));
        Assert.Equal(targetInfo.StartLine, procedureRead.StartLine);
        Assert.Equal(targetInfo.LineCount, procedureRead.TotalLineCount);
        Assert.False(procedureRead.HasMore);
        Assert.NotNull(procedureRead.SourceHash);
        Assert.Contains("ByVal value As String", procedureRead.Code);

        var lineRead = RequireSuccess(_scriptCommands.Read(
            batch,
            moduleName,
            null,
            null,
            1,
            2));
        Assert.Equal(1, lineRead.StartLine);
        Assert.Equal(2, lineRead.ReturnedLineCount);
        Assert.Contains("Option Explicit", lineRead.Code);

        var replacement = RequireSuccess(_scriptCommands.ReplaceProcedure(
            batch,
            moduleName,
            "Target",
            "Function",
            procedureRead.SourceHash!,
            replacementCode));
        Assert.Equal("Target", replacement.ProcedureName);
        Assert.Equal("Function", replacement.ProcedureKind);
        Assert.NotEqual(procedureRead.SourceHash, replacement.SourceHash);
        Assert.Contains("does not confirm", replacement.Message, StringComparison.OrdinalIgnoreCase);

        var fullModule = RequireSuccess(_scriptCommands.View(batch, moduleName));
        Assert.Contains("Option Explicit", fullModule.Code);
        Assert.Contains("Private Sub Retained()", fullModule.Code);
        Assert.Contains("changed-", fullModule.Code);
        Assert.DoesNotContain("Target = value", fullModule.Code);

        var updatedRead = RequireSuccess(_scriptCommands.Read(
            batch,
            moduleName,
            "Target",
            "Function",
            null,
            null));
        Assert.Equal(replacement.SourceHash, updatedRead.SourceHash);
    }

    [Fact]
    public void ScriptCommands_ReplaceProcedure_PreservesLeadingCommentsAndTrailingWhitespace()
    {
        const string moduleName = "CommentedProcedureModule";
        const string original = "Option Explicit\n\n' Target documentation\nPublic Sub Target()\n    Debug.Print \"old\"\nEnd Sub\n\n' Keep this note\nPrivate Sub Retained()\nEnd Sub\n\n";
        const string replacement = "Public Sub Target()\n    Debug.Print \"new\"\nEnd Sub";
        var batch = _fixture.BatchToken;
        Import(batch, moduleName, original);
        var before = RequireSuccess(_scriptCommands.View(batch, moduleName)).Code;
        var read = RequireSuccess(_scriptCommands.Read(batch, moduleName, "Target", "Sub", null, null));

        RequireSuccess(_scriptCommands.ReplaceProcedure(
            batch, moduleName, "Target", "Sub", read.SourceHash!, replacement));

        var after = RequireSuccess(_scriptCommands.View(batch, moduleName)).Code;
        Assert.Equal(before.Replace("\"old\"", "\"new\"", StringComparison.Ordinal), after);

        var lastRead = RequireSuccess(_scriptCommands.Read(batch, moduleName, "Retained", "Sub", null, null));
        RequireSuccess(_scriptCommands.ReplaceProcedure(
            batch, moduleName, "Retained", "Sub", lastRead.SourceHash!,
            "Private Sub Retained()\n    Debug.Print \"last\"\nEnd Sub"));
        Assert.Equal(
            after.Replace("Private Sub Retained()\r\nEnd Sub",
                "Private Sub Retained()\r\n    Debug.Print \"last\"\r\nEnd Sub", StringComparison.Ordinal),
            RequireSuccess(_scriptCommands.View(batch, moduleName)).Code);
    }

    [Fact]
    public void ScriptCommands_ReplaceProcedure_AcceptsLeadingBlankReplacementLines()
    {
        const string moduleName = "BlankPrefixProcedureModule";
        const string oldBody = "Public Sub Target()\r\n    Debug.Print \"old\"\r\nEnd Sub";
        const string replacement = "\r\n\r\nPublic Sub Target()\r\n    Debug.Print \"new\"\r\nEnd Sub";
        const string original = "Option Explicit\r\n\r\n' Existing target documentation\r\n" + oldBody
            + "\r\n\r\n' Existing retained documentation\r\nPrivate Sub Retained()\r\nEnd Sub";
        var batch = _fixture.BatchToken;
        Import(batch, moduleName, original);
        var before = RequireSuccess(_scriptCommands.View(batch, moduleName));
        var target = RequireSuccess(_scriptCommands.Read(batch, moduleName, "Target", "Sub", null, null));
        Assert.Contains(oldBody, before.Code, StringComparison.Ordinal);
        string expected = before.Code.Replace(oldBody, replacement, StringComparison.Ordinal);

        var replaced = RequireSuccess(_scriptCommands.ReplaceProcedure(
            batch, moduleName, "Target", "Sub", target.SourceHash!, replacement));

        var after = RequireSuccess(_scriptCommands.View(batch, moduleName));
        Assert.Equal(expected, after.Code);
        Assert.Equal(VbaModuleSource.ComputeHash(expected), after.SourceHash);
        var observed = RequireSuccess(_scriptCommands.Read(batch, moduleName, "Target", "Sub", null, null));
        Assert.False(observed.HasMore);
        Assert.Equal(VbaProcedureSource.ComputeHash(observed.Code), replaced.SourceHash);
        Assert.Equal(observed.SourceHash, replaced.SourceHash);
        Assert.Equal(observed.StartLine, replaced.StartLine);
        Assert.Equal(observed.TotalLineCount, replaced.LineCount);
    }

    [Fact]
    public async Task ScriptCommands_ReplaceProcedure_RejectsStaleSourceWithoutChangingIt()
    {
        const string moduleName = "StaleProcedureModule";
        const string original = "Public Sub UpdateMe()\nEnd Sub";
        const string concurrentUpdate = "Public Sub UpdateMe()\n    Debug.Print \"newer\"\nEnd Sub";
        const string attemptedUpdate = "Public Sub UpdateMe()\n    Debug.Print \"stale\"\nEnd Sub";

        var batch = _fixture.BatchToken;
        Import(batch, moduleName, original);
        var oldRead = RequireSuccess(_scriptCommands.Read(
            batch,
            moduleName,
            "UpdateMe",
            "Sub",
            null,
            null));
        RequireSuccess(_scriptCommands.Update(batch, moduleName, concurrentUpdate));

        var response = await _fixture.SendForFailureAsync(
            "vba.replace-procedure",
            new
            {
                moduleName,
                procedureName = "UpdateMe",
                procedureKind = "Sub",
                expectedSourceHash = oldRead.SourceHash,
                vbaCode = attemptedUpdate
            });
        Assert.Equal(OperationFailureCategory.Conflict.ToString(), response.ErrorCategory);
        AssertStoredCode(
            concurrentUpdate,
            RequireSuccess(_scriptCommands.Read(
                batch,
                moduleName,
                "UpdateMe",
                "Sub",
                null,
                null)).Code);
    }

    [Theory]
    [InlineData("Public Sub KeepMe()\nEnd Sub\nPublic Sub Extra()\nEnd Sub")]
    [InlineData("Public Sub KeepMe()\nEnd Sub: Public Sub Extra()\nEnd Sub")]
    public async Task ScriptCommands_ReplaceProcedure_RejectsExtraProcedureWithoutChangingTarget(string invalidReplacement)
    {
        const string moduleName = "InvalidProcedureModule";
        const string original = "Public Sub KeepMe()\nEnd Sub";

        var batch = _fixture.BatchToken;
        Import(batch, moduleName, original);
        var read = RequireSuccess(_scriptCommands.Read(batch, moduleName, "KeepMe", "Sub", null, null));

        var response = await _fixture.SendForFailureAsync(
            "vba.replace-procedure",
            new
            {
                moduleName,
                procedureName = "KeepMe",
                procedureKind = "Sub",
                expectedSourceHash = read.SourceHash,
                vbaCode = invalidReplacement
            });
        Assert.Equal(OperationFailureCategory.InvalidInput.ToString(), response.ErrorCategory);
        AssertStoredCode(
            original,
            RequireSuccess(_scriptCommands.Read(batch, moduleName, "KeepMe", "Sub", null, null)).Code);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task ScriptCommands_ReplaceProcedure_RejectsUnsupportedPhysicalSourceWithoutChangingHash(bool nul)
    {
        string moduleName = "PhysicalGuard" + (nul ? "Nul" : "Long");
        const string original = "Public Sub KeepMe()\nDebug.Print 1\nEnd Sub\nPublic Sub Retained()\nEnd Sub";
        var batch = _fixture.BatchToken;
        Import(batch, moduleName, original);
        var before = RequireSuccess(_scriptCommands.View(batch, moduleName));
        var procedure = RequireSuccess(_scriptCommands.Read(batch, moduleName, "KeepMe", "Sub", null, null));
        string body = nul ? "'bad\0value" : "'" + new string('x', 1023);
        var response = await _fixture.SendForFailureAsync("vba.replace-procedure", new
        {
            moduleName,
            procedureName = "KeepMe",
            procedureKind = "Sub",
            expectedSourceHash = procedure.SourceHash,
            vbaCode = "Public Sub KeepMe()\n" + body + "\nEnd Sub"
        });
        Assert.Equal(OperationFailureCategory.InvalidInput.ToString(), response.ErrorCategory);
        var after = RequireSuccess(_scriptCommands.View(batch, moduleName));
        Assert.Equal(before.SourceHash, after.SourceHash);
        Assert.Equal(before.Code, after.Code);
    }

    [Fact]
    public void ScriptCommands_ReadLongProcedure_ReturnsContinuationDetails()
    {
        const string moduleName = "LongProcedureModule";
        string body = string.Join(
            Environment.NewLine,
            Enumerable.Range(1, 510).Select(line => $"    Debug.Print {line}"));
        string code = $"Public Sub LongOne(){Environment.NewLine}{body}{Environment.NewLine}End Sub";

        var batch = _fixture.BatchToken;
        Import(batch, moduleName, code);

        var firstPart = RequireSuccess(_scriptCommands.Read(batch, moduleName, "LongOne", "Sub", null, null));

        Assert.Equal(500, firstPart.ReturnedLineCount);
        Assert.True(firstPart.HasMore);
        Assert.Equal(firstPart.StartLine + 500, firstPart.NextStartLine);

        var secondPart = RequireSuccess(_scriptCommands.Read(
            batch,
            moduleName,
            null,
            null,
            firstPart.NextStartLine,
            500));

        Assert.False(secondPart.HasMore);
        Assert.True(secondPart.ReturnedLineCount > 0);
        Assert.Contains("End Sub", secondPart.Code);
    }

    [Fact]
    public void ScriptCommands_Search_ReturnsLimitedLocatedMatchesAndHonorsOptions()
    {
        const string moduleName = "SearchModule";
        var batch = _fixture.BatchToken;
        Import(batch, moduleName,
            "Public Sub Target()\n    Debug.Print \"needle needle needles\"\n    Debug.Print \"NEEDLE\"\nEnd Sub");
        Import(batch, "OtherSearchModule", "Public Sub Other()\n    Debug.Print \"needle\"\nEnd Sub");
        var before = RequireSuccess(_scriptCommands.View(batch, moduleName)).Code;

        var response = _fixture.Send("vba.search", new { searchText = "needle", maxMatches = 2 });
        using var limited = JsonDocument.Parse(response.Result!);
        Assert.True(limited.RootElement.GetProperty("hasMore").GetBoolean());
        Assert.Equal(2, limited.RootElement.GetProperty("matches").GetArrayLength());
        Assert.All(limited.RootElement.GetProperty("matches").EnumerateArray(), match =>
        {
            Assert.Contains(match.GetProperty("moduleName").GetString(), new[] { moduleName, "OtherSearchModule" });
            Assert.Equal(2, match.GetProperty("line").GetInt32());
            Assert.True(match.GetProperty("column").GetInt32() > 1);
            Assert.Contains("needle", match.GetProperty("excerpt").GetString());
        });

        response = _fixture.Send("vba.search",
            new { searchText = "needle", moduleName, wholeWord = true, matchCase = true });
        using var exact = JsonDocument.Parse(response.Result!);
        var matches = exact.RootElement.GetProperty("matches").EnumerateArray().ToArray();
        Assert.Equal(2, matches.Length);
        Assert.False(exact.RootElement.GetProperty("hasMore").GetBoolean());
        Assert.All(matches, match => Assert.Equal(moduleName, match.GetProperty("moduleName").GetString()));
        Assert.Equal(18, matches[0].GetProperty("column").GetInt32());
        Assert.Equal(25, matches[1].GetProperty("column").GetInt32());

        response = _fixture.Send("vba.search", new { searchText = "not-present", moduleName });
        using var empty = JsonDocument.Parse(response.Result!);
        Assert.Empty(empty.RootElement.GetProperty("matches").EnumerateArray());
        Assert.False(empty.RootElement.GetProperty("hasMore").GetBoolean());
        Assert.Equal(before, RequireSuccess(_scriptCommands.View(batch, moduleName)).Code);
    }

    [Theory]
    [InlineData("", 50)]
    [InlineData("needle", 0)]
    [InlineData("needle", 101)]
    [InlineData("two\nlines", 50)]
    public async Task ScriptCommands_Search_RejectsInvalidLimitsAndText(string searchText, int maxMatches)
    {
        var response = await _fixture.SendForFailureAsync("vba.search", new { searchText, maxMatches });
        Assert.Equal(OperationFailureCategory.InvalidInput.ToString(), response.ErrorCategory);
    }

    [Fact]
    public async Task ScriptCommands_ReplacePropertyGet_PreservesLetAccessorAndComments()
    {
        const string moduleName = "PropertyProcedureModule";
        const string code = "Option Explicit\nPrivate mValue As String\n\n' Getter documentation\nPublic Property Get Value() As String\n    Value = mValue\nEnd Property\n\n' Setter documentation\nPublic Property Let Value(ByVal newValue As String)\n    mValue = newValue\nEnd Property";
        var batch = _fixture.BatchToken;
        Import(batch, moduleName, code);
        var before = RequireSuccess(_scriptCommands.View(batch, moduleName)).Code;
        var read = RequireSuccess(_scriptCommands.Read(batch, moduleName, "Value", "Property Get", null, null));
        var ambiguous = await _fixture.SendForFailureAsync("vba.read", new { moduleName, procedureName = "Value" });
        Assert.Equal(OperationFailureCategory.InvalidInput.ToString(), ambiguous.ErrorCategory);

        RequireSuccess(_scriptCommands.ReplaceProcedure(batch, moduleName, "Value", "Property Get", read.SourceHash!,
            "Public Property Get Value() As String\n    Value = \"changed\"\nEnd Property"));
        Assert.Equal(before.Replace("Value = mValue", "Value = \"changed\"", StringComparison.Ordinal),
            RequireSuccess(_scriptCommands.View(batch, moduleName)).Code);
        var details = Assert.Single(RequireSuccess(_scriptCommands.List(batch)).Scripts, script => script.Name == moduleName)
            .ProcedureDetails;
        Assert.Contains(details, procedure => procedure.Kind == "Property Get" && procedure.BodyStartLine > procedure.StartLine);
        Assert.Contains(details, procedure => procedure.Kind == "Property Let");
    }

    [Fact]
    public async Task ScriptCommands_Search_BoundsExcerptsAndReportsMissingModule()
    {
        const string moduleName = "LongSearchModule";
        var batch = _fixture.BatchToken;
        Import(batch, moduleName,
            "Public Sub Target()\n    Debug.Print \"" + new string('x', 250) + "needle" + new string('y', 250) + "\"\nEnd Sub");
        var response = _fixture.Send("vba.search", new { searchText = "needle", moduleName, maxMatches = 1 });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.False(result.RootElement.GetProperty("hasMore").GetBoolean());
        var match = Assert.Single(result.RootElement.GetProperty("matches").EnumerateArray());
        Assert.Equal(200, match.GetProperty("excerpt").GetString()!.Length);
        Assert.Contains("needle", match.GetProperty("excerpt").GetString());
        Assert.Equal(268, match.GetProperty("column").GetInt32());
        var missing = await _fixture.SendForFailureAsync("vba.search",
            new { searchText = "needle", moduleName = "NoSuchModule" });
        Assert.Equal(OperationFailureCategory.NotFound.ToString(), missing.ErrorCategory);
    }

    [Fact]
    public void ScriptCommands_References_ReportsExcelAndVbaLibrariesWithoutChangingCode()
    {
        var response = _fixture.Send("vba.references", new { });
        using var result = JsonDocument.Parse(response.Result!);
        var references = result.RootElement.GetProperty("references").EnumerateArray().ToArray();
        Assert.False(result.RootElement.GetProperty("hasBrokenReferences").GetBoolean());
        Assert.Contains(references, reference => reference.GetProperty("name").GetString() == "Excel");
        Assert.Contains(references, reference => reference.GetProperty("name").GetString() == "VBA");
        Assert.All(references, reference =>
        {
            Assert.False(reference.GetProperty("isBroken").GetBoolean());
            Assert.True(Guid.TryParse(reference.GetProperty("libraryId").GetString(), out _));
            Assert.True(reference.GetProperty("major").GetInt32() >= 0);
            Assert.True(reference.GetProperty("minor").GetInt32() >= 0);
        });
    }

    [Fact]
    public void ScriptCommands_Status_ReportsAccessibleUnlockedIdleProject()
    {
        var response = _fixture.Send("vba.status", new { });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("projectAccess").GetBoolean());
        Assert.Equal("None", result.RootElement.GetProperty("protection").GetString());
        Assert.Equal("Design", result.RootElement.GetProperty("mode").GetString());
        Assert.False(string.IsNullOrWhiteSpace(result.RootElement.GetProperty("projectName").GetString()));
        Assert.False(result.RootElement.TryGetProperty("accessMessage", out _));
    }

    [Fact]
    public void ScriptCommands_Run_WithParameters_PassesArgumentsCorrectly()
    {
        // Arrange
        // Import a macro that writes a parameter to a cell for verification
        string vbaCode = @"Sub TestWithParam(value As String)
    ThisWorkbook.Sheets(1).Range(""A1"").Value = value
End Sub";

        var batch = _fixture.BatchToken;
        Import(batch, "ParamTest", vbaCode);

        // Act - Run with parameter
        RequireSuccess(_scriptCommands.Run(batch, "ParamTest.TestWithParam", null, "HelloWorld"));

        // Assert - Verify the macro wrote the value
        var rangeCommands = _commands;
        var result = rangeCommands.GetValues(batch, "Sheet1", "A1");
        Assert.True(result.Success, $"GetValues should succeed. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Values);
        Assert.Single(result.Values);
        Assert.Equal("HelloWorld", result.Values[0][0]?.ToString());
    }

    [Fact]
    public void ScriptCommands_Run_WithThirtyParameters_PreservesPositionAndRestoresSecurity()
    {
        var parameters = Enumerable.Range(1, 30).Select(index => $"value-{index}").ToArray();
        string declaration = string.Join(", ", Enumerable.Range(1, 30).Select(index => $"ByVal p{index} As String"));
        string statements = string.Join("\n", Enumerable.Range(1, 30).Select(index =>
            $"    ThisWorkbook.Worksheets(1).Cells(1, {index}).Value = p{index}"));
        Import(_fixture.BatchToken, "ThirtyParameters", $"Public Sub Store({declaration})\n{statements}\nEnd Sub");
        var before = _fixture.ExecuteRawVerification((ctx, _) => ctx.App.AutomationSecurity);

        RequireSuccess(_scriptCommands.Run(_fixture.BatchToken, "ThirtyParameters.Store", null, parameters));

        var values = RequireSuccess(_commands.GetValues(_fixture.BatchToken, "Sheet1", "A1:AD1"));
        Assert.Equal(parameters, Assert.Single(values.Values).Select(value => value?.ToString()));
        var after = _fixture.ExecuteRawVerification((ctx, _) => ctx.App.AutomationSecurity);
        Assert.Equal(before, after);
    }

    [Fact]
    public void ScriptCommands_Run_WithMultipleParameters_PassesAllArguments()
    {
        // Arrange
        // Import a macro that writes two parameters to separate cells
        string vbaCode = @"Sub TestMultiParam(val1 As String, val2 As String)
    ThisWorkbook.Sheets(1).Range(""A1"").Value = val1
    ThisWorkbook.Sheets(1).Range(""B1"").Value = val2
End Sub";

        var batch = _fixture.BatchToken;
        Import(batch, "MultiParamTest", vbaCode);

        // Act - Run with two parameters
        RequireSuccess(_scriptCommands.Run(batch, "MultiParamTest.TestMultiParam", null, "First", "Second"));

        // Assert - Verify both parameters were passed correctly
        var rangeCommands = _commands;
        var result = rangeCommands.GetValues(batch, "Sheet1", "A1:B1");
        Assert.True(result.Success, $"GetValues should succeed. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Values);
        Assert.Single(result.Values); // one row
        Assert.Equal("First", result.Values[0][0]?.ToString());
        Assert.Equal("Second", result.Values[0][1]?.ToString());
    }

    [Fact]
    public async Task ScriptCommands_ReplaceModule_VerifiesSourceAndRejectsStaleHash()
    {
        const string moduleName = "GuardedWholeModule";
        Import(_fixture.BatchToken, moduleName, "Public Function Label() As String\nLabel = \"before\"\nEnd Function");
        var before = RequireSuccess(_scriptCommands.View(_fixture.BatchToken, moduleName));
        const string source = "Attribute VB_Name = \"GuardedWholeModule\"\r\r\nPublic Function Label() As String\r\r\nLabel = \"after\"\r\r\nEnd Function";
        var replaced = RequireSuccess(_scriptCommands.ReplaceModule(_fixture.BatchToken, moduleName, before.SourceHash, source));
        var after = RequireSuccess(_scriptCommands.View(_fixture.BatchToken, moduleName));
        Assert.Equal(after.SourceHash, replaced.SourceHash);
        Assert.NotEqual(before.SourceHash, after.SourceHash);
        Assert.DoesNotContain("Attribute VB_Name", after.Code);
        Assert.Contains("\"after\"", after.Code);
        Assert.Equal(after.LineCount, replaced.LineCount);
        var stale = await _fixture.SendForFailureAsync("vba.replace-module", new
        {
            moduleName,
            expectedSourceHash = before.SourceHash,
            vbaCode = "Public Sub Stale()\nEnd Sub"
        });
        Assert.Equal("Conflict", stale.ErrorCategory);
        Assert.Equal(after.Code, RequireSuccess(_scriptCommands.View(_fixture.BatchToken, moduleName)).Code);
    }

    [Fact]
    public void ScriptCommands_Invoke_PreservesTypedArgumentsAndTwoDimensionalReturn()
    {
        const string source = """
            Public Function TypedResult(ByVal text As Variant, ByVal number As Variant, ByVal flag As Variant, ByVal nullable As Variant, Optional absent As Variant) As Variant
                Dim output(1 To 2, 1 To 3) As Variant
                output(1, 1) = text
                output(1, 2) = number
                output(1, 3) = flag
                output(2, 1) = IsNull(nullable)
                output(2, 2) = IsMissing(absent)
                output(2, 3) = Null
                TypedResult = output
            End Function
            """;
        Import(_fixture.BatchToken, "TypedInvocation", source);
        var result = _fixture.Send("vba.invoke", new
        {
            procedureName = "TypedInvocation.TypedResult",
            arguments = new object[]
            {
                new { type = "string", value = "=literal" }, new { type = "number", value = -12.5 },
                new { type = "boolean", value = false }, new { type = "null" }, new { type = "missing" }
            }
        });
        Assert.True(result.Success, result.ErrorMessage);
        using var json = JsonDocument.Parse(result.Result!);
        Assert.Equal("array", json.RootElement.GetProperty("returnKind").GetString());
        var values = json.RootElement.GetProperty("values");
        Assert.Equal("=literal", values[0][0].GetString());
        Assert.Equal(-12.5, values[0][1].GetDouble());
        Assert.False(values[0][2].GetBoolean());
        Assert.True(values[1][0].GetBoolean());
        Assert.True(values[1][1].GetBoolean());
        Assert.Equal(JsonValueKind.Null, values[1][2].ValueKind);
    }

    [Fact]
    public async Task ScriptCommands_Invoke_RejectsObjectReturnAndRetainsMacroEffects()
    {
        Import(_fixture.BatchToken, "ObjectInvocation", """
            Public Function ReturnRange() As Object
                ThisWorkbook.Worksheets(1).Range("A1").Value = "invoked"
                Set ReturnRange = ThisWorkbook.Worksheets(1).Range("A1")
            End Function
            """);
        var result = await _fixture.SendForFailureAsync("vba.invoke", new { procedureName = "ObjectInvocation.ReturnRange" });
        Assert.Contains("unsupported", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        var values = RequireSuccess(_commands.GetValues(_fixture.BatchToken, "Sheet1", "A1"));
        Assert.Equal("invoked", Assert.Single(Assert.Single(values.Values)));
    }

    [Fact]
    public async Task ScriptCommands_Invoke_ReopenedMacrosStillRequireExplicitApproval()
    {
        Import(_fixture.BatchToken, "ApprovedInvocation", "Public Function ReturnLabel() As String\nReturnLabel = \"approved\"\nEnd Function");
        await _fixture.SaveAndReopenAsync();
        var blocked = await _fixture.SendForFailureAsync("vba.invoke", new { procedureName = "ApprovedInvocation.ReturnLabel" });
        Assert.Contains("enable-macros", blocked.ErrorMessage, StringComparison.Ordinal);
        _fixture.Send("session.enable-macros", new { approved = true, allowReopen = true });
        var result = _fixture.Send("vba.invoke", new { procedureName = "ApprovedInvocation.ReturnLabel" });
        using var json = JsonDocument.Parse(result.Result!);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal("string", json.RootElement.GetProperty("returnKind").GetString());
        Assert.Equal("approved", json.RootElement.GetProperty("value").GetString());
    }

    private static void AssertStoredCode(string expected, string actual)
    {
        expected = expected.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace("\n", "\r\n", StringComparison.Ordinal);
        actual = actual.Trim();
        // VBE canonicalizes identifier casing; literals and comments must remain exact.
        Assert.Equal(expected, actual, ignoreCase: true);
        const string literalsAndComments = "\"(?:[^\"\\r\\n]|\"\")*\"|'[^\\r\\n]*";
        Assert.Equal(
            System.Text.RegularExpressions.Regex.Matches(expected, literalsAndComments).Select(match => match.Value),
            System.Text.RegularExpressions.Regex.Matches(actual, literalsAndComments).Select(match => match.Value));
    }

    [Fact]
    public async Task ScriptCommands_Invoke_RejectsLargeResultAfterExecutionAndPreservesEffects()
    {
        Import(_fixture.BatchToken, "OversizedInvocation", """
            Public Function ReturnLarge() As Variant
                Dim output(1 To 101, 1 To 100) As Variant
                output(1, 1) = "synthetic result content"
                ThisWorkbook.Worksheets(1).Range("A1").Value2 = "execution completed"
                ReturnLarge = output
            End Function
            """);
        var result = await _fixture.SendForFailureAsync("vba.invoke", new { procedureName = "OversizedInvocation.ReturnLarge" });
        Assert.Contains("Macro execution completed", result.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("10000", result.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("not rolled back", result.ErrorMessage, StringComparison.Ordinal);
        Assert.DoesNotContain("synthetic result content", result.ErrorMessage, StringComparison.Ordinal);
        Assert.Null(result.Result);
        Assert.Equal("execution completed", RequireSuccess(_commands.GetValues(_fixture.BatchToken, "Sheet1", "A1")).Values[0][0]);
        RequireSuccess(_scriptCommands.Run(_fixture.BatchToken, "OversizedInvocation.ReturnLarge", timeout: null));
        Assert.Equal("execution completed", RequireSuccess(_commands.GetValues(_fixture.BatchToken, "Sheet1", "A1")).Values[0][0]);
    }

    [Fact]
    public async Task ScriptCommands_ImportCaseInsensitiveConflict_DoesNotCreateAnotherComponent()
    {
        const string moduleName = "ImportCaseGuard";
        Import(_fixture.BatchToken, moduleName, "Public Sub Keep()\nEnd Sub");
        var before = RequireSuccess(_scriptCommands.List(_fixture.BatchToken));
        var result = await _fixture.SendForFailureAsync("vba.import", new
        {
            moduleName = moduleName.ToLowerInvariant(),
            vbaCode = "Public Sub Replacement()\nEnd Sub"
        });
        Assert.Equal("Conflict", result.ErrorCategory);
        var after = RequireSuccess(_scriptCommands.List(_fixture.BatchToken));
        Assert.Equal(before.Scripts.Select(script => script.Name), after.Scripts.Select(script => script.Name));
        Assert.Contains("Keep", RequireSuccess(_scriptCommands.View(_fixture.BatchToken, moduleName)).Procedures);
    }

    private void Import(
        IExcelBatch batch,
        string moduleName,
        string vbaCode)
    {
        RequireSuccess(_scriptCommands.Import(batch, moduleName, vbaCode));
        _fixture.RegisterVbaModuleForCleanup(moduleName);
    }

    private void Delete(IExcelBatch batch, string moduleName)
    {
        RequireSuccess(_scriptCommands.Delete(batch, moduleName));
        _fixture.ForgetVbaModule(moduleName);
    }
}
