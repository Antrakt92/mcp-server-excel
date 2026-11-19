using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "VBA")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class VbaLabelDiagnosticsTests
{
    [Fact]
    public void Analyze_FindsLocalAndCrossModuleSubsInModuleOrder()
    {
        var diagnostics = VbaLabelScanner.Analyze(
        [
            new("Fixture", true, """"
                Private Static Sub LocalReset()
                End Sub
                Sub Probe()
                    localReset: started = Timer
                    H2_Invalidate: hits = hits + 1
                    Call H2_Invalidate
                End Sub
                """"),
            new("Availability", true, "Public Sub H2_Invalidate()\r\nEnd Sub")
        ]);

        Assert.False(diagnostics.HasMore);
        Assert.Collection(diagnostics.Warnings,
            warning => AssertWarning(warning, "Fixture", "Probe", 4, "localReset"),
            warning => AssertWarning(warning, "Fixture", "Probe", 5, "H2_Invalidate"));
    }

    [Theory]
    [InlineData("GoTo ResetTimer")]
    [InlineData("On state GoTo 10, ResetTimer")]
    [InlineData("GoSub ResetTimer")]
    [InlineData("Resume ResetTimer")]
    [InlineData("On Error GoTo ResetTimer")]
    [InlineData("On state GoSub First, _\n        ResetTimer")]
    [InlineData("If state Then GoTo _\n        ResetTimer")]
    public void Analyze_ActualTargetsInSameProcedureAreExempt(string reference)
    {
        const string prefix = "Sub ResetTimer()\nEnd Sub\nSub Probe()\n";
        const string body = "ResetTimer: done = True\nEnd Sub";
        var withoutJump = VbaLabelScanner.Analyze([new("Module1", true, prefix + body)]);
        AssertWarning(Assert.Single(withoutJump.Warnings), "Module1", "Probe", 4, "ResetTimer");

        var withJump = VbaLabelScanner.Analyze([new("Module1", true, prefix + "    " + reference + "\n" + body)]);
        Assert.Empty(withJump.Warnings);
    }

    [Fact]
    public void Analyze_CommentsStringsAndOtherProceduresDoNotHideAWarning()
    {
        var diagnostics = VbaLabelScanner.Analyze(
        [
            new("Module1", true, """"
                Sub ResetTimer()
                End Sub
                Sub Probe()
                    Debug.Print "quoted ""and GoTo ResetTimer""": Rem GoSub ResetTimer
                    If done Then Rem GoTo ResetTimer
                    ' Resume ResetTimer
                ResetTimer: done = True
                End Sub
                Sub Other()
                    GoTo ResetTimer
                ResetTimer: done = False
                End Sub
                """")
        ]);

        AssertWarning(Assert.Single(diagnostics.Warnings), "Module1", "Probe", 7, "ResetTimer");
    }

    [Fact]
    public void Analyze_IgnoresOtherModulePrivateClassAndConditionalDeclarations()
    {
        var diagnostics = VbaLabelScanner.Analyze(
        [
            new("Target", true, """"
                Sub Probe()
                PrivateHelper: test = 1
                ClassHelper: test = 1
                BuildHelper: test = 1
                TakesArgument: test = 1
                #If Win64 Then
                VisibleHelper: test = 1
                #End If
                    Call Other( _
                        VisibleHelper:=test)
                End Sub
                """"),
            new("Helpers", true, """"
                Private Sub PrivateHelper(): End Sub
                #If Win64 Then
                Public Sub BuildHelper(): End Sub
                #End If
                Sub TakesArgument(Optional ByVal state As Variant)
                End Sub
                Sub VisibleHelper(): End Sub
                """"),
            new("Class1", false, "Public Sub ClassHelper()\nEnd Sub")
        ]);

        Assert.Empty(diagnostics.Warnings);
    }

    [Theory]
    [InlineData(100, false)]
    [InlineData(101, true)]
    public void Analyze_BoundsWarningsAndReportsOmissions(int labelCount, bool hasMore)
    {
        string declarations = string.Join('\n', Enumerable.Range(1, labelCount)
            .Select(index => $"Sub ResetTimer{index}(): End Sub"));
        string labels = string.Join('\n', Enumerable.Range(1, labelCount)
            .Select(index => $"ResetTimer{index}: hit = {index}"));

        var diagnostics = VbaLabelScanner.Analyze(
        [
            new("Module1", true, $"{declarations}\nSub Probe()\n{labels}\nEnd Sub")
        ]);

        Assert.Equal(100, diagnostics.Warnings.Count);
        Assert.Equal(hasMore, diagnostics.HasMore);
    }

    private static void AssertWarning(
        Sbroenne.ExcelMcp.Core.Models.VbaLabelWarning warning, string module, string procedure, int line, string label)
    {
        Assert.Equal(module, warning.ModuleName);
        Assert.Equal(procedure, warning.ProcedureName);
        Assert.Equal(line, warning.Line);
        Assert.Equal(label, warning.Label);
        Assert.Contains("Call", warning.Message, StringComparison.Ordinal);
    }
}
