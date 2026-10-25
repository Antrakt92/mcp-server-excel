using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "PreCommit")]
[Trait("Speed", "Fast")]
public sealed class ForkValidationSelectionTests
{
    [Fact(Timeout = 150000)]
    public async Task ForkBuilder_SelectsPackagingAndScriptSafetyWithoutExcel()
    {
        var result = await ValidationSelectionTests.RunAsync("""
            foreach ($path in @('scripts/Build-ForkRelease.ps1', 'scripts\Build-ForkRelease.ps1')) {
                $plan = Get-ValidationPlan -Paths $path
                if (($plan.ToolingProjects -join ',') -cne 'Packaging,ScriptSafety') {
                    throw "Wrong fork builder owners: $($plan.ToolingProjects -join ',')"
                }
                if (-not $plan.ToolingFilters.ContainsKey('Packaging') -or
                    $plan.ToolingFilters.ScriptSafety -cne 'Feature=PreCommit') {
                    throw 'Fork builder regression filters are missing.'
                }
                if ($plan.Excel -or $plan.FullE2E -or $plan.ExcelGroups.Count) {
                    throw 'Packaging-only change selected Excel work.'
                }
                if ($plan.PackageComponents.Count -or $plan.NpmTests) {
                    throw 'Fork builder change selected unrelated upstream distributions.'
                }
            }
            """);
        Assert.True(result.ExitCode == 0, result.Output);
    }
}
