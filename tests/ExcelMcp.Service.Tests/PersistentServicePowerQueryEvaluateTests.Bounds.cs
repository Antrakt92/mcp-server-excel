using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePowerQueryEvaluateTests
{
    [Fact]
    public void Evaluate_MultiTileResult_KeepsHeadersAndNativeValues()
    {
        const string code = """
            let
                Source = Table.FromColumns({List.Numbers(1, 1001)}, {"Index"}),
                Typed = Table.TransformColumnTypes(Source, {{"Index", Int64.Type}}),
                Tagged = Table.AddColumn(Typed, "Even", each Number.Mod([Index], 2) = 0, type logical)
            in Tagged
            """;
        var rows = Enumerable.Range(1, 1001).Select(i => new object[] { i, i % 2 == 0 }).ToArray();
        EvaluateChecked(code, ["Index", "Even"], rows);
    }

    [Fact]
    public async Task Evaluate_OversizedResult_ReportsLimitAndCleansTemporaryArtifacts()
    {
        const string code = """
            Table.FromColumns({List.Numbers(1, 30000)}, {"Value"})
            """;
        var guard = CreateEvaluationGuard();
        var before = SnapshotNativeArtifacts();
        var failure = await _fixture.SendForFailureAsync("powerquery.evaluate", new { mCode = code });

        Assert.Equal("InvalidInput", failure.ErrorCategory);
        Assert.Contains("4 MiB response byte budget", failure.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("Table.FirstN", failure.ErrorMessage, StringComparison.Ordinal);
        Assert.Null(failure.Result);
        Assert.Equal(before, SnapshotNativeArtifacts());
        AssertEvaluationGuard(guard);
    }
}
