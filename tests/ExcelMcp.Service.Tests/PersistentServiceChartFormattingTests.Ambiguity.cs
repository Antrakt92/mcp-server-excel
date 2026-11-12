using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceChartFormattingTests
{
    [Theory]
    [InlineData("chart.read")]
    [InlineData("chart.delete")]
    [InlineData("chart.move")]
    [InlineData("chartconfig.set-title")]
    public async Task DuplicateChartName_RejectsLookupBeforeChoosingOrChangingEitherChart(string command)
    {
        const string chartName = "RepeatedChart";
        var batch = _fixture.BatchToken;
        string secondSheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, secondSheet, "A1:B3",
            [["Category", "SecondSeries"], ["Q1", 100], ["Q2", 200]]));
        RequireSuccess(_chartCommands.CreateFromRange(batch, _sheetName, "A1:B4",
            ChartType.ColumnClustered, 120, 60, 320, 220, chartName));
        RequireSuccess(_chartCommands.CreateFromRange(batch, secondSheet, "A1:B3",
            ChartType.ColumnClustered, 180, 90, 370, 250, chartName));
        var charts = RequireSuccess(_chartCommands.List(batch)).Charts;
        var matching = charts.Where(chart => chart.Name == chartName).ToArray();
        Assert.Equal(2, matching.Length);
        Assert.Contains(matching, chart => chart.SheetName == _sheetName);
        Assert.Contains(matching, chart => chart.SheetName == secondSheet);
        string before = JsonSerializer.Serialize(charts);

        object args = command switch
        {
            "chart.move" => new { chartName, left = 240.0 },
            "chartconfig.set-title" => new { chartName, title = "Must not change either chart" },
            _ => new { chartName }
        };
        var failure = await _fixture.SendForFailureAsync(command, args);

        Assert.Equal("Conflict", failure.ErrorCategory);
        Assert.Contains(chartName, failure.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains(_sheetName, failure.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains(secondSheet, failure.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(before, JsonSerializer.Serialize(RequireSuccess(_chartCommands.List(batch)).Charts));
    }
}
