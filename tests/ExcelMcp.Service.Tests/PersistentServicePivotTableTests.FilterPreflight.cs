using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    [Theory]
    [InlineData("mixed")]
    [InlineData("missing")]
    [InlineData("empty")]
    public async Task RegularMemberFilter_InvalidSelectionPreservesExistingFilter(string scenario)
    {
        var batch = _fixture.BatchToken;
        const string pivotName = "FilterPreflightPivot";
        RequireSuccess(_pivotCommands.CreateFromRange(batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", pivotName));
        RequireSuccess(_pivotCommands.AddRowField(batch, pivotName, "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, pivotName, "Sales"));
        RequireSuccess(_pivotCommands.SetFieldFilter(batch, pivotName, "Region", ["North"]));
        RequireSuccess(_pivotCommands.AddFieldFilter(batch, pivotName, "Region",
            new PivotFilterOptions { Type = PivotFilterType.CaptionEquals, Text1 = "North" }));

        var before = ReadRegionSelection();
        Assert.False(before.ShowAll);
        Assert.Equal(["North"], before.SelectedMemberUniqueNames);
        Assert.Single(before.Filters);
        string savedSelection = JsonSerializer.Serialize(before);
        string savedData = JsonSerializer.Serialize(RequireSuccess(_pivotCommands.GetData(batch, pivotName)).Values);
        string[] selectedValues = scenario switch
        {
            "mixed" => ["South", "MissingRegion"],
            "missing" => ["MissingRegion"],
            _ => []
        };

        var error = await _fixture.SendForFailureAsync("pivottablefield.set-field-filter",
            new { pivotTableName = pivotName, fieldName = "Region", selectedValues });

        Assert.Equal("InvalidInput", error.ErrorCategory);
        Assert.Contains(scenario == "empty" ? "at least one" : "MissingRegion",
            error.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(savedSelection, JsonSerializer.Serialize(ReadRegionSelection()));
        Assert.Equal(savedData, JsonSerializer.Serialize(RequireSuccess(_pivotCommands.GetData(batch, pivotName)).Values));

        PivotFieldSelectionSnapshot ReadRegionSelection() => Assert.Single(
            RequireSuccess(_pivotCommands.Read(batch, pivotName, _salesSheetName)).FieldSelections,
            field => field.FieldName == "Region");
    }

    [Fact]
    public void RegularMemberFilter_ReadbackReturnsTheExactSelectedMembers()
    {
        var batch = _fixture.BatchToken;
        const string pivotName = "FilterReadbackPivot";
        RequireSuccess(_pivotCommands.CreateFromRange(batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", pivotName));
        RequireSuccess(_pivotCommands.AddRowField(batch, pivotName, "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, pivotName, "Sales"));
        RequireSuccess(_pivotCommands.SetFieldFilter(batch, pivotName, "Region", ["North"]));

        var applied = RequireSuccess(_pivotCommands.SetFieldFilter(batch, pivotName, "Region", ["South"]));

        Assert.Equal(["South"], applied.SelectedItems);
        Assert.Equal(["North", "South"], applied.AvailableItems.Order(StringComparer.Ordinal));
        Assert.False(applied.ShowAll);
        var native = Assert.Single(RequireSuccess(_pivotCommands.Read(batch, pivotName, _salesSheetName)).FieldSelections);
        Assert.Equal(["South"], native.SelectedMemberUniqueNames);
        var rows = RequireSuccess(_pivotCommands.GetData(batch, pivotName)).Values;
        Assert.Equal("South", rows[1][0]);
        AssertOriginalSales();
    }

    [Theory]
    [InlineData(false, 325)]
    [InlineData(true, 650)]
    public void RegularMemberFilter_ReportFieldAppliesExactSelection(bool multiple, int expectedTotal)
    {
        var batch = _fixture.BatchToken;
        const string pivotName = "ReportMemberPivot";
        RequireSuccess(_commands.SetValues(batch, _salesSheetName, "A7:D7",
            [["West", "Widget", 999, "2025-04-01"]]));
        RequireSuccess(_pivotCommands.CreateFromRange(batch, _salesSheetName, "A1:D7", _salesSheetName, "F1", pivotName));
        RequireSuccess(_pivotCommands.AddFilterField(batch, pivotName, "Region"));
        RequireSuccess(_pivotCommands.AddRowField(batch, pivotName, "Product"));
        RequireSuccess(_pivotCommands.AddValueField(batch, pivotName, "Sales"));
        RequireSuccess(_pivotCommands.Refresh(batch, pivotName));
        List<string> requested = multiple ? ["North", "South"] : ["South"];

        var applied = RequireSuccess(_pivotCommands.SetFieldFilter(batch, pivotName, "Region", requested));

        Assert.Equal(requested, applied.SelectedItems.Order(StringComparer.Ordinal));
        Assert.Equal(["North", "South", "West"], applied.AvailableItems.Order(StringComparer.Ordinal));
        Assert.False(applied.ShowAll);
        var values = RequireSuccess(_pivotCommands.GetData(batch, pivotName)).Values;
        var totals = Assert.Single(values, row => Equals(row[0], "Grand Total"));
        Assert.NotNull(totals[1]);
        Assert.Equal(expectedTotal, Convert.ToInt32(totals[1], System.Globalization.CultureInfo.InvariantCulture));
    }
}
