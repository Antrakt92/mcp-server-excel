using Sbroenne.ExcelMcp.Core.Commands.Slicer;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("RequiresExcel", "false")]
[Trait("Category", "Unit")]
[Trait("Feature", "Slicer")]
[Trait("Layer", "Core")]
[Trait("Speed", "Fast")]
public sealed class NonOlapSlicerSelectionTests
{
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void NoMatchingNativeNames_FailsBeforeAnySelectionWrite(bool clearFirst)
    {
        var writes = new List<(string Name, bool Selected)>();
        var cache = CreateCache(writes);
        var error = Assert.Throws<ArgumentException>(() =>
            SlicerSelection.SetNonOlapSelection(cache, ["missing", "unknown"], clearFirst, CancellationToken.None));

        Assert.Equal("selectedItems", error.ParamName);
        Assert.Contains("None of the requested slicer items were found", error.Message);
        Assert.Empty(writes);
    }

    [Fact]
    public void ReplaceSelection_IgnoresUnknownNamesAndSelectsBeforeDeselecting()
    {
        var writes = new List<(string Name, bool Selected)>();
        var cache = CreateCache(writes);

        SlicerSelection.SetNonOlapSelection(cache, ["missing", "south", "SOUTH"], clearFirst: true, CancellationToken.None);

        Assert.Equal([("South", true), ("North", false)], writes);
    }

    [Fact]
    public void AddSelection_IgnoresUnknownNamesAndLeavesOtherMembersSelected()
    {
        var writes = new List<(string Name, bool Selected)>();
        var cache = CreateCache(writes);

        SlicerSelection.SetNonOlapSelection(cache, ["missing", "south"], clearFirst: false, CancellationToken.None);

        Assert.Equal([("South", true)], writes);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EmptySelection_SelectsAllMembers(bool clearFirst)
    {
        var writes = new List<(string Name, bool Selected)>();
        var cache = CreateCache(writes);

        SlicerSelection.SetNonOlapSelection(cache, [], clearFirst, CancellationToken.None);

        Assert.Equal([("North", true), ("South", true)], writes);
    }

    private static Excel.SlicerCache CreateCache(List<(string Name, bool Selected)> writes)
    {
        var members = new[] { CreateMember("North", writes), CreateMember("South", writes) };
        var items = ExcelInterfaceProxy.Create<Excel.SlicerItems>((method, args) => method.Name switch
        {
            "get_Count" => members.Length,
            "get_Item" => members[Assert.IsType<int>(args![0]) - 1],
            _ => throw new NotSupportedException(method.Name)
        });
        return ExcelInterfaceProxy.Create<Excel.SlicerCache>((method, _) => method.Name == "get_SlicerItems"
            ? items : throw new NotSupportedException(method.Name));
    }

    private static Excel.SlicerItem CreateMember(string name, List<(string Name, bool Selected)> writes)
    {
        return ExcelInterfaceProxy.Create<Excel.SlicerItem>((method, args) =>
        {
            if (method.Name == "get_Name")
                return name;
            Assert.Equal("set_Selected", method.Name);
            writes.Add((name, Assert.IsType<bool>(args![0])));
            return null;
        });
    }
}
