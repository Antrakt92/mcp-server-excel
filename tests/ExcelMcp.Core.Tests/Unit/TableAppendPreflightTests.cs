using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "Tables")]
public sealed class TableAppendPreflightTests
{
    [Fact]
    public void Append_EmptyJsonRowsFile_RejectsBeforeExecutingExcel()
    {
        string path = Path.Combine(Path.GetTempPath(), $"ExcelMcp_EmptyRows_{Guid.NewGuid():N}.json");
        try
        {
            File.WriteAllText(path, "[]");
            using var batch = ExcelInterfaceProxy.Create<IExcelBatch>((method, _) =>
                method.Name == nameof(IDisposable.Dispose)
                    ? null
                    : throw new InvalidOperationException($"Excel batch was accessed through {method.Name}."));

            var error = Assert.Throws<ArgumentException>(() =>
                new TableCommands().Append(batch, "SalesTable", rowsFile: path));

            Assert.Equal("rows", error.ParamName);
            Assert.Contains("No data to append", error.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
