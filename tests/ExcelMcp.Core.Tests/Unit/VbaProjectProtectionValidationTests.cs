using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "VBA")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class VbaProjectProtectionValidationTests
{
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("synthetic\0value")]
    [InlineData("synthetic\rvalue")]
    [InlineData("synthetic\nvalue")]
    [InlineData("synthetic\tvalue")]
    [InlineData("SyntheticCaf\u00e9!7")]
    [InlineData("Synthetic\u041f\u0440\u0438\u043c\u0435\u0440!7")]
    [InlineData("Synthetic\u30c6\u30b9\u30c8!7")]
    [InlineData("Synthetic\U0001F600!7")]
    public void InvalidPassword_IsRejectedBeforeBatchAccess_WithoutEchoingValue(string? password)
    {
        var commands = new VbaCommands();
        var error = Assert.ThrowsAny<ArgumentException>(() => commands.ProtectProject(null!, password!));
        Assert.Equal("password", error.ParamName);
        if (!string.IsNullOrEmpty(password)) Assert.DoesNotContain(password, error.ToString(), StringComparison.Ordinal);
    }
}
