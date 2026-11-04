using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "VBA")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class VbaProjectAuthenticationValidationTests
{
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("Synthetic\0Secret")]
    [InlineData("Synthetic\rSecret")]
    [InlineData("Synthetic\nSecret")]
    [InlineData("Synthetic\tSecret")]
    [InlineData("Synthetic\u007fSecret")]
    [InlineData("Synthetic\u00e9Secret")]
    [InlineData("Synthetic\u0416Secret")]
    [InlineData("Synthetic\U0001f600Secret")]
    public void InvalidPassword_BothActionsRejectBeforeBatchAccess_WithoutEcho(string? password)
    {
        var commands = new VbaCommands();
        AssertInvalid(() => commands.UnlockProject(null!, password!), password);
        AssertInvalid(() => commands.VerifyPassword(null!, password!), password);
    }

    private static void AssertInvalid(Action operation, string? password)
    {
        var error = Assert.ThrowsAny<ArgumentException>(operation);
        Assert.Equal("password", error.ParamName);
        if (!string.IsNullOrEmpty(password))
            Assert.DoesNotContain(password, error.ToString(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AttachedWorkbook_IsRefusedBeforeOwnerOrComAccess(bool verify)
    {
        using var attached = new AttachedWorkbookBatch(null!,
            new OpenWorkbookInfo("owned-selection", "Synthetic.xlsm", "C:/synthetic/Synthetic.xlsm", 1, true, false));
        var commands = new VbaCommands();
        var error = Assert.Throws<OperationFailureException>(() =>
        {
            if (verify) commands.VerifyPassword(attached, "Synthetic!7");
            else commands.UnlockProject(attached, "Synthetic!7");
        });
        Assert.Equal(OperationFailureCategory.Permissions, error.ErrorCategory);
        Assert.Contains("Attached", error.Message, StringComparison.Ordinal);
    }
}
