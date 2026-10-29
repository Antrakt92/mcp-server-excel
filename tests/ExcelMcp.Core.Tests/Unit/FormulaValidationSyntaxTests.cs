using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Range")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class FormulaValidationSyntaxTests
{
    [Theory]
    [InlineData("=\"(\"")]
    [InlineData("=\"text \"\"(\"\"\"")]
    [InlineData("=ExistingSheet!A1")]
    [InlineData("='O''Brien (Data)'!$A$1")]
    [InlineData("=SUM(Table1[Column (label)])")]
    [InlineData("=GETVM3(1)")]
    [InlineData("=CustomUdf(1)")]
    [InlineData("=IFERROR(#REF!,0)")]
    [InlineData("='Book.xlsm'!CustomUdf(1)")]
    [InlineData("='Book1'!CustomUdf(1)")]
    public void ValidLiteralOrLocalReference_IsNotRejected(string formula)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "ExistingSheet", "O'Brien (Data)" };
        Assert.Empty(RangeCommands.ValidateSingleFormula(formula, "A1", 1, 1, names));
    }

    [Theory]
    [InlineData("$A$1", 1, 1)]
    [InlineData("$aa$12", 12, 27)]
    [InlineData("XFD1048576", 1048576, 16384)]
    public void TargetAddress_AcceptsAbsoluteAndCaseInsensitiveCells(string address, int row, int column) =>
        Assert.Equal((row, column), RangeCommands.ParseCellAddress(address));
    [Theory]
    [InlineData("=SUM(A1", "syntax-error")]
    [InlineData("=)(", "syntax-error")]
    [InlineData("=\"unterminated", "syntax-error")]
    [InlineData("=Missing!$A$1", "invalid-reference")]
    [InlineData("='Missing (Data)'!A1", "invalid-reference")]
    [InlineData("=missing!a1", "invalid-reference")]
    public void DefiniteLexicalOrMissingLocalReference_ReturnsAnError(string formula, string category)
    {
        var errors = RangeCommands.ValidateSingleFormula(formula, "A1", 1, 1, new HashSet<string>());
        Assert.Contains(errors, error => error.Category == category);
    }

    [Theory]
    [InlineData("='C:\\data\\[Book.xlsx]Missing'!A1")]
    [InlineData("=[Book.xlsx]Missing!A1")]
    [InlineData("=\"Missing!A1(\"")]
    public void ExternalOrLiteralReference_IsNotReportedAsMissingLocalSheet(string formula) =>
        Assert.Empty(RangeCommands.ValidateSingleFormula(formula, "A1", 1, 1, new HashSet<string>()));

    [Theory]
    [InlineData("XFE1")]
    [InlineData("A1048577")]
    [InlineData("A0")]
    [InlineData("not-an-address")]
    public void InvalidTarget_IsRejectedBeforeValidation(string address) =>
        Assert.Throws<ArgumentException>(() => RangeCommands.ParseCellAddress(address));
}
