using System.Text.RegularExpressions;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;


namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range formula validation operations
/// Read-only lexical checks and local worksheet-name validation; not an Excel formula compiler.
/// </summary>
public partial class RangeCommands
{
    /// <inheritdoc />
    public RangeFormulaValidationResult ValidateFormulas(IExcelBatch batch, string sheetName, string rangeAddress, List<List<object?>>? formulas = null, string? formulasFile = null)
    {
        // Resolve formula cells (formulas, text, numbers, booleans, null) from inline parameter or file
        var resolvedFormulas = ParameterTransforms.ResolveFormulaCellsOrFile(formulas, formulasFile);

        var result = new RangeFormulaValidationResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = rangeAddress,
            Formulas = resolvedFormulas,
            FormulaCount = resolvedFormulas.Sum(row => row.Count),
            IsValid = true
        };

        return batch.Execute((ctx, ct) =>
        {
            var errors = new List<FormulaValidationError>();
            var sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Excel.Sheets? sheets = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                for (int index = 1; index <= sheets.Count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Worksheet? sheet = null;
                    try
                    {
                        sheet = (Excel.Worksheet)sheets[index];
                        sheetNames.Add(sheet.Name);
                    }
                    finally { ComUtilities.Release(ref sheet); }
                }
            }
            finally { ComUtilities.Release(ref sheets); }
            int validCount = 0;

            Excel.Range? target = null;
            int startRow;
            int startCol;
            try
            {
                target = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress, out string? targetError)
                    ?? throw new ArgumentException(targetError ?? RangeHelpers.GetResolveError(sheetName, rangeAddress), nameof(rangeAddress));
                // Same rectangle/payload geometry as set-formulas, without inspecting or changing content.
                ValidateContentWriteDimensions(target, resolvedFormulas, nameof(formulas), "Formula");
                (startRow, startCol) = ParseCellAddress(target.Address.Split(':')[0]);
            }
            finally { ComUtilities.Release(ref target); }

            int currentRow = startRow;
            foreach (var formulaRow in resolvedFormulas)
            {
                int currentCol = startCol;
                foreach (var formula in formulaRow)
                {
                    ct.ThrowIfCancellationRequested();
                    // Generate cell address for error reporting
                    string cellAddress = GetCellAddress(currentRow, currentCol);

                    // Empty cells and constants (text, numbers, booleans) are valid; only '=' formulas need checks
                    if (string.IsNullOrWhiteSpace(formula) || !formula.StartsWith('='))
                    {
                        validCount++;
                        currentCol++;
                        continue;
                    }

                    // Validate formula
                    var validationErrors = ValidateSingleFormula(formula, cellAddress, currentRow, currentCol, sheetNames);
                    if (validationErrors.Count == 0)
                    {
                        validCount++;
                    }
                    else
                    {
                        result.IsValid = false;
                        errors.AddRange(validationErrors);
                    }

                    currentCol++;
                }
                currentRow++;
            }

            result.ValidCount = validCount;
            result.ErrorCount = errors.Count;
            if (errors.Count > 0)
            {
                result.Errors = errors;
            }

            result.Success = true;
            return result;
        });
    }

    /// <summary>
    /// Validates a single formula and returns list of validation errors
    /// </summary>
    internal static List<FormulaValidationError> ValidateSingleFormula(string formula, string cellAddress, int row, int col, IReadOnlySet<string>? sheetNames = null)
    {
        var errors = new List<FormulaValidationError>();
        void Error(string message, string category = "syntax-error") => errors.Add(new FormulaValidationError
        {
            CellAddress = cellAddress,
            Row = row,
            Column = col,
            Formula = formula,
            Message = message,
            Category = category
        });
        int parentheses = 0;
        var bangs = new List<int>();
        for (int index = 1; index < formula.Length; index++)
        {
            char character = formula[index];
            if (character is '"' or '\'')
            {
                char quote = character;
                bool closed = false;
                while (++index < formula.Length)
                {
                    if (formula[index] != quote) continue;
                    if (index + 1 < formula.Length && formula[index + 1] == quote) { index++; continue; }
                    closed = true;
                    break;
                }
                if (!closed) Error("Unterminated quoted text or worksheet name.");
                continue;
            }
            if (character == '[')
            {
                // Structured-reference labels and external workbook labels are not formula syntax.
                int brackets = 1;
                while (++index < formula.Length && brackets != 0)
                {
                    if (formula[index] == '\'' && index + 1 < formula.Length) { index++; continue; }
                    if (formula[index] == '[') brackets++;
                    if (formula[index] == ']') brackets--;
                }
                index--;
                continue;
            }
            if (character == '(') parentheses++;
            if (character == ')')
            {
                if (parentheses == 0) Error("Extra closing parenthesis.");
                else parentheses--;
            }
            if (character == '!') bangs.Add(index);
        }
        if (parentheses > 0) Error($"Missing {parentheses} closing parenthesis(es).");
        foreach (int bang in bangs)
        {
            // Workbook-qualified UDF calls do not establish a local worksheet reference.
            if (Regex.IsMatch(formula[(bang + 1)..], @"^[\p{L}_][\p{L}\p{Nd}_.]*\s*\(", RegexOptions.CultureInvariant)) continue;
            var match = Regex.Match(formula[..bang], @"('(?:[^']|'')+'|(?:\[[^\]]+\])?[\p{L}\p{Nd}_.$\\:]+)$", RegexOptions.CultureInvariant);
            if (!match.Success || match.Index > 0 && formula[match.Index - 1] == '#') continue;
            string qualifier = match.Value;
            if (qualifier.StartsWith('\'')) qualifier = qualifier[1..^1].Replace("''", "'", StringComparison.Ordinal);
            // External references cannot be resolved without opening files, and are left unverified.
            if (qualifier.Contains('[') || qualifier.Contains(']') || sheetNames is null) continue;
            if (!sheetNames.Contains(qualifier) && Regex.IsMatch(qualifier, @"\.(?:xlsx?|xlsm|xlsb|xlam|xla|xll)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)) continue;
            foreach (string name in qualifier.Split(':'))
                if (!sheetNames.Contains(name)) Error($"Worksheet '{name}' does not exist in the selected workbook.", "invalid-reference");
        }
        return errors;
    }

    /// <summary>
    /// Converts row,col (1-based) to Excel cell address (e.g., "A1", "B5")
    /// </summary>
    private static string GetCellAddress(int row, int col)
    {
        // Convert column number to letter(s)
        string colLetter = "";
        int temp = col;
        while (temp > 0)
        {
            temp--;
            colLetter = (char)('A' + (temp % 26)) + colLetter;
            temp /= 26;
        }
        return colLetter + row;
    }

    /// <summary>
    /// Parses Excel cell address (e.g., "B1", "AA5") to (row, col) tuple
    /// </summary>
    internal static (int row, int col) ParseCellAddress(string cellAddress)
    {
        var match = Regex.Match(cellAddress, @"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$", RegexOptions.CultureInvariant);
        if (!match.Success) throw new ArgumentException("The validation target must start at an A1 cell reference.", nameof(cellAddress));
        int column = 0;
        foreach (char character in match.Groups[1].Value.ToUpperInvariant()) column = column * 26 + character - 'A' + 1;
        int row = int.Parse(match.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);
        if (column > 16384 || row > 1048576) throw new ArgumentException("The validation target is outside Excel's worksheet bounds.", nameof(cellAddress));
        return (row, column);
    }
}
