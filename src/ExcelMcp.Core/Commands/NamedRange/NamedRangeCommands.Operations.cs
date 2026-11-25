using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named range lifecycle operations (List, Read, Write, Create, Update, Delete)
/// </summary>
public partial class NamedRangeCommands
{
    private const long MaxListValuePreviewCellCount = 10_000;

    private static string ToR1C1Formula(Excel.Application application, string formula)
    {
        Excel.Range? activeCell = null;
        try
        {
            activeCell = (Excel.Range)application.ActiveCell;
            object? converted = application.ConvertFormula(
                formula,
                Excel.XlReferenceStyle.xlA1,
                Excel.XlReferenceStyle.xlR1C1,
                Type.Missing,
                activeCell);
            if (converted is string convertedFormula)
                return convertedFormula;

            // Excel cannot convert structured table references, which already use the same syntax in R1C1.
            if (formula.Contains('[') && !formula.Contains("]!", StringComparison.Ordinal))
                return formula;

            throw new InvalidOperationException("Excel could not convert the named formula to R1C1 notation.");
        }
        finally
        {
            ComUtilities.Release(ref activeCell);
        }
    }

    /// <inheritdoc />
    public NamedRangeListResult List(IExcelBatch batch)
    {
        var result = new NamedRangeListResult
        {
            FilePath = batch.WorkbookPath
        };

        return batch.Execute(ListNames);

        [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
            Justification = "RefersToRange can reject formula/external names; reports that specific getter failure per-name as Unavailable with native HRESULT. Other metadata and preview failures propagate.")]
        NamedRangeListResult ListNames(ExcelContext ctx, CancellationToken ct)
        {
            var budget = new RangeReadBudget("Read individual names with namedrange read.");
            budget.AddScalar(result.FilePath);
            var previews = new List<(int Index, NamedRangeInfo Info)>();
            Excel.Names? namesCollection = null;
            try
            {
                namesCollection = ctx.Book.Names;
                int count = namesCollection.Count;

                for (int i = 1; i <= count; i++)
                {
                    ct.ThrowIfCancellationRequested();

                    Excel.Name? nameObj = null;
                    Excel.Range? refersToRange = null;
                    try
                    {
                        nameObj = namesCollection.Item(i);
                        string name = nameObj.Name?.ToString() ?? string.Empty;

                        if (ShouldSkipNameFromList(nameObj, name))
                        {
                            continue;
                        }

                        string refersTo = ((object?)nameObj.RefersTo)?.ToString() ?? string.Empty;
                        budget.ReserveRecord(6);
                        budget.AddScalar(name);
                        budget.AddScalar(refersTo);

                        var info = new NamedRangeInfo
                        {
                            Name = name,
                            RefersTo = refersTo,
                            ValueType = "null"
                        };

                        try
                        {
                            refersToRange = nameObj.RefersToRange;
                        }
                        catch (COMException ex) when (ex.HResult == unchecked((int)0x800A03EC) || ComErrors.IsMemberUnavailable(ex))
                        {
                            // Named range may not have a valid RefersToRange (e.g., formula-based or external reference)
                            // Continue with metadata only - this is expected for some named ranges.
                            info.ValueType = "Unavailable";
                            info.ValueOmittedReason = $"{ex.Message} (HRESULT 0x{ex.HResult:X8}).";
                        }
                        if (refersToRange is not null)
                        {
                            if (RangeHelpers.BelongsToWorkbook(refersToRange, ctx.Book))
                            {
                                if (PrepareListValuePreview(refersToRange, info))
                                    previews.Add((i, info));
                            }
                            else
                            {
                                info.ValueType = "Unavailable";
                                info.ValueOmittedReason = RangeHelpers.ExternalNamedRangeMessage(name);
                            }
                        }

                        budget.AddScalar(info.ValueType);
                        budget.AddScalar(info.ValueOmittedReason);
                        result.NamedRanges.Add(info);
                    }
                    finally
                    {
                        ComUtilities.Release(ref refersToRange);
                        ComUtilities.Release(ref nameObj);
                    }
                }

                // Charge every entry's metadata and omission status before preview values.
                // Reacquire each range only while reading it.
                foreach (var (index, info) in previews)
                {
                    ct.ThrowIfCancellationRequested();
                    PopulateListValuePreview(ctx, namesCollection, index, info, budget, ct);
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref namesCollection);
            }
        }
    }

    private static bool ShouldSkipNameFromList(Excel.Name nameObj, string name)
    {
        if (IsHiddenName(nameObj))
        {
            return true;
        }

        return IsBuiltInName(name);
    }

    private static bool IsHiddenName(Excel.Name nameObj) => !nameObj.Visible;

    private static bool IsBuiltInName(string name)
    {
        var localName = GetLocalName(name);
        return localName.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)
            || localName.Equals("_FilterDatabase", StringComparison.OrdinalIgnoreCase);
    }

    private static string GetLocalName(string name)
    {
        var bangIndex = name.LastIndexOf('!');
        return bangIndex >= 0 ? name[(bangIndex + 1)..].Trim('\'') : name;
    }

    private static Excel.Name? FindNameByKey(Excel.Workbook workbook, string name)
    {
        Excel.Names? names = null;
        Excel.Name? nameObj = null;
        try
        {
            names = workbook.Names;
            int count = names.Count;
            for (int index = 1; index <= count; index++)
            {
                nameObj = names.Item(index);
                if (string.Equals(nameObj.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    var result = nameObj;
                    nameObj = null;
                    return result;
                }
                ComUtilities.Release(ref nameObj);
            }
            return null;
        }
        finally
        {
            ComUtilities.Release(ref nameObj);
            ComUtilities.Release(ref names);
        }
    }

    private static bool PrepareListValuePreview(Excel.Range refersToRange, NamedRangeInfo info)
    {
        var areaCount = GetAreaCount(refersToRange);
        if (areaCount > 1)
        {
            info.ValueType = "MultiAreaRange";
            info.ValueOmittedReason = "Named range resolves to multiple areas; list omits multi-area value previews.";
            return false;
        }

        var cellCount = GetCellCount(refersToRange);
        info.CellCount = cellCount;

        if (cellCount > MaxListValuePreviewCellCount)
        {
            info.ValueType = "RangeTooLarge";
            info.ValueOmittedReason =
                $"Named range contains {cellCount} cells, which exceeds the list preview limit of {MaxListValuePreviewCellCount}.";
            return false;
        }

        info.ValueType = "RangeTooLarge";
        info.ValueOmittedReason = "This preview would exceed the list response byte budget. Read the name separately.";
        return true;
    }

    private static void PopulateListValuePreview(ExcelContext context, Excel.Names names, int index,
        NamedRangeInfo info, RangeReadBudget budget, CancellationToken cancellationToken)
    {
        Excel.Name? name = null;
        Excel.Range? range = null;
        try
        {
            name = names.Item(index);
            if (!string.Equals(name.Name, info.Name, StringComparison.Ordinal) || !name.Visible ||
                !string.Equals(((object?)name.RefersTo)?.ToString(), info.RefersTo, StringComparison.Ordinal))
                throw new InvalidOperationException("Workbook names changed while reading the list. Retry to refresh the metadata.");

            range = name.RefersToRange;
            if (range is null || !RangeHelpers.BelongsToWorkbook(range, context.Book) ||
                GetAreaCount(range) != 1 || GetCellCount(range) != info.CellCount)
                throw new InvalidOperationException("A named range changed while reading the list. Retry to refresh the metadata.");

            if (budget.TryAddPreview(() => ReadPreview(context, range, info.CellCount, budget, cancellationToken),
                    out var preview))
            {
                info.Value = preview.Value;
                info.ValueType = preview.Type;
                info.ValueOmittedReason = null;
            }
        }
        finally
        {
            ComUtilities.Release(ref range);
            ComUtilities.Release(ref name);
        }
    }

    private static (object? Value, string Type) ReadPreview(ExcelContext context, Excel.Range range, long? cellCount,
        RangeReadBudget budget, CancellationToken cancellationToken)
    {
        if (cellCount != 1)
        {
            var size = RangeHelpers.GetDimensions(range);
            var rows = BoundedRangeValueReader.Read(context, range, size, budget, cancellationToken,
                convertValue: ConvertValueForJson);
            return (rows, "Array");
        }

        var (value, type) = ReadScalarValue(context, range);
        budget.AddScalar(value);
        budget.AddScalar(type);
        return (value, type);
    }

    private static (object? Value, string Type) ReadScalarValue(ExcelContext context, Excel.Range? range)
    {
        object? rawValue = range?.Value2;
        if (range is not null && RangeCommands.ReadNativeError(context, range, 1, 1, rawValue) is { } error)
            return (error.Name, "ExcelError");

        return (ConvertValueForJson(rawValue), rawValue?.GetType().Name ?? "null");
    }

    private static int GetAreaCount(Excel.Range range)
    {
        Excel.Areas? areas = null;
        try
        {
            areas = range.Areas;
            return areas.Count;
        }
        finally
        {
            ComUtilities.Release(ref areas);
        }
    }

    private static long GetCellCount(Excel.Range range)
    {
        return Convert.ToInt64((object)range.CountLarge, CultureInfo.InvariantCulture);
    }

    internal static object ParseWriteValue(string value)
    {
        if (double.TryParse(
            value,
            NumberStyles.Float,
            CultureInfo.InvariantCulture,
            out double numericValue))
        {
            return numericValue;
        }

        return bool.TryParse(value, out bool booleanValue)
            ? booleanValue
            : value;
    }

    /// <inheritdoc />
    public OperationResult Write(IExcelBatch batch, string name, string value)
    {
        var cellValue = RangeHelpers.ConvertToCellValue(ParseWriteValue(value), literalText: true);
        return batch.Execute((ctx, ct) =>
        {
            Excel.Name? nameObj = null;
            Excel.Range? refersToRange = null;
            int originalCalculation = -1;
            bool calculationChanged = false;
            Exception? operationFailure = null;

            try
            {
                nameObj = FindNameByKey(ctx.Book, name);
                if (nameObj == null)
                {
                    throw new InvalidOperationException($"Named range '{name}' not found.");
                }

                refersToRange = nameObj.RefersToRange;
                RangeHelpers.RequireNamedRangeWorkbook(refersToRange, ctx.Book, name);

                // Calculation suppressed here (not in ExcelWriteGuard) because Data Model ops need it enabled
                originalCalculation = (int)ctx.App.Calculation;
                if (originalCalculation != -4135) // xlCalculationManual
                {
                    ctx.App.Calculation = (Excel.XlCalculation)(-4135);
                    calculationChanged = true;
                }

                refersToRange.Value2 = cellValue;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Dummy return for batch.Execute
            }
            catch (Exception ex)
            {
                operationFailure = ex;
                throw;
            }
            finally
            {
                try
                {
                    if (calculationChanged && originalCalculation != -1)
                    {
                        OperationCleanup.Run(() => ctx.App.Calculation = (Excel.XlCalculation)originalCalculation,
                            operationFailure);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref refersToRange);
                    ComUtilities.Release(ref nameObj);
                }
            }
        });
    }

    /// <inheritdoc />
    public NamedRangeValue Read(IExcelBatch batch, string name, bool metadataOnly = false)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Name? nameObj = null;
            Excel.Range? refersToRange = null;
            try
            {
                nameObj = FindNameByKey(ctx.Book, name);
                if (nameObj == null)
                {
                    throw new InvalidOperationException($"Named range '{name}' not found.");
                }

                string refersTo = ((object?)nameObj.RefersTo)?.ToString() ?? "";
                var budget = new RangeReadBudget();
                budget.AddScalar(name);
                budget.AddScalar(refersTo);
                if (metadataOnly)
                    return ReadMetadata(ctx.Book, nameObj, name, refersTo, budget);

                refersToRange = nameObj.RefersToRange;
                if (refersToRange is not null)
                {
                    RangeHelpers.RequireNamedRangeWorkbook(refersToRange, ctx.Book, name);
                    if (GetAreaCount(refersToRange) != 1)
                        throw new ArgumentException("Named range reads require one rectangular range. Read each area separately.", nameof(name));

                    var size = RangeHelpers.GetDimensions(refersToRange);
                    if (size != (1, 1))
                    {
                        return new NamedRangeValue
                        {
                            Name = name,
                            RefersTo = refersTo,
                            Value = BoundedRangeValueReader.Read(ctx, refersToRange, size, budget, ct,
                                convertValue: ConvertValueForJson),
                            ValueType = "Array"
                        };
                    }
                }

                var (value, type) = ReadScalarValue(ctx, refersToRange);
                budget.AddScalar(value);
                return new NamedRangeValue
                {
                    Name = name,
                    RefersTo = refersTo,
                    Value = value,
                    ValueType = type
                };
            }
            finally
            {
                ComUtilities.Release(ref refersToRange);
                ComUtilities.Release(ref nameObj);
            }
        });
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Metadata reports the targeted RefersToRange getter failure with its original HRESULT and message. Unexpected COM failures still propagate.")]
    private static NamedRangeValue ReadMetadata(
        Excel.Workbook book, Excel.Name nameObj, string name, string refersTo, RangeReadBudget budget)
    {
        var result = new NamedRangeValue
        {
            Name = name,
            RefersTo = refersTo,
            Visible = nameObj.Visible,
            ValueType = "NotRead"
        };

        Excel.Range? range = null;
        try
        {
            try
            {
                range = nameObj.RefersToRange;
            }
            catch (COMException ex) when (ex.HResult == unchecked((int)0x800A03EC) || ComErrors.IsMemberUnavailable(ex))
            {
                budget.AddScalar(ex.Message);
                result.ResolutionStatus = "Unavailable";
                result.ResolutionHResult = ex.HResult;
                result.ResolutionMessage = ex.Message;
                return result;
            }

            if (range is null)
            {
                result.ResolutionStatus = "Unavailable";
            }
            else if (RangeHelpers.BelongsToWorkbook(range, book))
            {
                result.ResolutionStatus = "LocalRange";
                result.AreaCount = GetAreaCount(range);
            }
            else
            {
                result.ResolutionStatus = "ExternalRange";
            }

            return result;
        }
        finally
        {
            ComUtilities.Release(ref range);
        }
    }

    /// <inheritdoc />
    public OperationResult Create(IExcelBatch batch, string name, string reference, bool visible = true)
    {
        // Validate parameter name length (Excel limit: 255 characters)
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("Named range name cannot be empty or whitespace", nameof(name));
        }

        if (name.Length > 255)
        {
            throw new ArgumentException($"Named range name exceeds Excel's 255-character limit (current length: {name.Length})", nameof(name));
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.Name? existing = null;
            Excel.Name? created = null;
            Excel.Names? namesCollection = null;
            try
            {
                // Check if parameter already exists
                existing = FindExistingNameForCreation(ctx.Book, name);
                if (existing != null)
                {
                    throw new InvalidOperationException($"Named range '{name}' already exists");
                }

                namesCollection = ctx.Book.Names;
                string formula = $"={reference.TrimStart('=')}";
                string r1c1Formula = ToR1C1Formula(ctx.App, formula);
                // Names.Add parses decimal separators locally; the property setter accepts invariant formulas.
                created = namesCollection.Add(Name: name, RefersTo: "=0", Visible: visible);
                if (created is null)
                    throw new InvalidOperationException(
                        $"Excel did not confirm creating named range '{name}'. Inspect workbook names before retrying.");
                created.RefersToR1C1 = r1c1Formula;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Dummy return for batch.Execute
            }
            catch (Exception ex) when (created is not null && CreatedObjectFailure.CanReport(ex))
            {
                throw DescribeCreatedNameFailure(name, ex, () =>
                    (object)created.RefersToR1C1 is string retainedReference
                        ? retainedReference
                        : throw new InvalidOperationException("Excel did not return the retained name's reference."));
            }
            finally
            {
                ComUtilities.Release(ref created);
                ComUtilities.Release(ref namesCollection);
                ComUtilities.Release(ref existing);
            }
        });
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "A failed retained-name inspection is combined with the original creation failure and rethrown; neither failure is suppressed.")]
    internal static InvalidOperationException DescribeCreatedNameFailure(
        string name, Exception cause, Func<string> readReference)
    {
        string remainingState;
        try
        {
            remainingState = $"Named range '{name}' remains with R1C1 reference '{readReference()}'.";
        }
        catch (Exception inspectionFailure) when (inspectionFailure is not OperationCanceledException)
        {
            remainingState = $"The current reference of named range '{name}' could not be inspected.";
            cause = cause is COMException assignmentFailure
                ? new NamedRangeAssignmentFailure(assignmentFailure, inspectionFailure)
                : new AggregateException(cause, inspectionFailure);
        }
        return CreatedObjectFailure.CreateDescribed(
            "Assigning the requested reference",
            $"named range '{name}'",
            $"{remainingState} It was not deleted automatically; inspect workbook names before retrying.",
            cause);
    }

    private sealed class NamedRangeAssignmentFailure : COMException
    {
        public NamedRangeAssignmentFailure(COMException assignmentFailure, Exception inspectionFailure)
            : base($"{assignmentFailure.Message} Retained-reference inspection also failed: {inspectionFailure.Message}" +
                (inspectionFailure is COMException com ? $" [COM: 0x{com.HResult:X8}]" : string.Empty),
                new AggregateException(assignmentFailure, inspectionFailure))
        {
            // Keep the primary native code discoverable while retaining both exception objects.
            HResult = assignmentFailure.HResult;
        }
    }

    /// <inheritdoc />
    public OperationResult Update(IExcelBatch batch, string name, string reference)
    {
        // Validate parameter name length (Excel limit: 255 characters)
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("Named range name cannot be empty or whitespace", nameof(name));
        }

        if (name.Length > 255)
        {
            throw new ArgumentException($"Named range name exceeds Excel's 255-character limit (current length: {name.Length})", nameof(name));
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.Name? nameObj = null;
            try
            {
                nameObj = FindNameByKey(ctx.Book, name);
                if (nameObj == null)
                {
                    throw new InvalidOperationException($"Named range '{name}' not found.");
                }

                string formula = $"={reference.TrimStart('=')}";
                nameObj.RefersToR1C1 = ToR1C1Formula(ctx.App, formula);

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Dummy return for batch.Execute
            }
            finally
            {
                ComUtilities.Release(ref nameObj);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Delete(IExcelBatch batch, string name)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Name? nameObj = null;
            try
            {
                nameObj = FindNameByKey(ctx.Book, name);
                if (nameObj == null)
                {
                    throw new InvalidOperationException($"Named range '{name}' not found.");
                }

                nameObj.Delete();
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath }; // Dummy return for batch.Execute
            }
            finally
            {
                ComUtilities.Release(ref nameObj);
            }
        });
    }
}
