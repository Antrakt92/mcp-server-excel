using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Slicer;

internal static class SlicerConnections
{
    internal static Excel.PivotTable? Find(Excel.SlicerPivotTables members, string pivotTableName,
        string? pivotSheetName, CancellationToken ct)
    {
        Excel.PivotTable? match = null;
        string? matchSheet = null;
        try
        {
            int count = members.Count;
            for (int index = 1; index <= count; index++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.PivotTable? member = null;
                try
                {
                    member = members[index];
                    if (!string.Equals(member.Name, pivotTableName, StringComparison.OrdinalIgnoreCase))
                        continue;
                    string sheetName = CoreLookupHelpers.GetPivotWorksheetName(member);
                    if (pivotSheetName is not null &&
                        !string.Equals(sheetName, pivotSheetName, StringComparison.OrdinalIgnoreCase))
                        continue;
                    if (match is not null)
                    {
                        throw new OperationFailureException(OperationFailureCategory.Conflict,
                            $"PivotTable '{pivotTableName}' is connected on both '{matchSheet}' and '{sheetName}'. Specify pivotSheetName.");
                    }
                    matchSheet = sheetName;
                    match = member;
                    member = null;
                }
                finally
                {
                    ComUtilities.Release(ref member);
                }
            }
            var result = match;
            match = null;
            return result;
        }
        finally
        {
            ComUtilities.Release(ref match);
        }
    }
}
