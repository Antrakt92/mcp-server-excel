using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

internal static class VbaInvocationValues
{
    internal const int MaximumReturnCells = 10_000;
    private const long MaximumReturnBytes = 4 * 1024 * 1024;

    internal static object[] Arguments(List<VbaArgument>? arguments)
    {
        if (arguments?.Count > 30) throw new ArgumentException("Excel accepts at most 30 VBA arguments.", nameof(arguments));
        var values = new object[30];
        Array.Fill(values, Type.Missing);
        for (int index = 0; index < (arguments?.Count ?? 0); index++)
        {
            var argument = arguments![index] ?? throw new ArgumentException("An argument cannot be null; use type null explicitly.", nameof(arguments));
            object? value = argument.Value;
            if (value is JsonElement json)
                value = json.ValueKind switch
                {
                    JsonValueKind.Null => null,
                    JsonValueKind.String => json.GetString(),
                    JsonValueKind.Number when json.TryGetDouble(out double number) && double.IsFinite(number) => number,
                    JsonValueKind.True => true,
                    JsonValueKind.False => false,
                    _ => throw new ArgumentException("VBA arguments must contain scalar values.", nameof(arguments))
                };
            values[index] = argument.Type switch
            {
                "string" when value is string text => text,
                "boolean" when value is bool boolean => boolean,
                "number" when IsNumber(value) && double.IsFinite(Convert.ToDouble(value, CultureInfo.InvariantCulture)) => Convert.ToDouble(value, CultureInfo.InvariantCulture),
                "null" when value is null => DBNull.Value,
                "missing" when value is null => Type.Missing,
                _ => throw new ArgumentException($"Argument {index + 1} must have a supported type and a matching scalar value.", nameof(arguments))
            };
        }
        return values;
    }

    internal static VbaInvokeResult Result(object? returned)
    {
        ValidateReturnBudget(returned);
        if (returned is Array array)
        {
            var values = new List<List<object?>>(array.GetLength(0));
            for (int row = 0; row < array.GetLength(0); row++)
            {
                var cells = new List<object?>(array.GetLength(1));
                for (int column = 0; column < array.GetLength(1); column++)
                    cells.Add(Scalar(array.GetValue(array.GetLowerBound(0) + row, array.GetLowerBound(1) + column)));
                values.Add(cells);
            }
            return new VbaInvokeResult { Success = true, ReturnKind = "array", Values = values };
        }
        object? value = Scalar(returned);
        return new VbaInvokeResult
        {
            Success = true,
            ReturnKind = value switch { null => "null", string => "string", bool => "boolean", _ => "number" },
            Value = value
        };
    }

    private static void ValidateReturnBudget(object? returned)
    {
        // Cover structured content, JSON text, and the additional escaping of that text.
        const int copies = 3;
        long bytes = 4096 * copies;
        void Add(long estimate)
        {
            long encodedBytes = estimate * copies;
            if (encodedBytes > MaximumReturnBytes - bytes)
                throw ResultRejected("The conservative 4 MiB return-value transport budget was exceeded. Return a smaller value.");
            bytes += encodedBytes;
        }
        static long ScalarBytes(object? value) => Scalar(value) switch
        {
            string text => 2L + 6L * text.Length,
            null => 4,
            bool => 5,
            _ => 64
        };
        if (returned is Array array)
        {
            if (array.Rank != 2 || array.LongLength > MaximumReturnCells)
                throw ResultRejected($"Return arrays require two dimensions and at most {MaximumReturnCells} elements.");
            Add(32L * array.GetLength(0) + 16L * array.LongLength);
            foreach (var value in array) Add(ScalarBytes(value));
        }
        else Add(ScalarBytes(returned));
    }

    private static InvalidOperationException ResultRejected(string reason) =>
        new($"Macro execution completed, but its result was rejected. {reason} Workbook effects were not rolled back. Inspect the workbook before another invocation.");

    internal static void ReleaseReturnedObjects(object? value)
    {
        if (value is Array array)
            foreach (var item in array) ReleaseReturnedObjects(item);
        else if (value is not null && Marshal.IsComObject(value))
            ComUtilities.Release(ref value);
    }

    private static object? Scalar(object? value)
    {
        if (value is null or DBNull) return null;
        if (value is string or bool) return value;
        if (IsNumber(value) && double.IsFinite(Convert.ToDouble(value, CultureInfo.InvariantCulture))) return value;
        throw ResultRejected("An unsupported value or COM object was returned. Only scalar/null or two-dimensional scalar arrays are supported.");
    }

    private static bool IsNumber(object? value) => value is byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal;
}
