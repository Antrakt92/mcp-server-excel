using System.Globalization;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

internal static class RangeReadAssertions
{
    internal static void Number(double expected, object? value)
    {
        Assert.True(value is byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal,
            $"Expected a numeric CLR value, received {value?.GetType().Name ?? "null"}.");
        Assert.Equal(expected, Convert.ToDouble(value, CultureInfo.InvariantCulture));
    }
}
