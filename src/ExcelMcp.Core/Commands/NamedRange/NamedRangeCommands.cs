namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named range/parameter management commands implementation
/// </summary>
public partial class NamedRangeCommands : INamedRangeCommands
{
    private static object? ConvertValueForJson(object? value)
    {
        if (value == null || value == DBNull.Value)
        {
            return null;
        }

        return value switch
        {
            string or bool or int or long or decimal => value,
            DateTime dateTime => dateTime.ToString("O", System.Globalization.CultureInfo.InvariantCulture),
            double doubleValue when double.IsNaN(doubleValue) || double.IsInfinity(doubleValue)
                => doubleValue.ToString(System.Globalization.CultureInfo.InvariantCulture),
            float floatValue when float.IsNaN(floatValue) || float.IsInfinity(floatValue)
                => floatValue.ToString(System.Globalization.CultureInfo.InvariantCulture),
            double or float => value,
            _ => value.ToString()
        };
    }
}



