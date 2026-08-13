using System.Globalization;
using System.Text.RegularExpressions;

namespace md2visio.vsdx.@base;

internal static partial class VisioValueParser
{
    public static double ParseResult(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new FormatException("Visio returned an empty numeric result.");

        var match = NumericValue().Match(value);
        if (!match.Success)
            throw new FormatException($"Visio result '{value}' does not contain a number.");

        var normalized = match.Value.Replace(',', '.');
        return double.Parse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture);
    }

    [GeneratedRegex(@"[-+]?(?:\d+(?:[.,]\d+)?|[.,]\d+)(?:[eE][-+]?\d+)?")]
    private static partial Regex NumericValue();
}
