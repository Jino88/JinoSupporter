using System.Globalization;

namespace GraphMaker;

public static class GraphMakerParsingHelper
{
    private static readonly string[] SupportedDateFormats =
    {
        "yyyy-MM-dd",
        "yyyy/MM/dd",
        "MM-dd",
        "MM/dd",
        "M-d",
        "M/d",
        "d-MMM",
        "dd-MMM",
        "d-MMM-yyyy",
        "dd-MMM-yyyy",
        "d-MMM-yy",
        "dd-MMM-yy",
        "d/MMM",
        "dd/MMM"
    };

    public static bool TryParseDouble(string? text, out double value)
    {
        if (double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands,
                CultureInfo.InvariantCulture, out value))
        {
            return true;
        }

        return double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands,
            CultureInfo.CurrentCulture, out value);
    }

    public static bool TryParseDate(string? text, out DateTime value)
    {
        value = default;

        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        string trimmed = text.Trim();

        foreach (string format in SupportedDateFormats)
        {
            if (!DateTime.TryParseExact(trimmed, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsed))
            {
                continue;
            }

            if (!format.Contains("y", StringComparison.OrdinalIgnoreCase))
            {
                parsed = new DateTime(DateTime.Now.Year, parsed.Month, parsed.Day);
            }

            value = parsed;
            return true;
        }

        if (DateTime.TryParse(trimmed, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime invariantParsed) ||
            DateTime.TryParse(trimmed, CultureInfo.CurrentCulture, DateTimeStyles.None, out invariantParsed))
        {
            value = invariantParsed;
            return true;
        }

        return false;
    }
}
