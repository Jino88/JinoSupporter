using System.Globalization;
using JinoSupporter.Web.Services.BmesReports.Contracts;

namespace JinoSupporter.Web.Services.BmesReports;

internal static class BmesReportProjection
{
    public static double? Finite(double value) => double.IsFinite(value) ? value : null;
    public static double? PositiveOrNull(double value) => value > 0 && double.IsFinite(value) ? value : null;

    public static Dictionary<string, double?> Map(
        IEnumerable<ReportPeriodDto> periods,
        Func<string, double> valueOf) =>
        periods.ToDictionary(p => p.Key, p => Finite(valueOf(p.Key)), StringComparer.Ordinal);

    public static IReadOnlyList<ReportPeriodDto> FromNgRate(
        NgRateReportService.NgRateReport report,
        bool includeDates = true)
    {
        var result = new List<ReportPeriodDto>();
        int sortOrder = 0;
        if (includeDates)
            AddNgPeriods(result, report.DateCols, "date", ref sortOrder);
        AddNgPeriods(result, report.WeekCols, "week", ref sortOrder);
        AddNgPeriods(result, report.MonthCols, "month", ref sortOrder);
        return result;
    }

    public static IReadOnlyList<ReportPeriodDto> FromFCost(FCostReport report)
    {
        var result = new List<ReportPeriodDto>(report.Columns.Count);
        for (int index = 0; index < report.Columns.Count; index++)
        {
            FCostColumnMeta column = report.Columns[index];
            string kind = column.Kind switch
            {
                FCostPeriodKind.Day => "date",
                FCostPeriodKind.Week => "week",
                FCostPeriodKind.Month => "month",
                _ => "date",
            };
            string key = FCostPeriodKey(column);
            (DateOnly? start, DateOnly? end) = PeriodBounds(key, kind);
            result.Add(new ReportPeriodDto
            {
                Key = key,
                Kind = kind,
                Header = column.Header,
                SortOrder = index,
                StartDate = start,
                EndDateExclusive = end,
                SourceIndex = index,
                SourceCode = NullIfEmpty(column.Code),
                SourcePDate = NullIfEmpty(column.PDate),
            });
        }
        return result;
    }

    public static BmesReportRequestDto ProjectRequest(BmesReportRequest request) => new()
    {
        StartDate = request.StartDate,
        EndDate = request.EndDate,
        Groups = request.Groups.Select(group => new ReportSelectionGroupDto
        {
            Id = group.Id,
            Name = group.Name,
            MidGroups = group.MidGroups.Select(mid => new ReportSelectionMidDto
            {
                Material = mid.Material,
                LineShifts = mid.LineShifts
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Select(value => value.Trim())
                    .Distinct(StringComparer.Ordinal)
                    .ToArray(),
            }).ToArray(),
        }).ToArray(),
    };

    public static string FCostPeriodKey(FCostColumnMeta column)
    {
        string digits = PeriodDigits(column.PDate, column.Code);
        return column.Kind switch
        {
            FCostPeriodKind.Day when digits.Length >= 8 =>
                $"{digits[..4]}-{digits.Substring(4, 2)}-{digits.Substring(6, 2)}",
            FCostPeriodKind.Week when digits.Length >= 6 => "W:" + digits[..6],
            FCostPeriodKind.Month when digits.Length >= 6 => "M:" + digits[..6],
            _ => $"date:{column.Index:D2}",
        };
    }

    private static void AddNgPeriods(
        ICollection<ReportPeriodDto> target,
        IEnumerable<NgRateReportService.PeriodColumn> source,
        string kind,
        ref int sortOrder)
    {
        foreach (NgRateReportService.PeriodColumn column in source)
        {
            (DateOnly? start, DateOnly? end) = PeriodBounds(column.Key, kind);
            target.Add(new ReportPeriodDto
            {
                Key = column.Key,
                Kind = kind,
                Header = column.Header,
                SortOrder = sortOrder++,
                StartDate = start,
                EndDateExclusive = end,
            });
        }
    }

    private static (DateOnly? Start, DateOnly? End) PeriodBounds(string key, string kind)
    {
        if (kind == "date" && DateOnly.TryParseExact(
                key,
                "yyyy-MM-dd",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out DateOnly day))
            return (day, day.AddDays(1));

        string raw = key.Length > 2 && key[1] == ':' ? key[2..] : key;
        if (kind == "month" && raw.Length >= 6 &&
            int.TryParse(raw[..4], out int year) && int.TryParse(raw.Substring(4, 2), out int month) &&
            year is >= 1 and <= 9999 && month is >= 1 and <= 12)
        {
            var start = new DateOnly(year, month, 1);
            return (start, start.AddMonths(1));
        }

        if (kind == "week" && raw.Length >= 6 &&
            int.TryParse(raw[..4], out year) && int.TryParse(raw.Substring(4, 2), out int week) &&
            year is >= 1 and <= 9999 && week >= 1 && week <= ISOWeek.GetWeeksInYear(year))
        {
            DateTime monday = ISOWeek.ToDateTime(year, week, DayOfWeek.Monday);
            var start = DateOnly.FromDateTime(monday);
            return (start, start.AddDays(7));
        }

        return (null, null);
    }

    private static string PeriodDigits(params string[] values)
    {
        foreach (string value in values)
        {
            string digits = new((value ?? string.Empty).Where(char.IsDigit).ToArray());
            if (digits.Length >= 6)
                return digits;
        }
        return string.Empty;
    }

    private static string? NullIfEmpty(string value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;
}
