using System.Globalization;
using JinoSupporter.Web.Services.BmesReports.Contracts;

namespace JinoSupporter.Web.Services.BmesReports;

public sealed class BmesKpiReportCalculationService
{
    public KpiTabDto Calculate(
        BmesFCostCalculationSnapshot fcost,
        FCostCorePartsKpiSnapshot? coreParts,
        IpgDefectKpiSnapshot? ipg,
        DateOnly reportEnd,
        ICollection<ReportIssueDto> warnings)
    {
        if (coreParts is null)
            warnings.Add(new ReportIssueDto(
                "kpi-core-parts-unavailable",
                "Core-parts KPI source is unavailable.",
                "kpi-core-parts",
                true));
        if (ipg is null)
            warnings.Add(new ReportIssueDto(
                "kpi-ipg-unavailable",
                "IPG KPI source is unavailable.",
                "kpi-ipg",
                true));

        IReadOnlyList<ReportPeriodDto> periods = BmesReportProjection.FromFCost(fcost.Report)
            .Where(period => period.Kind is "week" or "month")
            .OrderBy(period => period.StartDate ?? DateOnly.MaxValue)
            .ThenByDescending(period => period.Kind == "month")
            .ThenBy(period => period.Key, StringComparer.Ordinal)
            .ToArray();
        Dictionary<int, BmesFCostKpiPeriodValue> fcostByIndex =
            fcost.KpiPeriods.ToDictionary(value => value.ColumnIndex);
        Dictionary<string, FCostCorePartsKpiPeriod> coreByKey = BuildCoreMap(coreParts);
        Dictionary<string, IpgDefectKpiPeriod> ipgByKey = BuildIpgMap(ipg, reportEnd.Year);

        return new KpiTabDto
        {
            Periods = periods,
            Metrics =
            [
                new KpiMetricDto
                {
                    Id = "main-excess-material-cost",
                    Name = "초과투입 재료비(Main)",
                    Type = "법인",
                    BaselineValue = 1.79,
                    TargetValue = 1.0,
                    Unit = "percent",
                    Lines =
                    [
                        Line("actual", "실적", "percent", periods, period =>
                            FcostValue(period, fcostByIndex)?.TotalRate),
                        Line("achievement", "달성율", "percent", periods, period =>
                            Achievement(1.0, FcostValue(period, fcostByIndex)?.TotalRate)),
                        Line("excess-cost", "초과 투입 재료비", "usd", periods, period =>
                            FcostValue(period, fcostByIndex)?.TotalCost),
                        Line("sales-share", "매출대비비중", "none", periods, _ => null),
                    ],
                },
                new KpiMetricDto
                {
                    Id = "internalized-excess-material-cost",
                    Name = "초과투입 재료비(내재화)",
                    Type = "법인",
                    BaselineValue = 0.95,
                    TargetValue = 0.76,
                    Unit = "percent",
                    Lines =
                    [
                        Line("actual", "실적", "percent", periods, period =>
                            coreByKey.GetValueOrDefault(period.Key)?.TotalRatePercent),
                        Line("achievement", "달성율", "percent", periods, period =>
                            Achievement(0.76, coreByKey.GetValueOrDefault(period.Key)?.TotalRatePercent)),
                        Line("excess-cost", "초과 투입 재료비", "usd", periods, period =>
                            coreByKey.GetValueOrDefault(period.Key)?.TotalCostUsd),
                        Line("sales-share", "매출대비비중", "none", periods, _ => null),
                    ],
                },
                new KpiMetricDto
                {
                    Id = "main-process-defect-improvement",
                    Name = "Main 공정불량 개선율",
                    Type = "법인",
                    BaselineValue = 73_934,
                    TargetValue = 40_000,
                    Unit = "ppm",
                    Lines =
                    [
                        Line("actual", "실적", "ppm", periods, period =>
                            FcostValue(period, fcostByIndex)?.MainDefectAveragePpm),
                        Line("achievement", "달성율", "percent", periods, period =>
                            Achievement(40_000, FcostValue(period, fcostByIndex)?.MainDefectAveragePpm)),
                    ],
                },
                new KpiMetricDto
                {
                    Id = "ipg-process-defect-improvement",
                    Name = "IPG 공정불량 개선율",
                    Type = "법인",
                    BaselineValue = 1_403,
                    TargetValue = 1_000,
                    Unit = "ppm",
                    Lines =
                    [
                        Line("actual", "실적", "ppm", periods, period =>
                            ipgByKey.GetValueOrDefault(period.Key)?.AveragePpm, ipg?.AnnualAveragePpm),
                        Line("achievement", "달성율", "percent", periods, period =>
                            Achievement(1_000, ipgByKey.GetValueOrDefault(period.Key)?.AveragePpm),
                            Achievement(1_000, ipg?.AnnualAveragePpm)),
                    ],
                },
            ],
        };
    }

    private static KpiLineDto Line(
        string kind,
        string label,
        string unit,
        IReadOnlyList<ReportPeriodDto> periods,
        Func<ReportPeriodDto, double?> valueOf,
        double? annualValue = null) => new()
    {
        Kind = kind,
        Label = label,
        Unit = unit,
        AnnualValue = Finite(annualValue),
        ValuesByPeriod = periods.ToDictionary(
            period => period.Key,
            period => Finite(valueOf(period)),
            StringComparer.Ordinal),
    };

    private static BmesFCostKpiPeriodValue? FcostValue(
        ReportPeriodDto period,
        IReadOnlyDictionary<int, BmesFCostKpiPeriodValue> values) =>
        period.SourceIndex is int index ? values.GetValueOrDefault(index) : null;

    private static double? Achievement(double target, double? actual) =>
        actual is > 0 ? Finite(target / actual.Value * 100d) : null;

    private static double? Finite(double? value) =>
        value is not null && double.IsFinite(value.Value) ? value : null;

    private static Dictionary<string, FCostCorePartsKpiPeriod> BuildCoreMap(FCostCorePartsKpiSnapshot? snapshot)
    {
        var result = new Dictionary<string, FCostCorePartsKpiPeriod>(StringComparer.Ordinal);
        if (snapshot is null)
            return result;
        foreach (FCostCorePartsKpiPeriod period in snapshot.Periods)
        {
            string? key = PeriodKey(period.Kind, period.PDate, period.Code);
            if (key is not null)
                result.TryAdd(key, period);
        }
        return result;
    }

    private static Dictionary<string, IpgDefectKpiPeriod> BuildIpgMap(IpgDefectKpiSnapshot? snapshot, int fallbackYear)
    {
        var result = new Dictionary<string, IpgDefectKpiPeriod>(StringComparer.Ordinal);
        if (snapshot is null)
            return result;
        foreach (IpgDefectKpiPeriod period in snapshot.Periods)
        {
            string? key = period.Kind.Equals("Week", StringComparison.OrdinalIgnoreCase)
                ? IpgWeekKey(period.Header, fallbackYear)
                : period.Kind.Equals("Month", StringComparison.OrdinalIgnoreCase)
                    ? IpgMonthKey(period.Header, fallbackYear)
                    : null;
            if (key is not null)
                result.TryAdd(key, period);
        }
        return result;
    }

    private static string? PeriodKey(string kind, params string[] candidates)
    {
        string prefix = kind.Equals("Week", StringComparison.OrdinalIgnoreCase) ? "W:"
            : kind.Equals("Month", StringComparison.OrdinalIgnoreCase) ? "M:"
            : string.Empty;
        if (prefix.Length == 0)
            return null;
        foreach (string candidate in candidates)
        {
            string digits = new((candidate ?? string.Empty).Where(char.IsDigit).ToArray());
            if (digits.Length >= 6)
                return prefix + digits[..6];
        }
        return null;
    }

    private static string? IpgWeekKey(string header, int fallbackYear)
    {
        string text = (header ?? string.Empty).Trim().ToUpperInvariant();
        int marker = text.IndexOf('W');
        if (marker < 0)
            return null;
        string yearDigits = new(text[..marker].Where(char.IsDigit).ToArray());
        string weekDigits = new(text[(marker + 1)..].Where(char.IsDigit).ToArray());
        if (!int.TryParse(weekDigits, NumberStyles.None, CultureInfo.InvariantCulture, out int week) || week is < 1 or > 54)
            return null;
        int year = fallbackYear;
        if (yearDigits.Length >= 4)
            int.TryParse(yearDigits[^4..], out year);
        else if (yearDigits.Length == 2 && int.TryParse(yearDigits, out int shortYear))
            year = 2000 + shortYear;
        return year is >= 1 and <= 9999 ? $"W:{year:D4}{week:D2}" : null;
    }

    private static string? IpgMonthKey(string header, int fallbackYear)
    {
        string digits = new((header ?? string.Empty).Where(char.IsDigit).ToArray());
        int year = fallbackYear;
        int month;
        if (digits.Length >= 6)
        {
            if (!int.TryParse(digits[..4], out year) || !int.TryParse(digits.Substring(4, 2), out month))
                return null;
        }
        else if (digits.Length == 4)
        {
            if (!int.TryParse(digits[..2], out int shortYear) || !int.TryParse(digits.Substring(2, 2), out month))
                return null;
            year = 2000 + shortYear;
        }
        else
        {
            return null;
        }
        return year is >= 1 and <= 9999 && month is >= 1 and <= 12 ? $"M:{year:D4}{month:D2}" : null;
    }
}
