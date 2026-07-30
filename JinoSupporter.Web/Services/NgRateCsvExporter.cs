namespace JinoSupporter.Web.Services;

public static class NgRateCsvExporter
{
    public static byte[] ExportReports(
        List<(string Label, NgRateReportService.NgRateReport Report)> reports,
        List<(string Label, NgRateReportService.LineShiftNgReport Report)>? detailReports = null,
        bool includeDateColumns = true)
        => CsvExportUtility.Build(ReportRows(reports, detailReports, includeDateColumns));

    public static byte[] ExportWeeklyModelSummary(NgRateExcelExporter.WeeklyModelSummaryExport export)
        => CsvExportUtility.Build(WeeklyModelSummaryRows(export));

    private static IEnumerable<IEnumerable<string?>> ReportRows(
        List<(string Label, NgRateReportService.NgRateReport Report)> reports,
        List<(string Label, NgRateReportService.LineShiftNgReport Report)>? detailReports,
        bool includeDateColumns)
    {
        bool firstSection = true;
        foreach (var (label, report) in reports)
        {
            foreach (var row in NgRateReportRows(label, report, includeDateColumns, firstSection))
            {
                yield return row;
                firstSection = false;
            }
        }

        if (detailReports is null) yield break;

        foreach (var (label, report) in detailReports)
        {
            foreach (var row in LineShiftDetailRows(label, report, includeDateColumns, firstSection))
            {
                yield return row;
                firstSection = false;
            }
        }
    }

    private static IEnumerable<IEnumerable<string?>> NgRateReportRows(
        string label,
        NgRateReportService.NgRateReport report,
        bool includeDateColumns,
        bool firstSection)
    {
        var cols = AllCols(report, includeDateColumns);

        if (!firstSection) yield return Array.Empty<string?>();
        yield return new string?[] { label, "Summary - NG PPM by Process Type" };
        yield return new[] { "Process Type" }.Concat(cols.Select(c => c.Header));
        foreach (var row in report.Summary)
            yield return new[] { row.ProcessType }.Concat(PpmValues(cols, row.Ppm));

        yield return Array.Empty<string?>();
        yield return new string?[] { label, "Process 10 NG - Top 10 Worst Processes by PPM" };
        yield return new[] { "#", "Type", "Process Name" }.Concat(cols.Select(c => c.Header));
        foreach (var proc in report.Top10Process)
        {
            yield return new[] { proc.Rank.ToString(), proc.ProcessType, proc.ProcessName }
                .Concat(PpmValues(cols, proc.Ppm));
            foreach (var group in proc.Groups)
                yield return new[] { string.Empty, string.Empty, "  - " + group.GroupName }
                    .Concat(PpmValues(cols, group.Ppm));
        }

        yield return Array.Empty<string?>();
        yield return new string?[] { label, "Worst 10 NG Names" };
        yield return new[] { "#", "Type", "Process Name", "NG Name" }.Concat(cols.Select(c => c.Header));
        foreach (var ng in report.Top10Ng)
        {
            yield return new[] { ng.Rank.ToString(), ng.ProcessType, ng.ProcessName, ng.NgName }
                .Concat(PpmValues(cols, ng.Ppm));
            foreach (var group in ng.Groups)
                yield return new[] { string.Empty, string.Empty, string.Empty, "  - " + group.GroupName }
                    .Concat(PpmValues(cols, group.Ppm));
        }

        if (report.ReasonRows.Count == 0) yield break;

        yield return Array.Empty<string?>();
        yield return new string?[] { label, "Reason" };
        yield return new[] { "Reason", "#", "Type", "Process Name", "NG Name" }
            .Concat(cols.Select(c => c.Header));
        foreach (var reason in report.ReasonRows)
        {
            yield return new[]
                {
                    reason.Reason,
                    reason.IsTotal ? "Total" : reason.No.ToString(),
                    reason.ProcessType,
                    reason.ProcessName,
                    reason.NgName,
                }
                .Concat(PpmValues(cols, reason.Ppm));
        }
    }

    private static IEnumerable<IEnumerable<string?>> LineShiftDetailRows(
        string label,
        NgRateReportService.LineShiftNgReport report,
        bool includeDateColumns,
        bool firstSection)
    {
        var cols = AllCols(report, includeDateColumns);

        if (!firstSection) yield return Array.Empty<string?>();
        yield return new string?[] { label, "Detail - Top NG by Model / Group" };
        yield return new[] { "Model / Group", "Type", "Process Name", "NG Name" }
            .Concat(cols.Select(c => c.Header));

        foreach (var detail in report.Details
                     .OrderBy(d => d.LineShift, StringComparer.Ordinal)
                     .ThenBy(d => d.ProcessType, StringComparer.Ordinal)
                     .ThenBy(d => d.ProcessName, StringComparer.Ordinal)
                     .ThenBy(d => d.NgName, StringComparer.Ordinal))
        {
            yield return new[] { detail.LineShift, detail.ProcessType, detail.ProcessName, detail.NgName }
                .Concat(DetailPpmValues(cols, detail));
        }
    }

    private static IEnumerable<IEnumerable<string?>> WeeklyModelSummaryRows(
        NgRateExcelExporter.WeeklyModelSummaryExport export)
    {
        yield return new string?[] { export.Title };

        var header = new List<string?> { "Group / Sub Group", "Level" };
        header.AddRange(export.WeekColumns.Select(c => c.Header));
        header.AddRange(export.MonthColumns.Select(c => c.Header));
        yield return header;

        foreach (var row in export.Rows)
        {
            var values = new List<string?> { row.Label, row.Level };
            AddWeeklyModelSummaryValues(values, export.WeekColumns, row.WeekPpm, export.IncludeDelta);
            AddWeeklyModelSummaryValues(values, export.MonthColumns, row.MonthPpm, export.IncludeDelta);
            yield return values;
        }
    }

    private static void AddWeeklyModelSummaryValues(
        List<string?> values,
        IReadOnlyList<NgRateReportService.PeriodColumn> columns,
        IReadOnlyDictionary<string, double> ppm,
        bool includeDelta)
    {
        for (int i = 0; i < columns.Count; i++)
        {
            double current = ppm.GetValueOrDefault(columns[i].Key);
            double? previous = PreviousNonZeroPpm(columns, ppm, i);
            values.Add(includeDelta ? FormatPpmWithDeltaText(current, previous) : FormatPpmText(current));
        }
    }

    private static double? PreviousNonZeroPpm(
        IReadOnlyList<NgRateReportService.PeriodColumn> columns,
        IReadOnlyDictionary<string, double> ppm,
        int currentIndex)
    {
        for (int i = currentIndex + 1; i < columns.Count; i++)
        {
            double value = ppm.GetValueOrDefault(columns[i].Key);
            if (value > 0) return value;
        }
        return null;
    }

    private static List<NgRateReportService.PeriodColumn> AllCols(
        NgRateReportService.NgRateReport report,
        bool includeDateColumns)
    {
        var cols = new List<NgRateReportService.PeriodColumn>();
        if (includeDateColumns) cols.AddRange(report.DateCols);
        cols.AddRange(report.WeekCols);
        cols.AddRange(report.MonthCols);
        return cols;
    }

    private static List<NgRateReportService.PeriodColumn> AllCols(
        NgRateReportService.LineShiftNgReport report,
        bool includeDateColumns)
    {
        var cols = new List<NgRateReportService.PeriodColumn>();
        if (includeDateColumns) cols.AddRange(report.DateCols);
        cols.AddRange(report.WeekCols);
        cols.AddRange(report.MonthCols);
        return cols;
    }

    private static IEnumerable<string?> PpmValues(
        IEnumerable<NgRateReportService.PeriodColumn> columns,
        IReadOnlyDictionary<string, double> ppm)
        => columns.Select(c => FormatPpmText(ppm.GetValueOrDefault(c.Key)));

    private static IEnumerable<string?> DetailPpmValues(
        IEnumerable<NgRateReportService.PeriodColumn> columns,
        NgRateReportService.LineShiftNgDetail detail)
    {
        foreach (var column in columns)
        {
            double value =
                detail.DatePpm.GetValueOrDefault(column.Key) +
                detail.WeekPpm.GetValueOrDefault(column.Key) +
                detail.MonthPpm.GetValueOrDefault(column.Key);
            yield return FormatPpmText(value);
        }
    }

    private static string FormatPpmWithDeltaText(double current, double? previous)
    {
        string main = FormatPpmText(current);
        if (!previous.HasValue || (current == 0 && previous.Value == 0))
            return main;

        double delta = current - previous.Value;
        if (Math.Abs(delta) < 0.5)
            return main;

        long absRounded = (long)Math.Round(Math.Abs(delta));
        string sign = delta > 0 ? "+" : "-";
        return $"{main} ({sign}{absRounded:N0})";
    }

    private static string FormatPpmText(double value)
        => value == 0 ? "-" : ((long)Math.Round(value)).ToString("N0");
}
