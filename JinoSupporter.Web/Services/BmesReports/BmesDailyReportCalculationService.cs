using JinoSupporter.Web.Services.BmesReports.Contracts;

namespace JinoSupporter.Web.Services.BmesReports;

public sealed class BmesDailyReportCalculationService(
    NgRateService ngRate,
    NgRateSettingsService settings,
    NgRateReportService reportService)
{
    public async Task<BmesDailyCalculationSnapshot> CalculateAsync(
        BmesReportRequest request,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        ValidateRequest(request);
        if (!settings.IsCredentialsConfigured)
            throw new InvalidOperationException("BMES credentials are not configured.");

        cancellationToken.ThrowIfCancellationRequested();
        List<string> selectedLineShifts = request.Groups
            .SelectMany(group => group.MidGroups)
            .SelectMany(mid => mid.LineShifts)
            .Where(lineShift => !string.IsNullOrWhiteSpace(lineShift))
            .Select(lineShift => lineShift.Trim())
            .Distinct(StringComparer.Ordinal)
            .ToList();

        progress?.Report($"Fetching data: {request.StartDate:yyyy-MM-dd} ~ {request.EndDate:yyyy-MM-dd}");
        string? dbPath = await ngRate.GetOrFetchAsync(
            request.StartDateTime,
            request.EndDateTime,
            progress,
            selectedLineShifts);
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(dbPath))
            throw new InvalidOperationException("Daily data fetch or save failed.");

        progress?.Report($"Using DB: {Path.GetFileName(dbPath)}");
        HierReports hierarchy = await HierReportBuilder.BuildAsync(
            reportService,
            dbPath,
            request.Groups,
            progress,
            includeMidDetailReport: true);
        cancellationToken.ThrowIfCancellationRequested();

        if (!hierarchy.HasData || hierarchy.ByGroup is null || hierarchy.ByMid is null)
            throw new InvalidOperationException("Daily hierarchy did not produce data for the selected groups.");

        return new BmesDailyCalculationSnapshot
        {
            Hierarchy = hierarchy,
            // This intentionally matches the legacy callback, which handed the group
            // summary to F-COST for its group-key trend lookup.
            TrendByMid = hierarchy.ByGroup,
            Tab = Project(hierarchy, request.Groups),
            SourceDbPath = dbPath,
        };
    }

    public DailyTabDto Project(HierReports hierarchy, IReadOnlyList<ModelGroupRecord> groups)
    {
        NgRateReportService.NgRateReport report = hierarchy.ByGroup
            ?? throw new InvalidOperationException("Daily group report is missing.");
        NgRateReportService.NgRateReport byMid = hierarchy.ByMid
            ?? throw new InvalidOperationException("Daily model report is missing.");
        IReadOnlyList<ReportPeriodDto> periods = BmesReportProjection.FromNgRate(report);
        var summaryRows = new List<DailySummaryRowDto>();
        var sections = new List<DailyModelSectionDto>();

        foreach (ModelGroupRecord group in groups)
        {
            string groupRowId = "G::" + group.Name;
            summaryRows.Add(new DailySummaryRowDto
            {
                RowId = groupRowId,
                Depth = 0,
                Level = "group",
                GroupName = group.Name,
                Display = group.Name,
                PpmByPeriod = PpmMap(report, group.Name, periods),
            });

            foreach (MidGroupRecord mid in group.MidGroups)
            {
                if (string.IsNullOrWhiteSpace(mid.Material) || mid.LineShifts.Count == 0)
                    continue;

                string modelKey = NgRateModeSupport.ModelKey(group.Name, mid.Material);
                string modelRowId = "M::" + modelKey;
                summaryRows.Add(new DailySummaryRowDto
                {
                    RowId = modelRowId,
                    ParentRowId = groupRowId,
                    Depth = 1,
                    Level = "model",
                    GroupName = group.Name,
                    ModelName = mid.Material,
                    Display = mid.Material,
                    PpmByPeriod = PpmMap(byMid, modelKey, periods),
                });

                IReadOnlyList<NgRateReportService.SummaryPivotRow> modelRows =
                    byMid.GroupSummary.GetValueOrDefault(modelKey) ?? [];
                int processOrdinal = 0;
                foreach (NgRateReportService.SummaryPivotRow process in modelRows
                    .Where(row => !row.IsTotal)
                    .OrderBy(row => row.ProcessType, StringComparer.Ordinal))
                {
                    summaryRows.Add(new DailySummaryRowDto
                    {
                        RowId = $"PT::{modelKey}::{processOrdinal++}",
                        ParentRowId = modelRowId,
                        Depth = 2,
                        Level = "process-type",
                        GroupName = group.Name,
                        ModelName = mid.Material,
                        Display = process.ProcessType,
                        ProcessType = process.ProcessType,
                        PpmByPeriod = PeriodMap(periods, process.Ppm),
                    });
                }

                sections.Add(new DailyModelSectionDto
                {
                    Id = modelKey,
                    GroupName = group.Name,
                    ModelName = mid.Material,
                    LineShiftCount = mid.LineShifts.Count,
                    ReasonRows = ProjectReasonRows(byMid, modelKey, periods),
                });
            }
        }

        return new DailyTabDto
        {
            Periods = periods,
            SummaryRows = summaryRows,
            ModelSections = sections,
            ReferenceDatePeriodKey = periods.FirstOrDefault(period => period.Kind == "date")?.Key,
        };
    }

    private static IReadOnlyList<DailyReasonRowDto> ProjectReasonRows(
        NgRateReportService.NgRateReport report,
        string modelKey,
        IReadOnlyList<ReportPeriodDto> periods)
    {
        var result = new List<DailyReasonRowDto>();
        int sectionOrdinal = 0;
        foreach (IGrouping<string, NgRateReportService.ReasonRow> section in report.ReasonRows
                     .Where(row => !row.IsTotal)
                     .GroupBy(row => row.Reason)
                     .OrderBy(group => group.Key, StringComparer.Ordinal))
        {
            var details = new List<DailyReasonRowDto>();
            int rank = 0;
            foreach (NgRateReportService.ReasonRow row in section)
            {
                NgRateReportService.GroupPivotRow? group = row.Groups.FirstOrDefault(candidate =>
                    string.Equals(candidate.GroupName, modelKey, StringComparison.Ordinal));
                if (group is null)
                    continue;

                Dictionary<string, double?> ppm = PeriodMap(periods, group.Ppm);
                if (!ppm.Values.Any(value => value > 0))
                    continue;

                string parentId = $"R::{modelKey}::{sectionOrdinal}";
                details.Add(new DailyReasonRowDto
                {
                    RowId = $"{parentId}::{rank + 1}",
                    ParentRowId = parentId,
                    Reason = section.Key,
                    Rank = ++rank,
                    ProcessType = NullIfEmpty(row.ProcessType),
                    ProcessName = NullIfEmpty(row.ProcessName),
                    NgName = NullIfEmpty(row.NgName),
                    PpmByPeriod = ppm,
                });
            }

            if (details.Count == 0)
                continue;

            string totalRowId = $"R::{modelKey}::{sectionOrdinal++}";
            result.Add(new DailyReasonRowDto
            {
                RowId = totalRowId,
                Reason = section.Key,
                IsTotal = true,
                ProcessName = "Total",
                PpmByPeriod = periods.ToDictionary(
                    period => period.Key,
                    period => BmesReportProjection.Finite(details.Sum(detail =>
                        detail.PpmByPeriod.GetValueOrDefault(period.Key) ?? 0)),
                    StringComparer.Ordinal),
            });
            result.AddRange(details.Select(detail => new DailyReasonRowDto
            {
                RowId = detail.RowId.Replace(detail.ParentRowId!, totalRowId, StringComparison.Ordinal),
                ParentRowId = totalRowId,
                Reason = detail.Reason,
                Rank = detail.Rank,
                ProcessType = detail.ProcessType,
                ProcessName = detail.ProcessName,
                NgName = detail.NgName,
                PpmByPeriod = detail.PpmByPeriod,
            }));
        }
        return result;
    }

    private static Dictionary<string, double?> PpmMap(
        NgRateReportService.NgRateReport report,
        string key,
        IReadOnlyList<ReportPeriodDto> periods)
    {
        NgRateReportService.SummaryPivotRow? total = report.GroupSummary
            .GetValueOrDefault(key)?
            .FirstOrDefault(row => row.IsTotal);
        return PeriodMap(periods, total?.Ppm);
    }

    private static Dictionary<string, double?> PeriodMap(
        IReadOnlyList<ReportPeriodDto> periods,
        IReadOnlyDictionary<string, double>? values) =>
        periods.ToDictionary(
            period => period.Key,
            period => values is not null && values.TryGetValue(period.Key, out double value)
                ? BmesReportProjection.Finite(value)
                : null,
            StringComparer.Ordinal);

    private static string? NullIfEmpty(string value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;

    private static void ValidateRequest(BmesReportRequest request)
    {
        if (request.EndDate < request.StartDate)
            throw new ArgumentException("End date must be on or after start date.", nameof(request));
        if (request.Groups.Count == 0)
            throw new ArgumentException("At least one model group must be selected.", nameof(request));
        if (!request.Groups.SelectMany(group => group.MidGroups).SelectMany(mid => mid.LineShifts).Any())
            throw new ArgumentException("The selected groups do not contain a LineShift.", nameof(request));
    }
}
