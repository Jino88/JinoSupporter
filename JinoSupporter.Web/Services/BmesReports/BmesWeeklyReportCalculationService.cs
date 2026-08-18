using JinoSupporter.Web.Services.BmesReports.Contracts;

namespace JinoSupporter.Web.Services.BmesReports;

public sealed class BmesWeeklyReportCalculationService(NgRateReportService reportService)
{
    private const string BaselineMonthKey = "202512";

    public async Task<BmesWeeklyCalculationSnapshot> CalculateAsync(
        BmesReportRequest request,
        BmesDailyCalculationSnapshot daily,
        IReadOnlyDictionary<string, WeeklyReportFormSettingRecord> formSettings,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        if (!daily.Hierarchy.HasData || daily.Hierarchy.ByGroup is null)
            throw new InvalidOperationException("Daily hierarchy is required for Weekly calculation.");
        if (string.IsNullOrWhiteSpace(daily.SourceDbPath) || !File.Exists(daily.SourceDbPath))
            throw new InvalidOperationException("Shared NG Rate DB was not found.");

        NgRateGroupMappings mappings = NgRateModeSupport.BuildGroupMappings(request.Groups);
        if (!mappings.HasData)
            throw new InvalidOperationException("No LineShift found in the selected groups.");

        progress?.Report($"Reusing Daily hierarchy report: {Path.GetFileName(daily.SourceDbPath)}");
        Task<NgRateReportService.LineShiftNgReport> lineShiftTask =
            reportService.ComputeLineShiftNgDetailsAsync(
                daily.SourceDbPath,
                mappings.LineShiftList,
                progress,
                request.StartDateTime,
                request.EndDateTime);
        Task<NgRateReportService.LineShiftNgReport> midTask =
            reportService.ComputeGroupedNgDetailsAsync(
                daily.SourceDbPath,
                mappings.MidMapping,
                progress,
                request.StartDateTime,
                request.EndDateTime);

        await Task.WhenAll(lineShiftTask, midTask);
        cancellationToken.ThrowIfCancellationRequested();
        NgRateReportService.LineShiftNgReport lineShift = await lineShiftTask;
        NgRateReportService.LineShiftNgReport mid = await midTask;

        return new BmesWeeklyCalculationSnapshot
        {
            Hierarchy = daily.Hierarchy,
            LineShiftDetails = lineShift,
            MidDetails = mid,
            Tab = Project(request, daily.Hierarchy, mid, formSettings),
        };
    }

    public WeeklyTabDto Project(
        BmesReportRequest request,
        HierReports hierarchy,
        NgRateReportService.LineShiftNgReport midDetails,
        IReadOnlyDictionary<string, WeeklyReportFormSettingRecord> formSettings)
    {
        NgRateReportService.NgRateReport byGroup = hierarchy.ByGroup
            ?? throw new InvalidOperationException("Weekly group report is missing.");
        NgRateReportService.NgRateReport byMid = hierarchy.ByMid
            ?? throw new InvalidOperationException("Weekly model report is missing.");
        IReadOnlyList<ReportPeriodDto> periods = BmesReportProjection.FromNgRate(byGroup);
        IReadOnlyList<ReportPeriodDto> weeklyPeriods = periods
            .Where(period => period.Kind is "week" or "month")
            .ToArray();
        var lookup = HierPpmLookup.From(hierarchy);

        List<WeeklyTargetRowDto> targets = BuildTargetRows(
            request.Groups,
            weeklyPeriods,
            lookup,
            formSettings);
        List<WeeklySummaryRowDto> summaries = BuildSummaryRows(request.Groups, periods, hierarchy);
        List<WeeklyTrendSeriesDto> trends = summaries
            .Where(row => row.Level is "group" or "model")
            .Select(row => new WeeklyTrendSeriesDto
            {
                Id = row.RowId,
                Label = row.Display,
                GroupName = row.GroupName,
                PpmByPeriod = new Dictionary<string, double?>(row.PpmByPeriod, StringComparer.Ordinal),
            })
            .ToList();

        string? sortKey = byGroup.WeekCols.Count >= 2
            ? byGroup.WeekCols[1].Key
            : byGroup.WeekCols.FirstOrDefault()?.Key ?? byGroup.DateCols.FirstOrDefault()?.Key;
        List<WeeklyDefectRowDto> defects = BuildTopDefects(request.Groups, periods, midDetails, sortKey);

        return new WeeklyTabDto
        {
            Periods = periods,
            TargetRows = targets,
            SummaryRows = summaries,
            TrendSeries = trends,
            TopDefects = defects,
            SortReferencePeriodKey = sortKey,
        };
    }

    private static List<WeeklyTargetRowDto> BuildTargetRows(
        IReadOnlyList<ModelGroupRecord> groups,
        IReadOnlyList<ReportPeriodDto> periods,
        HierPpmLookup lookup,
        IReadOnlyDictionary<string, WeeklyReportFormSettingRecord> settings)
    {
        var rows = new List<WeeklyTargetRowDto>();
        foreach (ModelGroupRecord group in groups)
        foreach (MidGroupRecord mid in group.MidGroups)
        {
            if (string.IsNullOrWhiteSpace(mid.Material) || mid.LineShifts.Count == 0)
                continue;
            string settingKey = "weekly:model:" + NgRateModeSupport.ModelKey(group.Name, mid.Material);
            bool useBGroup = !settings.TryGetValue(settingKey, out WeeklyReportFormSettingRecord? setting) ||
                             !string.Equals(setting.DisplayMode, "Group", StringComparison.OrdinalIgnoreCase);
            useBGroup &= mid.SubGroups.Any(sub => !string.IsNullOrWhiteSpace(sub.Name) && sub.AllLineShifts.Any());

            if (useBGroup)
            {
                foreach (SubGroupRecord sub in mid.SubGroups)
                    EmitSubRows(group.Name, mid.Material, sub, string.Empty, 0, null, periods, lookup, rows);
                continue;
            }

            string modelKey = NgRateModeSupport.ModelKey(group.Name, mid.Material);
            double baseline = lookup.Mid(modelKey.Split("::", 2)[0], mid.Material, BaselineMonthKey);
            rows.Add(new WeeklyTargetRowDto
            {
                RowId = "M::" + modelKey,
                Display = mid.Material,
                RowKind = "group",
                BaselinePpm = BmesReportProjection.PositiveOrNull(baseline),
                TargetPpm = setting?.Target,
                AchievementPercent = Achievement(setting?.Target, LatestPpm(periods, key => lookup.Mid(group.Name, mid.Material, key))),
                PpmByPeriod = BmesReportProjection.Map(periods, key => lookup.Mid(group.Name, mid.Material, key)),
            });
        }
        return rows;
    }

    private static void EmitSubRows(
        string groupName,
        string material,
        SubGroupRecord sub,
        string parentPath,
        int depth,
        string? parentRowId,
        IReadOnlyList<ReportPeriodDto> periods,
        HierPpmLookup lookup,
        ICollection<WeeklyTargetRowDto> rows)
    {
        string nodeName = sub.Name ?? string.Empty;
        string path = string.IsNullOrEmpty(nodeName)
            ? parentPath
            : string.IsNullOrEmpty(parentPath) ? nodeName : parentPath + "::" + nodeName;
        string? currentParent = parentRowId;

        if (!string.IsNullOrEmpty(nodeName) && sub.AllLineShifts.Any())
        {
            bool top = depth <= 1;
            string groupKey = top
                ? ModelGroupPickerHelpers.SubGroupKeyOf(groupName, material, sub)
                : ModelGroupPickerHelpers.SubLeafKeyOf(groupName, material, path);
            string rowId = (top ? "S1::" : $"S{depth}::") + groupKey;
            Func<string, double> valueOf = top
                ? key => lookup.Sub1(groupKey, key)
                : key => lookup.Sub2(groupKey, key);
            rows.Add(new WeeklyTargetRowDto
            {
                RowId = rowId,
                ParentRowId = parentRowId,
                Display = nodeName,
                RowKind = top ? "b-group" : "line-group",
                Depth = depth,
                BaselinePpm = BmesReportProjection.PositiveOrNull(valueOf(BaselineMonthKey)),
                PpmByPeriod = BmesReportProjection.Map(periods, valueOf),
            });
            currentParent = rowId;

            foreach (string lineShift in sub.LineShifts.Where(value => !string.IsNullOrWhiteSpace(value)))
            {
                rows.Add(new WeeklyTargetRowDto
                {
                    RowId = $"LS::{rowId}::{lineShift}",
                    ParentRowId = rowId,
                    Display = lineShift,
                    RowKind = "shift-group",
                    Depth = depth + 1,
                    IsLineShift = true,
                    BaselinePpm = BmesReportProjection.PositiveOrNull(lookup.Ls(lineShift, BaselineMonthKey)),
                    PpmByPeriod = BmesReportProjection.Map(periods, key => lookup.Ls(lineShift, key)),
                });
            }
        }

        int childDepth = string.IsNullOrEmpty(nodeName) ? depth : depth + 1;
        foreach (SubGroupRecord child in sub.SubGroups)
            EmitSubRows(groupName, material, child, path, childDepth, currentParent, periods, lookup, rows);
    }

    private static List<WeeklySummaryRowDto> BuildSummaryRows(
        IReadOnlyList<ModelGroupRecord> groups,
        IReadOnlyList<ReportPeriodDto> periods,
        HierReports hierarchy)
    {
        var rows = new List<WeeklySummaryRowDto>();
        foreach (ModelGroupRecord group in groups)
        {
            string groupId = "G::" + group.Name;
            rows.Add(SummaryRow(groupId, null, 0, "group", group.Name, null, null, group.Name,
                periods, hierarchy.ByGroup, group.Name));
            foreach (MidGroupRecord mid in group.MidGroups)
            {
                if (string.IsNullOrWhiteSpace(mid.Material) || mid.LineShifts.Count == 0)
                    continue;
                string key = NgRateModeSupport.ModelKey(group.Name, mid.Material);
                string modelId = "M::" + key;
                rows.Add(SummaryRow(modelId, groupId, 1, "model", group.Name, mid.Material, null, mid.Material,
                    periods, hierarchy.ByMid, key));

                IReadOnlyList<NgRateReportService.SummaryPivotRow> processRows =
                    hierarchy.ByMid?.GroupSummary.GetValueOrDefault(key) ?? [];
                int ordinal = 0;
                foreach (NgRateReportService.SummaryPivotRow process in processRows.Where(row => !row.IsTotal))
                {
                    rows.Add(new WeeklySummaryRowDto
                    {
                        RowId = $"PT::{key}::{ordinal++}",
                        ParentRowId = modelId,
                        Depth = 2,
                        Level = "process-type",
                        GroupName = group.Name,
                        ModelName = mid.Material,
                        ProcessType = process.ProcessType,
                        Display = process.ProcessType,
                        PpmByPeriod = periods.ToDictionary(
                            period => period.Key,
                            period => process.Ppm.TryGetValue(period.Key, out double value)
                                ? BmesReportProjection.Finite(value)
                                : null,
                            StringComparer.Ordinal),
                    });
                }
            }
        }
        return rows;
    }

    private static WeeklySummaryRowDto SummaryRow(
        string rowId,
        string? parentId,
        int depth,
        string level,
        string? groupName,
        string? modelName,
        string? processType,
        string display,
        IReadOnlyList<ReportPeriodDto> periods,
        NgRateReportService.NgRateReport? report,
        string key)
    {
        NgRateReportService.SummaryPivotRow? total = report?.GroupSummary.GetValueOrDefault(key)?
            .FirstOrDefault(row => row.IsTotal);
        return new WeeklySummaryRowDto
        {
            RowId = rowId,
            ParentRowId = parentId,
            Depth = depth,
            Level = level,
            GroupName = groupName,
            ModelName = modelName,
            ProcessType = processType,
            Display = display,
            PpmByPeriod = periods.ToDictionary(
                period => period.Key,
                period => total is not null && total.Ppm.TryGetValue(period.Key, out double value)
                    ? BmesReportProjection.Finite(value)
                    : null,
                StringComparer.Ordinal),
        };
    }

    private static List<WeeklyDefectRowDto> BuildTopDefects(
        IReadOnlyList<ModelGroupRecord> groups,
        IReadOnlyList<ReportPeriodDto> periods,
        NgRateReportService.LineShiftNgReport report,
        string? sortKey)
    {
        var result = new List<WeeklyDefectRowDto>();
        foreach (ModelGroupRecord group in groups)
        {
            string prefix = group.Name + "::";
            int rank = 0;
            foreach (NgRateReportService.LineShiftNgDetail detail in report.Details
                         .Where(detail => detail.LineShift.StartsWith(prefix, StringComparison.Ordinal))
                         .Select(detail => new { Detail = detail, Reference = SortValue(detail, sortKey, report) })
                         .Where(item => item.Reference > 0)
                         .OrderByDescending(item => item.Reference)
                         .Take(10)
                         .Select(item => item.Detail))
            {
                result.Add(new WeeklyDefectRowDto
                {
                    Rank = ++rank,
                    LineShift = detail.LineShift,
                    ProcessType = detail.ProcessType,
                    ProcessName = detail.ProcessName,
                    NgName = detail.NgName,
                    PpmByPeriod = periods.ToDictionary(
                        period => period.Key,
                        period => BmesReportProjection.Finite(period.Kind switch
                        {
                            "date" => detail.DatePpm.GetValueOrDefault(period.Key),
                            "week" => detail.WeekPpm.GetValueOrDefault(period.Key),
                            "month" => detail.MonthPpm.GetValueOrDefault(period.Key),
                            _ => 0,
                        }),
                        StringComparer.Ordinal),
                });
            }
        }
        return result;
    }

    private static double SortValue(
        NgRateReportService.LineShiftNgDetail detail,
        string? sortKey,
        NgRateReportService.LineShiftNgReport report)
    {
        if (string.IsNullOrEmpty(sortKey))
            return 0;
        int start = report.WeekCols.FindIndex(column => column.Key == sortKey);
        if (start >= 0)
        {
            for (int index = start; index < report.WeekCols.Count; index++)
            {
                double value = detail.WeekPpm.GetValueOrDefault(report.WeekCols[index].Key);
                if (value > 0)
                    return value;
            }
        }
        return detail.DatePpm.GetValueOrDefault(sortKey);
    }

    private static double LatestPpm(IReadOnlyList<ReportPeriodDto> periods, Func<string, double> valueOf)
    {
        foreach (ReportPeriodDto period in periods.Where(period => period.Kind == "week"))
        {
            double value = valueOf(period.Key);
            if (value > 0)
                return value;
        }
        return 0;
    }

    private static double? Achievement(double? target, double actual) =>
        target is > 0 && actual > 0 ? BmesReportProjection.Finite(target.Value / actual * 100d) : null;
}
