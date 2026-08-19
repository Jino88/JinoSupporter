using System.Globalization;
using System.Text.Json;
using BmesNgRateStandalone.Services.BmesReports.Contracts;

namespace BmesNgRateStandalone.Services.BmesReports;

public sealed class BmesFCostReportCalculationService(
    FCostService fcost,
    FCostReportService reportService,
    BmesFcostActualService actualService,
    NgRateService ngRate,
    NgRateReportService ngRateReportService,
    NgRateSettingsService ngRateSettings,
    AppPathsService appPaths,
    WebRepository repository)
{
    private const double DefaultBaselineRate = 6.12;
    private const double TrendRemainderMinSharePercent = 0.05;
    private static readonly double[] DefaultTargetRates = [2.0, 0.7];
    private const string DefaultActionItems = "목표불량률 모델 정렬 161016, X526/626 기본값 T/F 체크";

    public async Task<BmesFCostCalculationSnapshot> CalculateAsync(
        BmesReportRequest request,
        NgRateReportService.NgRateReport? dailyTrendByMid,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        DateTime start = request.StartDateTime;
        DateTime end = request.EndDateTime;
        DateTime today = DateTime.Today;
        if (end > today)
            end = today;
        if (start > end)
            throw new ArgumentException("F-COST report start date must be before or equal to end date.", nameof(request));

        DateTime forceFrom = end.AddDays(-2);
        progress?.Report($"Preparing RAW cache. Cached through {forceFrom.AddDays(-1):yyyy-MM-dd}; refreshing {forceFrom:yyyy-MM-dd} - {end:yyyy-MM-dd}.");
        FCostRawBackfillResult backfill = await fcost.BackfillRawAsync(
            start,
            end,
            force: false,
            forceFromDate: forceFrom,
            forceRefreshTtl: TimeSpan.FromMinutes(15),
            progress: progress);
        cancellationToken.ThrowIfCancellationRequested();

        progress?.Report(backfill.FailedDays > 0
            ? $"RAW cache completed with {backfill.FailedDays:N0} failed day(s). Building report from available data..."
            : "RAW cache ready. Building report...");
        FCostReport report = await reportService.GenerateRawRangeReportAsync(
            fcost.GetRawDbPath(),
            start,
            end,
            groupsForSubGroupRollup: request.Groups,
            progress: progress);
        cancellationToken.ThrowIfCancellationRequested();

        var warnings = new List<ReportIssueDto>();
        BmesFcostRawBreakdownResult? rawBreakdown = await LoadRawBreakdownAsync(
            report,
            request.Groups,
            progress,
            warnings,
            cancellationToken);

        NgRateReportService.NgRateReport? trend = dailyTrendByMid;
        if (trend is not null)
        {
            progress?.Report("Reusing Daily NG Rate trend summary.");
        }
        else
        {
            trend = await BuildStandaloneNgRateTrendAsync(request, progress, cancellationToken);
        }

        FCostDatasetDto dataset = ProjectDataset(report, trend, request.Groups, rawBreakdown);
        IReadOnlyList<BmesFCostKpiPeriodValue> kpiPeriods = BuildKpiPeriodValues(report, dataset);
        progress?.Report($"Report ready: {start:yyyy-MM-dd} - {end:yyyy-MM-dd}.");

        return new BmesFCostCalculationSnapshot
        {
            Report = report,
            NgRateTrendByMid = trend,
            ModelGroups = request.Groups.ToArray(),
            RawMaterialBreakdown = rawBreakdown,
            RawStatus = fcost.GetRawStatus(),
            KpiPeriods = kpiPeriods,
            Dataset = dataset,
            StartDate = DateOnly.FromDateTime(start),
            EndDate = DateOnly.FromDateTime(end),
            Warnings = warnings,
        };
    }

    public FCostDatasetDto ProjectDataset(
        FCostReport report,
        NgRateReportService.NgRateReport? trend,
        IReadOnlyList<ModelGroupRecord> groups,
        BmesFcostRawBreakdownResult? rawBreakdown)
    {
        IReadOnlyList<ReportPeriodDto> periods = BmesReportProjection.FromFCost(report);
        IReadOnlyList<int> indexes = Enumerable.Range(0, periods.Count).ToArray();
        (double[] totalInput, double[] totalCost) = ComputeTotals(report, indexes);
        double[] totalRates = ComputeTotalRates(report, indexes, totalInput, totalCost);
        List<HierarchyValue> hierarchy = BuildHierarchy(report);
        List<FCostTrendRowDto> trendRows = BuildTrendRows(
            report,
            trend,
            groups,
            periods,
            hierarchy,
            totalInput,
            totalCost);

        return new FCostDatasetDto
        {
            Periods = periods,
            Totals = new FCostTotalsDto
            {
                InputQtyByPeriod = MapByIndex(periods, index => totalInput[index]),
                FcostUsdByPeriod = MapByIndex(periods, index => totalCost[index]),
                RatePercentByPeriod = MapByIndex(periods, index => totalRates[index]),
            },
            TrendRows = trendRows,
            HierarchyRows = hierarchy.Select(row => new FCostHierarchyRowDto
            {
                RowId = row.RowId,
                ParentRowId = row.ParentRowId,
                Depth = row.Depth,
                GroupName = row.GroupName,
                ModelName = row.ModelName,
                NgGroupKey = row.NgGroupKey,
                Display = row.Display,
                MatchedMaterialCount = row.MatchedMaterialCount,
                InputQtyByPeriod = MapByIndex(periods, index => ValueAt(row.InputByCol, index)),
                FcostUsdByPeriod = MapByIndex(periods, index => ValueAt(row.CostByCol, index)),
                SourceRatePercentByPeriod = MapByIndex(periods, index =>
                    ResolveRate(ValueAt(row.SourceRateByCol, index), ValueAt(row.InputByCol, index), ValueAt(row.CostByCol, index))),
            }).ToArray(),
            Materials = report.Materials.Select(material => ProjectMaterial(material, periods)).ToArray(),
            UnmappedMaterials = report.UnmappedMaterials.Select(material => ProjectMaterial(material, periods)).ToArray(),
            RawBreakdown = rawBreakdown is null ? null : ProjectRawBreakdown(rawBreakdown),
            TargetDefectRate = BuildTargetDefectRate(report, trend, groups, periods, hierarchy, totalRates),
        };
    }

    private async Task<BmesFcostRawBreakdownResult?> LoadRawBreakdownAsync(
        FCostReport report,
        IReadOnlyList<ModelGroupRecord> groups,
        IProgress<string>? progress,
        ICollection<ReportIssueDto> warnings,
        CancellationToken cancellationToken)
    {
        List<int> selectedIndexes = LatestPeriodIndexes(report);
        List<BmesFcostRawBreakdownPeriod> periods = BuildRawPeriods(report, selectedIndexes);
        List<BmesFcostRawBreakdownLineShift> lineShifts = BuildRawLineShifts(groups);
        if (periods.Count == 0 || lineShifts.Count == 0)
            return null;

        BmesFcostDbConnection connection = BuildReadOnlyConnection();
        if (string.IsNullOrWhiteSpace(connection.Server) || string.IsNullOrWhiteSpace(connection.UserId))
        {
            warnings.Add(new ReportIssueDto(
                "fcost-raw-breakdown-unavailable",
                "Network DB connection is not configured for FCOST raw material breakdown.",
                "fcost",
                false));
            return null;
        }

        try
        {
            progress?.Report($"Loading network DB FCOST raw material breakdown ({periods.Count:N0} period(s), {lineShifts.Count:N0} line shift(s))...");
            BmesFcostRawBreakdownResult result = await actualService.FetchRawBreakdownAsync(new BmesFcostRawBreakdownQuery
            {
                Connection = connection,
                Fact = "GN",
                Plant = "3200",
                Periods = periods,
                LineShifts = lineShifts,
            });
            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report($"Network DB FCOST raw material breakdown ready: {result.Rows.Count:N0} row(s).");
            return result;
        }
        catch (Exception)
        {
            warnings.Add(new ReportIssueDto(
                "fcost-raw-breakdown-unavailable",
                "Network DB FCOST raw material breakdown is unavailable.",
                "fcost",
                true));
            progress?.Report("[WARN] Network DB FCOST raw material breakdown unavailable.");
            return null;
        }
    }

    private async Task<NgRateReportService.NgRateReport?> BuildStandaloneNgRateTrendAsync(
        BmesReportRequest request,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        if (!ngRateSettings.IsCredentialsConfigured || !ngRateSettings.IsNgRateStorageConfigured)
            return null;
        (Dictionary<string, IReadOnlyList<string>> mapping, List<string> groupNames) = BuildNgMapping(request.Groups);
        if (mapping.Count == 0 || groupNames.Count == 0)
            return null;
        try
        {
            progress?.Report("Building NG Rate trend for registered models...");
            string? dbPath = await ngRate.GetOrFetchAsync(request.StartDateTime, request.EndDateTime, progress);
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(dbPath))
                return null;
            return await ngRateReportService.GenerateSummaryReportAsync(
                dbPath,
                mapping,
                groupNames,
                progress,
                request.StartDateTime,
                request.EndDateTime,
                weightedGroupSummary: true);
        }
        catch (Exception ex)
        {
            progress?.Report("[WARN] NG Rate trend skipped: " + SafeMessage(ex));
            return null;
        }
    }

    private TargetDefectRateDto BuildTargetDefectRate(
        FCostReport report,
        NgRateReportService.NgRateReport? trend,
        IReadOnlyList<ModelGroupRecord> groups,
        IReadOnlyList<ReportPeriodDto> periods,
        IReadOnlyList<HierarchyValue> hierarchy,
        IReadOnlyList<double> totalRates)
    {
        Dictionary<string, double> targetRates = ReadTargetRates();
        IReadOnlyList<string> actions = (repository.GetSetting("FCostWeeklyReport:ActionItems") ?? DefaultActionItems)
            .Split(["\r\n", "\n"], StringSplitOptions.None)
            .Select(value => value.Trim())
            .Where(value => value.Length > 0)
            .ToArray();
        var rows = new List<TargetDefectRateRowDto>();

        if (trend is not null)
        {
            foreach (ModelGroupRecord group in groups.OrderBy(group => group.SortOrder).ThenBy(group => group.Name, StringComparer.Ordinal))
            foreach (MidGroupRecord mid in group.MidGroups)
            {
                if (string.IsNullOrWhiteSpace(mid.Material))
                    continue;
                string trendKey = NgRateModeSupport.ModelKey(group.Name, mid.Material);
                Dictionary<string, double?> actual = periods.ToDictionary(
                    period => period.Key,
                    period => BmesReportProjection.Finite(LookupTrendPpm(trend, report, trendKey, period.SourceIndex ?? -1)),
                    StringComparer.Ordinal);
                double latest = periods.Where(period => period.Kind == "week")
                    .Select(period => actual.GetValueOrDefault(period.Key) ?? 0)
                    .FirstOrDefault(value => value > 0);
                double? targetPpm = targetRates.TryGetValue(mid.Material, out double target) ? target : null;
                rows.Add(new TargetDefectRateRowDto
                {
                    ModelName = mid.Material,
                    TargetPpm = targetPpm,
                    AchievementPercent = targetPpm is > 0 && latest > 0
                        ? BmesReportProjection.Finite(targetPpm.Value / latest * 100d)
                        : null,
                    ActualPpmByPeriod = actual,
                });
            }
        }
        else
        {
            foreach (HierarchyValue row in hierarchy.Where(row => row.Depth == 0))
            {
                Dictionary<string, double?> actual = periods.ToDictionary(
                    period => period.Key,
                    period => BmesReportProjection.Finite(
                        ResolveRate(
                            ValueAt(row.SourceRateByCol, period.SourceIndex ?? -1),
                            ValueAt(row.InputByCol, period.SourceIndex ?? -1),
                            ValueAt(row.CostByCol, period.SourceIndex ?? -1)) * 10_000d),
                    StringComparer.Ordinal);
                double latest = periods.Where(period => period.Kind == "week")
                    .Select(period => actual.GetValueOrDefault(period.Key) ?? 0)
                    .FirstOrDefault(value => value > 0);
                double? targetPpm = targetRates.TryGetValue(row.Display, out double target) ? target : null;
                rows.Add(new TargetDefectRateRowDto
                {
                    ModelName = row.Display,
                    TargetPpm = targetPpm,
                    AchievementPercent = targetPpm is > 0 && latest > 0
                        ? BmesReportProjection.Finite(targetPpm.Value / latest * 100d)
                        : null,
                    ActualPpmByPeriod = actual,
                });
            }
        }

        return new TargetDefectRateDto
        {
            DefaultBaselineRatePercent = ReadDoubleSetting("FCostWeeklyReport:BaselineRate", DefaultBaselineRate),
            DefaultTargetRatePercent =
            [
                ReadDoubleSetting("FCostWeeklyReport:FirstTargetRate", DefaultTargetRates[0]),
                ReadDoubleSetting("FCostWeeklyReport:SecondTargetRate", DefaultTargetRates[1]),
            ],
            ActionItems = actions,
            Rows = rows,
            TotalActualPpmByPeriod = periods.ToDictionary(
                period => period.Key,
                period => period.SourceIndex is int index && index >= 0 && index < totalRates.Count
                    ? BmesReportProjection.Finite(totalRates[index] * 10_000d)
                    : null,
                StringComparer.Ordinal),
        };
    }

    private static IReadOnlyList<BmesFCostKpiPeriodValue> BuildKpiPeriodValues(
        FCostReport report,
        FCostDatasetDto dataset)
    {
        var result = new List<BmesFCostKpiPeriodValue>();
        for (int index = 0; index < report.Columns.Count; index++)
        {
            if (report.Columns[index].Kind is not (FCostPeriodKind.Week or FCostPeriodKind.Month))
                continue;
            string periodKey = dataset.Periods[index].Key;
            double[] modelPpm = dataset.TrendRows
                .Where(row => row.ParentRowId is null && row.Depth == 0 &&
                              !string.Equals(row.ModelName, "기타_누락", StringComparison.Ordinal))
                .Select(row => row.NgPpmByPeriod.GetValueOrDefault(periodKey) ?? 0)
                .Where(ppm => ppm > 0)
                .ToArray();
            result.Add(new BmesFCostKpiPeriodValue(
                index,
                report.Total?.Cost?.GetCol(index + 1) ?? 0,
                report.Total?.Rate?.GetCol(index + 1) ?? 0,
                modelPpm.Length > 0 ? modelPpm.Average() : 0));
        }
        return result;
    }

    private static List<FCostTrendRowDto> BuildTrendRows(
        FCostReport report,
        NgRateReportService.NgRateReport? trend,
        IReadOnlyList<ModelGroupRecord> groups,
        IReadOnlyList<ReportPeriodDto> periods,
        IReadOnlyList<HierarchyValue> hierarchy,
        IReadOnlyList<double> totalInput,
        IReadOnlyList<double> totalCost)
    {
        var result = new List<FCostTrendRowDto>();
        var emitted = new HashSet<string>(StringComparer.Ordinal);
        foreach (ModelGroupRecord selectedGroup in groups
                     .OrderBy(group => group.SortOrder)
                     .ThenBy(group => group.Name, StringComparer.Ordinal))
        {
            if (string.IsNullOrWhiteSpace(selectedGroup.Name) || !emitted.Add(selectedGroup.Name))
                continue;
            FCostSubGroupAggregate[] groupRows = report.SubGroups
                .Where(row => string.Equals(row.GroupName, selectedGroup.Name, StringComparison.Ordinal))
                .ToArray();
            if (groupRows.Length == 0)
                continue;
            int count = periods.Count;
            var input = new double[count];
            var cost = new double[count];
            foreach (FCostSubGroupAggregate source in groupRows)
            for (int index = 0; index < count; index++)
            {
                input[index] += ValueAt(source.InputByCol, index);
                cost[index] += ValueAt(source.CostByCol, index);
            }
            if (!input.Any(value => value > 0) && !cost.Any(value => value > 0))
                continue;

            result.Add(new FCostTrendRowDto
            {
                RowId = "G::" + selectedGroup.Name,
                Depth = 0,
                GroupName = selectedGroup.Name,
                ModelName = selectedGroup.Name,
                NgGroupKey = selectedGroup.Name,
                InputQtyByPeriod = MapByIndex(periods, index => input[index]),
                FcostUsdByPeriod = MapByIndex(periods, index => cost[index]),
                NgPpmByPeriod = MapByIndex(periods, index => trend is null ? 0 : LookupTrendPpm(trend, report, selectedGroup.Name, index)),
                FcostSharePercentByPeriod = MapByIndex(periods, index =>
                    totalCost[index] > 0 ? cost[index] / totalCost[index] * 100d : 0),
            });
        }

        if (result.Count == 0)
        {
            foreach (HierarchyValue row in hierarchy
                         .Where(row => row.Depth == 0)
                         .Where(row => HasTrendData(row.InputByCol, row.CostByCol))
                         .OrderBy(row => row.Display, StringComparer.Ordinal))
            {
                result.Add(new FCostTrendRowDto
                {
                    RowId = row.RowId,
                    Depth = 0,
                    GroupName = string.Empty,
                    ModelName = row.Display,
                    NgGroupKey = row.NgGroupKey,
                    InputQtyByPeriod = MapByIndex(periods, index => ValueAt(row.InputByCol, index)),
                    FcostUsdByPeriod = MapByIndex(periods, index => ValueAt(row.CostByCol, index)),
                    NgPpmByPeriod = MapByIndex(periods, index => trend is null ? 0 : LookupTrendPpm(trend, report, row.NgGroupKey, index)),
                    FcostSharePercentByPeriod = MapByIndex(periods, index =>
                        totalCost[index] > 0 ? ValueAt(row.CostByCol, index) / totalCost[index] * 100d : 0),
                });
            }
        }

        AppendRemainderTrendRow(result, periods, totalInput, totalCost);
        return result;
    }

    private static void AppendRemainderTrendRow(
        ICollection<FCostTrendRowDto> rows,
        IReadOnlyList<ReportPeriodDto> periods,
        IReadOnlyList<double> totalInput,
        IReadOnlyList<double> totalCost)
    {
        if (rows.Count == 0 || periods.Count == 0)
            return;
        var remainderInput = new double[periods.Count];
        var remainderCost = new double[periods.Count];
        bool meaningful = false;
        for (int index = 0; index < periods.Count; index++)
        {
            string key = periods[index].Key;
            double shownInput = rows.Sum(row => row.InputQtyByPeriod.GetValueOrDefault(key) ?? 0);
            double shownCost = rows.Sum(row => row.FcostUsdByPeriod.GetValueOrDefault(key) ?? 0);
            double inputDifference = totalInput[index] - shownInput;
            double costDifference = totalCost[index] - shownCost;
            if (inputDifference < 0 && Math.Abs(inputDifference) < 0.5)
                inputDifference = 0;
            if (costDifference < 0 && Math.Abs(costDifference) < 0.5)
                costDifference = 0;
            remainderInput[index] = Math.Max(0, inputDifference);
            remainderCost[index] = Math.Max(0, costDifference);
            if (totalCost[index] > 0 &&
                remainderCost[index] / totalCost[index] * 100d >= TrendRemainderMinSharePercent)
                meaningful = true;
        }
        if (!meaningful)
            return;

        rows.Add(new FCostTrendRowDto
        {
            RowId = "M::__BMES_REMAINDER__",
            Depth = 0,
            GroupName = string.Empty,
            ModelName = "기타_누락",
            NgGroupKey = string.Empty,
            InputQtyByPeriod = MapByIndex(periods, index => remainderInput[index]),
            FcostUsdByPeriod = MapByIndex(periods, index => remainderCost[index]),
            NgPpmByPeriod = MapByIndex(periods, _ => 0),
            FcostSharePercentByPeriod = MapByIndex(periods, index =>
                totalCost[index] > 0 ? remainderCost[index] / totalCost[index] * 100d : 0),
        });
    }

    private static bool HasTrendData(IReadOnlyList<double> input, IReadOnlyList<double> cost) =>
        input.Any(value => value > 0) || cost.Any(value => value > 0);

    private static (double[] Input, double[] Cost) ComputeTotals(FCostReport report, IReadOnlyList<int> indexes)
    {
        var input = new double[indexes.Count];
        var cost = new double[indexes.Count];
        if (report.Total is not null)
        {
            for (int position = 0; position < indexes.Count; position++)
            {
                input[position] = report.Total.Input?.GetCol(indexes[position] + 1) ?? 0;
                cost[position] = report.Total.Cost?.GetCol(indexes[position] + 1) ?? 0;
            }
            if (input.Any(value => value > 0) && cost.Any(value => value > 0))
                return (input, cost);
            Array.Clear(input);
            Array.Clear(cost);
        }

        foreach (FCostMaterialBlock material in report.Materials)
        for (int position = 0; position < indexes.Count; position++)
        {
            input[position] += material.Input?.GetCol(indexes[position] + 1) ?? 0;
            cost[position] += material.Cost?.GetCol(indexes[position] + 1) ?? 0;
        }
        if (input.Any(value => value > 0) && cost.Any(value => value > 0))
            return (input, cost);
        Array.Clear(input);
        Array.Clear(cost);

        foreach (FCostSubGroupAggregate subgroup in report.SubGroups)
        for (int position = 0; position < indexes.Count; position++)
        {
            input[position] += ValueAt(subgroup.InputByCol, indexes[position]);
            cost[position] += ValueAt(subgroup.CostByCol, indexes[position]);
        }
        return (input, cost);
    }

    private static double[] ComputeTotalRates(
        FCostReport report,
        IReadOnlyList<int> indexes,
        IReadOnlyList<double> input,
        IReadOnlyList<double> cost)
    {
        var result = new double[indexes.Count];
        for (int position = 0; position < indexes.Count; position++)
        {
            int index = indexes[position];
            double numerator = 0;
            double sourceInput = 0;
            foreach (FCostSubGroupAggregate subgroup in report.SubGroups)
            {
                double weight = ValueAt(subgroup.InputByCol, index);
                if (weight <= 0)
                    continue;
                numerator += ValueAt(subgroup.SourceRateByCol, index) * weight;
                sourceInput += weight;
            }
            result[position] = sourceInput > 0
                ? numerator / sourceInput
                : input[position] > 0 ? cost[position] / input[position] * 100d : 0;
        }
        return result;
    }

    private static List<HierarchyValue> BuildHierarchy(FCostReport report)
    {
        var rows = new List<HierarchyValue>();
        int count = report.Columns.Count;
        foreach (var midRows in report.SubGroups
                     .GroupBy(row => new { row.GroupName, row.MidGroupName })
                     .OrderBy(row => row.Key.GroupName, StringComparer.Ordinal)
                     .ThenBy(row => row.Key.MidGroupName, StringComparer.Ordinal))
        {
            string midId = $"M::{midRows.Key.GroupName}::{midRows.Key.MidGroupName}";
            HierarchyValue mid = Aggregate(midRows, count, midId, null, midRows.Key.MidGroupName,
                midRows.Key.GroupName, midRows.Key.MidGroupName,
                $"{midRows.Key.GroupName}::{midRows.Key.MidGroupName}", 0);
            rows.Add(mid);

            foreach (var subRows in midRows
                         .Select(row => new { Row = row, Segments = ParseSubPath(row.SubGroupKey) })
                         .GroupBy(item => item.Segments.Length > 0 ? item.Segments[0] : item.Row.Display, StringComparer.Ordinal))
            {
                string subId = $"S1::{midId}::{subRows.Key}";
                HierarchyValue sub = Aggregate(subRows.Select(item => item.Row), count, subId, midId, subRows.Key,
                    midRows.Key.GroupName, midRows.Key.MidGroupName,
                    $"{midRows.Key.GroupName}::{midRows.Key.MidGroupName}::{subRows.Key}", 1);
                rows.Add(sub);
                foreach (var item in subRows.Where(item => item.Segments.Length > 1))
                {
                    rows.Add(new HierarchyValue(
                        "S2::" + item.Row.SubGroupKey,
                        subId,
                        item.Segments[^1],
                        midRows.Key.GroupName,
                        midRows.Key.MidGroupName,
                        item.Row.SubGroupKey.Replace(" / ", "::", StringComparison.Ordinal),
                        2,
                        item.Row.MatchedMaterialCount,
                        item.Row.InputByCol,
                        item.Row.CostByCol,
                        item.Row.SourceRateByCol));
                }
            }
        }
        return rows;
    }

    private static HierarchyValue Aggregate(
        IEnumerable<FCostSubGroupAggregate> source,
        int count,
        string id,
        string? parentId,
        string display,
        string groupName,
        string modelName,
        string ngKey,
        int depth)
    {
        var input = new double[count];
        var cost = new double[count];
        var rateNumerator = new double[count];
        var rateInput = new double[count];
        int materials = 0;
        foreach (FCostSubGroupAggregate row in source)
        {
            materials += row.MatchedMaterialCount;
            for (int index = 0; index < count; index++)
            {
                double weight = ValueAt(row.InputByCol, index);
                input[index] += weight;
                cost[index] += ValueAt(row.CostByCol, index);
                rateNumerator[index] += ValueAt(row.SourceRateByCol, index) * weight;
                rateInput[index] += weight;
            }
        }
        var rates = new double[count];
        for (int index = 0; index < count; index++)
            rates[index] = rateInput[index] > 0 ? rateNumerator[index] / rateInput[index] : 0;
        return new HierarchyValue(id, parentId, display, groupName, modelName, ngKey, depth, materials, input, cost, rates);
    }

    private static FCostMaterialRowDto ProjectMaterial(FCostMaterialBlock material, IReadOnlyList<ReportPeriodDto> periods) => new()
    {
        DisplayName = material.DisplayName,
        ProductGroup = NullIfEmpty(material.ProductGroup),
        ModelNo = NullIfEmpty(material.ModelNo),
        Material = NullIfEmpty(material.Material),
        Verid = NullIfEmpty(material.Verid),
        InputQtyByPeriod = MapByIndex(periods, index => material.Input?.GetCol(index + 1) ?? 0),
        FcostUsdByPeriod = MapByIndex(periods, index => material.Cost?.GetCol(index + 1) ?? 0),
        SourceRatePercentByPeriod = MapByIndex(periods, index =>
            ResolveRate(material.Rate?.GetCol(index + 1) ?? 0, material.Input?.GetCol(index + 1) ?? 0, material.Cost?.GetCol(index + 1) ?? 0)),
    };

    private static FCostRawBreakdownDto ProjectRawBreakdown(BmesFcostRawBreakdownResult source) => new()
    {
        SourceTable = source.SourceTable,
        NameSource = source.NameSource,
        WarningMessage = NullIfEmpty(source.WarningMessage),
        Periods = source.Periods.Select((period, index) => new ReportPeriodDto
        {
            Key = period.Kind.Equals("Week", StringComparison.OrdinalIgnoreCase)
                ? "W:" + period.Key.Replace("-", string.Empty, StringComparison.Ordinal)
                : "M:" + period.Key.Replace("-", string.Empty, StringComparison.Ordinal),
            Kind = period.Kind.Equals("Week", StringComparison.OrdinalIgnoreCase) ? "week" : "month",
            Header = period.Header,
            SortOrder = index,
            StartDate = DateOnly.FromDateTime(period.StartDate),
            EndDateExclusive = DateOnly.FromDateTime(period.EndDateExclusive),
        }).ToArray(),
        ExchangeRates = source.ExchangeRates.Select(rate => new FCostExchangeRateDto
        {
            PeriodKey = NormalizeRawPeriodKey(rate.PeriodKey),
            StandardDate = DateOnly.FromDateTime(rate.StandardDate),
            KrwPerUsd = RoundExchange(rate.KrwPerUsd),
            KrwPerVnd = RoundExchange(rate.KrwPerVnd),
        }).ToArray(),
        Rows = source.Rows.Select(row => new FCostRawMaterialRowDto
        {
            GroupName = row.GroupName,
            ModelName = row.ModelName,
            MaterialCode = row.MaterialCode,
            MaterialName = row.MaterialName,
            FcostVndByPeriod = row.FCostByPeriod.ToDictionary(
                item => NormalizeRawPeriodKey(item.Key), item => (decimal?)item.Value, StringComparer.Ordinal),
            EquivalentQtyByPeriod = row.EquivalentQtyByPeriod.ToDictionary(
                item => NormalizeRawPeriodKey(item.Key), item => (decimal?)item.Value, StringComparer.Ordinal),
            PriceByPeriod = row.PriceByPeriod.ToDictionary(
                item => NormalizeRawPeriodKey(item.Key),
                item => (FCostRawPriceDto?)new FCostRawPriceDto
                {
                    UnitPrice = item.Value.UnitPrice,
                    Currency = NullIfEmpty(item.Value.Currency),
                    PriceUnit = NullIfEmpty(item.Value.PriceUnit),
                    UnitPriceVnd = item.Value.UnitPriceVnd,
                    IsMixed = item.Value.IsMixed,
                },
                StringComparer.Ordinal),
            TotalFcostVnd = row.TotalFCostVnd,
            SourceRows = row.SourceRows,
        }).ToArray(),
    };

    private BmesFcostDbConnection BuildReadOnlyConnection()
    {
        var config = appPaths.Current;
        return new BmesFcostDbConnection
        {
            Server = config.AdminDbQueryServer,
            Port = config.AdminDbQueryPort > 0 ? config.AdminDbQueryPort : 1430,
            Database = string.IsNullOrWhiteSpace(config.AdminDbQueryDatabase) ||
                       string.Equals(config.AdminDbQueryDatabase, "master", StringComparison.OrdinalIgnoreCase)
                ? "BMES_LIV"
                : config.AdminDbQueryDatabase,
            UserId = config.AdminDbQueryUserId,
            Password = config.AdminDbQueryPassword,
            TimeoutSeconds = Math.Clamp(Math.Max(config.AdminDbQueryTimeoutSeconds, 300), 1, 600),
            Encrypt = config.AdminDbQueryEncrypt,
            TrustServerCertificate = config.AdminDbQueryTrustServerCertificate,
        };
    }

    private static List<int> LatestPeriodIndexes(FCostReport report)
    {
        IEnumerable<int> indexes = Enumerable.Range(0, report.Columns.Count);
        return indexes.Where(index => report.Columns[index].Kind == FCostPeriodKind.Week)
            .OrderByDescending(index => SortKey(report.Columns[index])).Take(4)
            .Concat(indexes.Where(index => report.Columns[index].Kind == FCostPeriodKind.Month)
                .OrderByDescending(index => SortKey(report.Columns[index])).Take(3))
            .ToList();
    }

    private static List<BmesFcostRawBreakdownPeriod> BuildRawPeriods(FCostReport report, IReadOnlyList<int> indexes)
    {
        var result = new List<BmesFcostRawBreakdownPeriod>();
        foreach (int index in indexes)
        {
            FCostColumnMeta column = report.Columns[index];
            string digits = new((column.PDate + column.Code).Where(char.IsDigit).ToArray());
            if (column.Kind == FCostPeriodKind.Week && digits.Length >= 6 &&
                int.TryParse(digits[..4], out int year) && int.TryParse(digits.Substring(4, 2), out int week) &&
                week >= 1 && week <= ISOWeek.GetWeeksInYear(year))
            {
                DateTime start = ISOWeek.ToDateTime(year, week, DayOfWeek.Monday).Date;
                result.Add(new BmesFcostRawBreakdownPeriod(result.Count, $"{year:0000}-{week:00}", column.Header, "Week", start, start.AddDays(7)));
            }
            else if (column.Kind == FCostPeriodKind.Month && digits.Length >= 6 &&
                     int.TryParse(digits[..4], out year) && int.TryParse(digits.Substring(4, 2), out int month) && month is >= 1 and <= 12)
            {
                DateTime start = new(year, month, 1);
                result.Add(new BmesFcostRawBreakdownPeriod(result.Count, $"{year:0000}{month:00}", column.Header, "Month", start, start.AddMonths(1)));
            }
        }
        return result;
    }

    private static List<BmesFcostRawBreakdownLineShift> BuildRawLineShifts(IReadOnlyList<ModelGroupRecord> groups)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<BmesFcostRawBreakdownLineShift>();
        foreach (ModelGroupRecord group in groups)
        foreach (MidGroupRecord mid in group.MidGroups)
        foreach (string raw in mid.LineShifts)
        {
            string lineShift = raw.Trim();
            if (lineShift.Length == 0 || !seen.Add(lineShift))
                continue;
            result.Add(new BmesFcostRawBreakdownLineShift(
                string.IsNullOrWhiteSpace(group.Name) ? "-" : group.Name.Trim(),
                string.IsNullOrWhiteSpace(mid.Material) ? group.Name.Trim() : mid.Material.Trim(),
                lineShift));
        }
        return result;
    }

    private static (Dictionary<string, IReadOnlyList<string>>, List<string>) BuildNgMapping(IReadOnlyList<ModelGroupRecord> groups)
    {
        var mapping = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var names = new List<string>();
        void Add(string lineShift, string name)
        {
            if (string.IsNullOrWhiteSpace(lineShift) || string.IsNullOrWhiteSpace(name))
                return;
            if (!mapping.TryGetValue(lineShift, out List<string>? list))
                mapping[lineShift] = list = [];
            if (!list.Contains(name, StringComparer.Ordinal))
                list.Add(name);
            if (!names.Contains(name, StringComparer.Ordinal))
                names.Add(name);
        }
        foreach (ModelGroupRecord group in groups)
        foreach (MidGroupRecord mid in group.MidGroups)
        foreach (string lineShift in mid.LineShifts)
        {
            Add(lineShift, group.Name);
            Add(lineShift, NgRateModeSupport.ModelKey(group.Name, mid.Material));
        }
        return (mapping.ToDictionary(item => item.Key, item => (IReadOnlyList<string>)item.Value, StringComparer.Ordinal), names);
    }

    private Dictionary<string, double> ReadTargetRates()
    {
        string? json = repository.GetSetting("FCostWeeklyReport:ModelTargetRates");
        if (string.IsNullOrWhiteSpace(json))
            return new Dictionary<string, double>(StringComparer.Ordinal);
        try
        {
            return JsonSerializer.Deserialize<Dictionary<string, double>>(json)?
                       .Where(item => !string.IsNullOrWhiteSpace(item.Key))
                       .ToDictionary(item => item.Key.Trim(), item => item.Value, StringComparer.Ordinal)
                   ?? new Dictionary<string, double>(StringComparer.Ordinal);
        }
        catch
        {
            return new Dictionary<string, double>(StringComparer.Ordinal);
        }
    }

    private double ReadDoubleSetting(string key, double fallback)
    {
        string? value = repository.GetSetting(key);
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ? parsed : fallback;
    }

    private static double LookupTrendPpm(
        NgRateReportService.NgRateReport trend,
        FCostReport report,
        string groupKey,
        int columnIndex)
    {
        if (columnIndex < 0 || columnIndex >= report.Columns.Count)
            return 0;
        FCostColumnMeta column = report.Columns[columnIndex];
        foreach (string periodKey in TrendPeriodCandidates(column))
        foreach (string key in TrendGroupCandidates(groupKey))
        {
            NgRateReportService.SummaryPivotRow? total = trend.GroupSummary.GetValueOrDefault(key)?
                .FirstOrDefault(row => row.IsTotal);
            if (total?.Ppm.TryGetValue(periodKey, out double value) == true)
                return value;
        }
        return 0;
    }

    private static IEnumerable<string> TrendGroupCandidates(string key)
    {
        string current = key.Trim();
        while (current.Length > 0)
        {
            yield return current;
            int separator = current.LastIndexOf("::", StringComparison.Ordinal);
            if (separator <= 0)
                yield break;
            current = current[..separator];
        }
    }

    private static IEnumerable<string> TrendPeriodCandidates(FCostColumnMeta column)
    {
        string digits = new((column.PDate + column.Code).Where(char.IsDigit).ToArray());
        if (column.Kind == FCostPeriodKind.Week && digits.Length >= 6)
            yield return "W:" + digits[..6];
        if (column.Kind == FCostPeriodKind.Month && digits.Length >= 6)
            yield return "M:" + digits[..6];
        if (column.Kind == FCostPeriodKind.Day && digits.Length >= 8)
            yield return $"{digits[..4]}-{digits.Substring(4, 2)}-{digits.Substring(6, 2)}";
        if (!string.IsNullOrWhiteSpace(column.Code))
            yield return column.Code;
        if (!string.IsNullOrWhiteSpace(column.PDate))
            yield return column.PDate;
    }

    private static Dictionary<string, double?> MapByIndex(IReadOnlyList<ReportPeriodDto> periods, Func<int, double> valueOf) =>
        periods.ToDictionary(
            period => period.Key,
            period => BmesReportProjection.Finite(valueOf(period.SourceIndex ?? period.SortOrder)),
            StringComparer.Ordinal);

    private static double ResolveRate(double sourceRate, double input, double cost) =>
        sourceRate > 0 ? sourceRate : input > 0 ? cost / input * 100d : 0;

    private static double ValueAt(IReadOnlyList<double>? values, int index) =>
        values is not null && index >= 0 && index < values.Count ? values[index] : 0;

    private static string[] ParseSubPath(string key)
    {
        int first = key.IndexOf("::", StringComparison.Ordinal);
        int second = first < 0 ? -1 : key.IndexOf("::", first + 2, StringComparison.Ordinal);
        return second < 0 || second + 2 >= key.Length
            ? []
            : key[(second + 2)..].Split(" / ", StringSplitOptions.None);
    }

    private static long SortKey(FCostColumnMeta column)
    {
        foreach (string candidate in new[] { column.PDate, column.Code })
        {
            string digits = new((candidate ?? string.Empty).Where(char.IsDigit).ToArray());
            if (long.TryParse(digits, out long value))
                return value;
        }
        return column.Index;
    }

    private static string NormalizeRawPeriodKey(string key)
    {
        string digits = new((key ?? string.Empty).Where(char.IsDigit).ToArray());
        return (key ?? string.Empty).Contains('-', StringComparison.Ordinal)
            ? "W:" + digits
            : "M:" + digits;
    }

    private static decimal? RoundExchange(decimal? value) =>
        value is null ? null : Math.Round(value.Value, 2, MidpointRounding.AwayFromZero);

    private static string SafeMessage(Exception exception) =>
        string.IsNullOrWhiteSpace(exception.Message) ? "Source unavailable." : exception.Message;

    private static string? NullIfEmpty(string value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;

    private sealed record HierarchyValue(
        string RowId,
        string? ParentRowId,
        string Display,
        string GroupName,
        string ModelName,
        string NgGroupKey,
        int Depth,
        int MatchedMaterialCount,
        double[] InputByCol,
        double[] CostByCol,
        double[] SourceRateByCol);
}
