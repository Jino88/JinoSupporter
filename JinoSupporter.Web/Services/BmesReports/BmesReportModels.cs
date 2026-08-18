using JinoSupporter.Web.Services.BmesReports.Contracts;

namespace JinoSupporter.Web.Services.BmesReports;

public sealed record BmesReportRequest(
    DateOnly StartDate,
    DateOnly EndDate,
    IReadOnlyList<ModelGroupRecord> Groups,
    IReadOnlyDictionary<string, WeeklyReportFormSettingRecord>? WeeklyFormSettings = null)
{
    public static BmesReportRequest Create(
        DateTime startDate,
        DateTime endDate,
        IReadOnlyList<ModelGroupRecord> groups,
        IReadOnlyDictionary<string, WeeklyReportFormSettingRecord>? weeklyFormSettings = null) =>
        new(DateOnly.FromDateTime(startDate.Date), DateOnly.FromDateTime(endDate.Date), groups, weeklyFormSettings);

    public DateTime StartDateTime => StartDate.ToDateTime(TimeOnly.MinValue);
    public DateTime EndDateTime => EndDate.ToDateTime(TimeOnly.MinValue);
}

public sealed class BmesDailyCalculationSnapshot
{
    public required HierReports Hierarchy { get; init; }
    public required NgRateReportService.NgRateReport TrendByMid { get; init; }
    public required DailyTabDto Tab { get; init; }
    public required string SourceDbPath { get; init; }
}

public sealed class BmesWeeklyCalculationSnapshot
{
    public required WeeklyTabDto Tab { get; init; }
    public required HierReports Hierarchy { get; init; }
    public required NgRateReportService.LineShiftNgReport LineShiftDetails { get; init; }
    public required NgRateReportService.LineShiftNgReport MidDetails { get; init; }
}

public sealed record BmesFCostKpiPeriodValue(
    int ColumnIndex,
    double TotalCost,
    double TotalRate,
    double MainDefectAveragePpm);

public sealed class BmesFCostCalculationSnapshot
{
    public required FCostReport Report { get; init; }
    public NgRateReportService.NgRateReport? NgRateTrendByMid { get; init; }
    public required IReadOnlyList<ModelGroupRecord> ModelGroups { get; init; }
    public BmesFcostRawBreakdownResult? RawMaterialBreakdown { get; init; }
    public FCostRawStatus? RawStatus { get; init; }
    public required IReadOnlyList<BmesFCostKpiPeriodValue> KpiPeriods { get; init; }
    public required FCostDatasetDto Dataset { get; init; }
    public required DateOnly StartDate { get; init; }
    public required DateOnly EndDate { get; init; }
    public IReadOnlyList<ReportIssueDto> Warnings { get; init; } = [];
}

public sealed class BmesReportGenerationResult
{
    public required BmesReportDocumentDto Document { get; init; }
    public required BmesDailyCalculationSnapshot Daily { get; init; }
    public required CauseMonthlyTabDto CauseMonthly { get; init; }
    public required BmesWeeklyCalculationSnapshot Weekly { get; init; }
    public required BmesFCostCalculationSnapshot Fcost { get; init; }
    public FCostCorePartsKpiSnapshot? CorePartsKpi { get; init; }
    public IpgDefectKpiSnapshot? IpgKpi { get; init; }
}
