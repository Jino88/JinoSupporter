using JinoSupporter.Web.Services.BmesReports.Contracts;

namespace JinoSupporter.Web.Services.BmesReports;

public sealed class BmesReportOrchestrator(
    BmesDailyReportCalculationService dailyService,
    BmesCauseMonthlyCalculationService causeService,
    BmesWeeklyReportCalculationService weeklyService,
    BmesFCostReportCalculationService fcostService,
    BmesKpiReportCalculationService kpiService,
    FCostCorePartsService corePartsService,
    IpgDefectService ipgService)
{
    public const string StageDaily = "Daily";
    public const string StageCause = "원인 비중";
    public const string StageWeekly = "Weekly";
    public const string StageFCost = "F-COST";

    public async Task<BmesReportDocumentDto> GenerateAsync(
        BmesReportRequest request,
        ReportProgressTracker? tracker,
        Action<string>? log,
        CancellationToken cancellationToken) =>
        (await GenerateWithContextAsync(request, tracker, log, cancellationToken)).Document;

    public async Task<BmesReportGenerationResult> GenerateWithContextAsync(
        BmesReportRequest request,
        ReportProgressTracker? tracker,
        Action<string>? log,
        CancellationToken cancellationToken)
    {
        ValidateRequest(request);
        IProgress<string> Stage(string name) =>
            tracker?.For(name, log) ?? new Progress<string>(message => log?.Invoke(message));

        BmesDailyCalculationSnapshot daily = await RunFatalStageAsync(
            StageDaily,
            tracker,
            () => dailyService.CalculateAsync(request, Stage(StageDaily), cancellationToken));
        CauseMonthlyTabDto cause = await RunFatalStageAsync(
            StageCause,
            tracker,
            () => Task.FromResult(causeService.Calculate(request, daily)));
        BmesWeeklyCalculationSnapshot weekly = await RunFatalStageAsync(
            StageWeekly,
            tracker,
            () => weeklyService.CalculateAsync(
                request,
                daily,
                request.WeeklyFormSettings ?? new Dictionary<string, WeeklyReportFormSettingRecord>(StringComparer.Ordinal),
                Stage(StageWeekly),
                cancellationToken));

        DateTime queryDate = request.EndDateTime > DateTime.Today ? DateTime.Today : request.EndDateTime;
        IProgress<string> fcostProgress = Stage(StageFCost);
        log?.Invoke("Starting F-COST and KPI source loading in parallel.");
        Task<FCostCorePartsKpiSnapshot?> corePartsTask = LoadCorePartsAsync(
            request.StartDateTime,
            queryDate,
            fcostProgress,
            cancellationToken);
        Task<IpgDefectKpiSnapshot?> ipgTask = LoadIpgAsync(
            request.StartDateTime,
            queryDate,
            fcostProgress,
            cancellationToken);

        BmesFCostCalculationSnapshot fcost = await RunFatalStageAsync(
            StageFCost,
            tracker,
            () => fcostService.CalculateAsync(request, daily.TrendByMid, fcostProgress, cancellationToken),
            markDone: false);
        await Task.WhenAll(corePartsTask, ipgTask);
        cancellationToken.ThrowIfCancellationRequested();
        FCostCorePartsKpiSnapshot? coreParts = await corePartsTask;
        IpgDefectKpiSnapshot? ipg = await ipgTask;

        var kpiWarnings = new List<ReportIssueDto>();
        KpiTabDto kpi = kpiService.Calculate(fcost, coreParts, ipg, request.EndDate, kpiWarnings);
        tracker?.MarkDone(StageFCost);

        IReadOnlyList<ReportIssueDto> allWarnings = fcost.Warnings.Concat(kpiWarnings).ToArray();
        ReportStatusDto rootStatus = allWarnings.Count == 0
            ? ReportStatusDto.Complete()
            : ReportStatusDto.Partial(allWarnings, "The report was generated with unavailable optional sources.");
        ReportStatusDto fcostStatus = fcost.Warnings.Count == 0
            ? ReportStatusDto.Complete()
            : ReportStatusDto.Partial(fcost.Warnings);
        ReportStatusDto kpiStatus = kpiWarnings.Count == 0
            ? ReportStatusDto.Complete()
            : ReportStatusDto.Partial(kpiWarnings);

        var document = new BmesReportDocumentDto
        {
            GeneratedAtUtc = DateTimeOffset.UtcNow,
            Request = BmesReportProjection.ProjectRequest(request),
            Status = rootStatus,
            Tabs = new BmesReportTabsDto
            {
                Daily = ReportTabEnvelope<DailyTabDto>.Complete(daily.Tab),
                CauseMonthly = ReportTabEnvelope<CauseMonthlyTabDto>.Complete(cause),
                Weekly = ReportTabEnvelope<WeeklyTabDto>.Complete(weekly.Tab),
                Fcost = new ReportTabEnvelope<FCostTabDto>
                {
                    Status = fcostStatus,
                    Data = new FCostTabDto
                    {
                        View = new FCostViewDto { Mode = "regular", AllPeriods = false },
                        Dataset = fcost.Dataset,
                    },
                },
                FcostAll = ReportTabEnvelope<FCostFollowerTabDto>.Complete(new FCostFollowerTabDto
                {
                    View = new FCostViewDto { Mode = "regular", AllPeriods = true, SourceTab = "fcost" },
                }),
                FcostWeekly = ReportTabEnvelope<FCostFollowerTabDto>.Complete(new FCostFollowerTabDto
                {
                    View = new FCostViewDto { Mode = "target-defect-rate", AllPeriods = false, SourceTab = "fcost" },
                }),
                FcostWeeklyAll = ReportTabEnvelope<FCostFollowerTabDto>.Complete(new FCostFollowerTabDto
                {
                    View = new FCostViewDto { Mode = "target-defect-rate", AllPeriods = true, SourceTab = "fcost" },
                }),
                Kpi = new ReportTabEnvelope<KpiTabDto> { Status = kpiStatus, Data = kpi },
            },
        };

        // Serialize once before publication. System.Text.Json rejects NaN/Infinity by
        // default, making a bad numeric projection a fatal generation error instead of
        // publishing an invalid JSON artifact.
        _ = BmesReportJson.SerializeToUtf8Bytes(document);

        return new BmesReportGenerationResult
        {
            Document = document,
            Daily = daily,
            CauseMonthly = cause,
            Weekly = weekly,
            Fcost = fcost,
            CorePartsKpi = coreParts,
            IpgKpi = ipg,
        };
    }

    private async Task<FCostCorePartsKpiSnapshot?> LoadCorePartsAsync(
        DateTime start,
        DateTime queryDate,
        IProgress<string> progress,
        CancellationToken cancellationToken)
    {
        try
        {
            bool refresh = queryDate >= DateTime.Today.AddDays(-2);
            FCostCorePartsBackfillResult pull = await corePartsService.BackfillAsync(
                start,
                queryDate,
                force: false,
                forceFromDate: refresh ? queryDate : null,
                forceRefreshTtl: refresh ? TimeSpan.FromMinutes(15) : null,
                delayMs: 0,
                queryIntervalDays: 21,
                progress: progress);
            cancellationToken.ThrowIfCancellationRequested();
            FCostCorePartsKpiSnapshot? snapshot = corePartsService.GetKpiRangeSnapshot(start, queryDate);
            if (pull.FailedDays > 0 || snapshot is null)
                progress.Report("[WARN] MES072410 KPI snapshot is unavailable.");
            return snapshot;
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch
        {
            progress.Report("[WARN] MES072410 KPI source is unavailable.");
            return null;
        }
    }

    private async Task<IpgDefectKpiSnapshot?> LoadIpgAsync(
        DateTime start,
        DateTime queryDate,
        IProgress<string> progress,
        CancellationToken cancellationToken)
    {
        try
        {
            IpgDefectKpiSnapshot snapshot = await ipgService.FetchKpiRangeAsync(start, queryDate, progress);
            cancellationToken.ThrowIfCancellationRequested();
            return snapshot;
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch
        {
            progress.Report("[WARN] MES050032 KPI source is unavailable.");
            return null;
        }
    }

    private static async Task<T> RunFatalStageAsync<T>(
        string stage,
        ReportProgressTracker? tracker,
        Func<Task<T>> action,
        bool markDone = true)
    {
        tracker?.MarkRunning(stage);
        try
        {
            T result = await action();
            if (markDone)
                tracker?.MarkDone(stage);
            return result;
        }
        catch
        {
            tracker?.MarkFailed(stage);
            throw;
        }
    }

    private static void ValidateRequest(BmesReportRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        if (request.EndDate < request.StartDate)
            throw new ArgumentException("End date must be on or after start date.", nameof(request));
        if (request.Groups.Count == 0)
            throw new ArgumentException("At least one model group must be selected.", nameof(request));
    }
}
