using System.Diagnostics;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Web;
using JinoSupporter.Web.Components.Pages;
using JinoSupporter.Web.Services.BmesReports;
using JinoSupporter.Web.Services.BmesReports.Contracts;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Renders the whole BMES report — all report tabs plus the tab menu — into a SINGLE
/// self-contained static HTML file. The Blazor page just shows that one file in an
/// iframe; tab switching happens entirely inside the HTML (no server round-trips, no
/// live Blazor components).
///
/// Memory safety: every tab body is stored in an inert &lt;template&gt; (parsed but not
/// rendered). Clicking a tab clones just that one body into the DOM and drops the
/// previous one, so only a single report is ever laid out/painted at a time — which is
/// what fixes the Chrome memory blow-up from keeping every report live at once.
///
/// Each tab component is rendered via Blazor's standalone <see cref="HtmlRenderer"/> in
/// "export mode": it auto-generates from the supplied inputs and emits every row/column
/// (a static snapshot cannot expand). The wrapper links the app's own CSS/JS (Chart.js
/// + app.js) from the same origin.
/// </summary>
public sealed class BmesReportHtmlExportService
{
    private readonly IServiceScopeFactory _scopeFactory;
    private readonly ILoggerFactory _loggerFactory;

    /// <summary>Tab key → display label, in menu order.</summary>
    private static readonly (string Key, string Label)[] Tabs =
    {
        ("daily", "Daily"),
        ("weekly", "Weekly"),
        ("kpi", "KPI"),
        ("cause-monthly", "원인 비중"),
        ("fcost", "F-COST"),
        ("fcost-all", "F-COST(전체)"),
        ("fcost-weekly", "목표 불량률"),
        ("fcost-weekly-all", "목표 불량률(전체)"),
    };

    private const string ReportFileName = "report.html";
    private const string ReportJsonFileName = "report.json";
    private const string ReportCacheMetadataFileName = "cache.json";
    private const int ReportCacheFormatVersion = 2;
    private static readonly TimeSpan CompleteReportCacheTtl = TimeSpan.FromMinutes(15);

    /// <summary>Serializes generation so one run cannot clean up another's token folder.</summary>
    private static readonly SemaphoreSlim GenerationLock = new(1, 1);
    private static CompletedReportCacheEntry? _lastCompletedReport;

    /// <summary>Starting value of the viewer toolbar's Minimum PPM filter. The export itself
    /// carries every reason row, so this only decides what the report shows on first open.</summary>
    private const double DefaultReasonPpmThreshold = 500d;

    /// <summary>Progress stages of one generation run, in order. The page builds its
    /// <see cref="ReportProgressTracker"/> from these so stage names cannot drift.</summary>
    public const string StageDaily  = "Daily";
    public const string StageCause  = "원인 비중";
    public const string StageWeekly = "Weekly";
    public const string StageFCost  = "F-COST";

    public static readonly string[] StageNames = [StageDaily, StageCause, StageWeekly, StageFCost];

    public BmesReportHtmlExportService(IServiceScopeFactory scopeFactory, ILoggerFactory loggerFactory)
    {
        _scopeFactory = scopeFactory;
        _loggerFactory = loggerFactory;
    }

    private static string ExportRoot => AppStoragePaths.Combine("_temp", "bmes-report");

    /// <summary>
    /// Resolve a published token without ever treating the token as a path. Published
    /// metadata (disk cache or same-process fallback) and legacy HTML are required, and
    /// the 15-minute TTL is enforced on every view/data request as well as cache lookup.
    /// </summary>
    public BmesReportArtifacts? ResolveReportArtifacts(string token, DateTimeOffset? nowUtc = null) =>
        ResolvePublishedArtifacts(token, nowUtc ?? DateTimeOffset.UtcNow);

    /// <summary>Validate token, publication metadata and TTL, then return legacy HTML.</summary>
    public string? ResolveReportFile(string token)
        => ResolveReportArtifacts(token)?.LegacyHtmlPath;

    /// <summary>Validate token, publication metadata, TTL and schema before returning JSON.</summary>
    public string? ResolveReportJsonFile(string token)
        => ResolveReportArtifacts(token)?.ReportJsonPath;

    public static bool IsValidToken(string token) =>
        !string.IsNullOrEmpty(token) &&
        token.Length is > 0 and <= 64 &&
        token.All(c => (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f'));

    public static bool IsWithinCompleteReportTtl(DateTimeOffset createdAtUtc, DateTimeOffset nowUtc) =>
        createdAtUtc <= nowUtc && nowUtc - createdAtUtc <= CompleteReportCacheTtl;

    public static bool IsCurrentReportContract(int cacheVersion, string schemaVersion, string calculationVersion) =>
        cacheVersion == ReportCacheFormatVersion &&
        string.Equals(schemaVersion, BmesReportContract.SchemaVersion, StringComparison.Ordinal) &&
        string.Equals(calculationVersion, BmesReportContract.CalculationVersion, StringComparison.Ordinal);

    private static bool IsPathWithinRoot(string root, string candidate)
    {
        string rootedPrefix = Path.TrimEndingDirectorySeparator(root) + Path.DirectorySeparatorChar;
        return candidate.StartsWith(rootedPrefix, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Generate the single combined report HTML under a fresh token folder (deleting
    /// previous runs first) and return the new token.
    /// </summary>
    public async Task<string> GenerateAllTabsAsync(
        DateTime start,
        DateTime end,
        IReadOnlyList<ModelGroupRecord> groups,
        ReportProgressTracker? tracker = null,
        Action<string>? log = null)
    {
        // Two overlapping runs used to destroy each other: cleanup ran first and deleted
        // *every* token folder, including the one an in-flight run was about to write into,
        // which surfaced as "Could not find a part of the path ...\report.html".
        await GenerationLock.WaitAsync();
        try
        {
            string cacheKey = BuildCompleteReportCacheKey(start, end, groups);
            if (TryGetCompletedReport(cacheKey, out string cachedToken))
            {
                log?.Invoke("Reusing complete report cache (same period and model selection).");
                foreach (string stageName in StageNames)
                    tracker?.MarkDone(stageName);
                return cachedToken;
            }

            string token = await GenerateAllTabsCoreAsync(cacheKey, start, end, groups, tracker, log);
            return token;
        }
        finally
        {
            GenerationLock.Release();
        }
    }

    private async Task<string> GenerateAllTabsCoreAsync(
        string cacheKey,
        DateTime start,
        DateTime end,
        IReadOnlyList<ModelGroupRecord> groups,
        ReportProgressTracker? tracker,
        Action<string>? log)
    {
        // End-to-end timing for the log line at the bottom of this method. The token and its
        // directory are created after the tabs have rendered, so only the stopwatch starts here.
        var totalSw = Stopwatch.StartNew();

        using IServiceScope scope = _scopeFactory.CreateScope();
        await using var renderer = new HtmlRenderer(scope.ServiceProvider, _loggerFactory);

        var bodies = new Dictionary<string, string>(StringComparer.Ordinal);
        var repository = scope.ServiceProvider.GetRequiredService<WebRepository>();
        var orchestrator = scope.ServiceProvider.GetRequiredService<BmesReportOrchestrator>();
        var request = BmesReportRequest.Create(
            start,
            end,
            groups,
            repository.GetWeeklyReportFormSettings());
        BmesReportGenerationResult generation = await orchestrator.GenerateWithContextAsync(
            request,
            tracker,
            log,
            CancellationToken.None);

        bodies["daily"] = await RenderTabAsync<NgRateForDailyReportPage>(renderer, new()
        {
            ["ShowTitle"] = false,
            ["ShowSetup"] = false,
            // Export every reason row regardless of PPM: the viewer toolbar filters them
            // live, and a row dropped here could never be recovered from a static file.
            ["DailyReasonPpmThreshold"] = 0d,
            ["ExportMode"] = true,
            ["ExportStart"] = start,
            ["ExportEnd"] = end,
            ["ExportGroups"] = groups,
            // The orchestrator already computed this tab and reported its own progress, so the
            // render only formats the snapshot. That is what replaces the OnExportComputed
            // callbacks this branch used to share one tab's result with the next.
            ["ExportCalculationSnapshot"] = generation.Daily,
        }, log);

        // ── 원인 비중 ────────────────────────────────────────────────────────────
        bodies["cause-monthly"] = await RenderTabAsync<BmesCauseMonthlyReportPage>(renderer, new()
        {
            ["ExportMode"] = true,
            ["ExportStart"] = start,
            ["ExportEnd"] = end,
            ["ExportGroups"] = groups,
            ["ExportSharedHierarchy"] = generation.Daily.Hierarchy,
            ["ExportCalculationSnapshot"] = generation.CauseMonthly,
        }, log);

        // ── Weekly ───────────────────────────────────────────────────────────────
        bodies["weekly"] = await RenderTabAsync<NgRateForWeeklyReportPage>(renderer, new()
        {
            ["ShowTitle"] = false,
            ["ShowSetup"] = false,
            ["ExportMode"] = true,
            ["ExportStart"] = start,
            ["ExportEnd"] = end,
            ["ExportGroups"] = groups,
            ["ExportSharedHierarchy"] = generation.Daily.Hierarchy,
            ["ExportCalculationSnapshot"] = generation.Weekly,
        }, log);

        // ── F-COST ×4 ────────────────────────────────────────────────────────────
        // Only the first variant builds the report (RAW backfill + report build + raw
        // material breakdown); the other three render the same snapshot with different
        // display flags.
        foreach ((string key, bool weeklyReport, bool allColumns, bool trendCard) in FCostVariants)
        {
            var parameters = new Dictionary<string, object?>
            {
                ["ShowPageTitle"] = false,
                ["ShowPageHeader"] = false,
                ["ShowSetup"] = false,
                ["ShowFCostWeeklyReport"] = weeklyReport,
                ["ShowAllPeriodColumns"] = allColumns,
                ["ShowRegularFCostDetails"] = !weeklyReport,
                ["ShowMajorTrendCard"] = trendCard,
                ["ExportMode"] = true,
                ["ExportStart"] = start,
                ["ExportEnd"] = end,
                ["ExportGroups"] = groups,
                ["ExportCalculationSnapshot"] = generation.Fcost,
            };

            bodies[key] = await RenderTabAsync<BmesFCostPage>(renderer, parameters, log);
        }

        // ── KPI ──────────────────────────────────────────────────────────────────
        // KPI uses the exact Total FCOST / TOTAL RATE values and visible-model defect
        // rates published by the F-COST leader snapshot, so it must be rendered after
        // the first F-COST export component has finished computing.
        bodies["kpi"] = BuildKpiBody(generation.Fcost, generation.CorePartsKpi, generation.IpgKpi, end);

        string html = BuildCombinedHtml(bodies, DefaultReasonPpmThreshold);
        byte[] reportJson = BmesReportJson.SerializeToUtf8Bytes(generation.Document);
        string token = Guid.NewGuid().ToString("N");
        string dir = Path.Combine(ExportRoot, token);
        try
        {
            Directory.CreateDirectory(dir);
            string htmlTemp = Path.Combine(dir, ReportFileName + ".tmp");
            string jsonTemp = Path.Combine(dir, ReportJsonFileName + ".tmp");
            await File.WriteAllTextAsync(htmlTemp, html, new UTF8Encoding(false));
            await File.WriteAllBytesAsync(jsonTemp, reportJson);
            File.Move(htmlTemp, Path.Combine(dir, ReportFileName));
            File.Move(jsonTemp, Path.Combine(dir, ReportJsonFileName));
            PublishCompletedReport(cacheKey, token);
        }
        catch
        {
            try { if (Directory.Exists(dir)) Directory.Delete(dir, recursive: true); }
            catch { /* best-effort removal; an unpublished folder is never route-visible */ }
            throw;
        }

        // Only a fully published report (HTML + JSON + cache.json) may replace older tokens.
        CleanupOldTokens(token);
        log?.Invoke($"BMES report export complete in {totalSw.ElapsedMilliseconds:N0} ms (tabs={bodies.Count:N0})");

        return token;
    }

    /// <summary>Tab key → the display flags that distinguish the four F-COST views.
    /// The plain F-COST view comes first because it is the one that builds the report.
    ///
    /// <c>TrendCard</c> puts the "주요 모델 불량률 TREND" table above the F-COST tables,
    /// where it belongs. It is gated behind <c>BmesFCostPage.TrendReportOnly</c>, which had
    /// no caller anywhere in the web project, so the card had never reached the export (or
    /// any route) at all — see the 2026-07-22 handoff §12/§14.</summary>
    private static async Task<FCostCorePartsKpiSnapshot?> LoadCorePartsKpiAsync(
        IServiceProvider services,
        DateTime startDate,
        DateTime queryDate,
        IProgress<string> progress)
    {
        try
        {
            var coreParts = services.GetRequiredService<FCostCorePartsService>();
            bool refresh = queryDate >= DateTime.Today.AddDays(-2);
            FCostCorePartsBackfillResult pull = await coreParts.BackfillAsync(
                startDate,
                queryDate,
                force: false,
                forceFromDate: refresh ? queryDate : null,
                forceRefreshTtl: refresh ? TimeSpan.FromMinutes(15) : null,
                delayMs: 0,
                queryIntervalDays: 21,
                progress: progress);
            FCostCorePartsKpiSnapshot? snapshot =
                coreParts.GetKpiRangeSnapshot(startDate, queryDate);
            if (pull.FailedDays > 0 || snapshot is null)
                progress.Report("[WARN] MES072410 KPI snapshot is unavailable.");
            return snapshot;
        }
        catch (Exception ex)
        {
            progress.Report("[WARN] MES072410 KPI: " + ex.Message);
            return null;
        }
    }

    private static async Task<IpgDefectKpiSnapshot?> LoadIpgKpiAsync(
        IServiceProvider services,
        DateTime startDate,
        DateTime queryDate,
        IProgress<string> progress)
    {
        try
        {
            var ipg = services.GetRequiredService<IpgDefectService>();
            return await ipg.FetchKpiRangeAsync(startDate, queryDate, progress);
        }
        catch (Exception ex)
        {
            progress.Report("[WARN] MES050032 KPI: " + ex.Message);
            return null;
        }
    }

    private static string BuildCompleteReportCacheKey(
        DateTime start,
        DateTime end,
        IReadOnlyList<ModelGroupRecord> groups)
    {
        string payload = JsonSerializer.Serialize(new
        {
            Version = 2,
            SchemaVersion = BmesReportContract.SchemaVersion,
            CalculationVersion = BmesReportContract.CalculationVersion,
            Start = start.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            End = end.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            Groups = groups,
        });
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(payload)));
    }

    private static bool TryGetCompletedReport(string cacheKey, out string token)
    {
        DateTimeOffset now = DateTimeOffset.UtcNow;
        if (_lastCompletedReport is { } memoryEntry &&
            string.Equals(memoryEntry.CacheKey, cacheKey, StringComparison.Ordinal) &&
            ResolvePublishedArtifacts(memoryEntry.Token, now) is { IsCurrentContract: true, ReportJsonPath: not null })
        {
            token = memoryEntry.Token;
            return true;
        }

        try
        {
            if (!Directory.Exists(ExportRoot))
            {
                token = string.Empty;
                return false;
            }

            foreach (string directory in Directory
                         .EnumerateDirectories(ExportRoot)
                         .OrderByDescending(Directory.GetLastWriteTimeUtc))
            {
                string candidateToken = Path.GetFileName(directory);
                BmesReportArtifacts? artifacts = ResolvePublishedArtifacts(candidateToken, now);
                if (artifacts is not { IsCurrentContract: true, ReportJsonPath: not null } ||
                    !string.Equals(artifacts.CacheKey, cacheKey, StringComparison.Ordinal))
                {
                    continue;
                }

                _lastCompletedReport = new CompletedReportCacheEntry(
                    ReportCacheFormatVersion,
                    artifacts.CacheKey,
                    artifacts.Token,
                    artifacts.CreatedAtUtc,
                    BmesReportContract.SchemaVersion,
                    BmesReportContract.CalculationVersion,
                    HasLegacyHtml: true,
                    HasReportJson: true);
                token = candidateToken;
                return true;
            }
        }
        catch
        {
            // A cache read must never prevent a fresh report from being generated.
        }

        token = string.Empty;
        return false;
    }

    private static void PublishCompletedReport(string cacheKey, string token)
    {
        var entry = new CompletedReportCacheEntry(
            ReportCacheFormatVersion,
            cacheKey,
            token,
            DateTimeOffset.UtcNow,
            BmesReportContract.SchemaVersion,
            BmesReportContract.CalculationVersion,
            HasLegacyHtml: true,
            HasReportJson: true);
        string directory = Path.Combine(ExportRoot, token);
        string metadataPath = Path.Combine(directory, ReportCacheMetadataFileName);
        string metadataTempPath = metadataPath + ".tmp";
        try
        {
            File.WriteAllText(metadataTempPath, JsonSerializer.Serialize(entry), new UTF8Encoding(false));
            File.Move(metadataTempPath, metadataPath);
        }
        catch (IOException)
        {
            try { if (File.Exists(metadataTempPath)) File.Delete(metadataTempPath); }
            catch { /* best-effort cache metadata cleanup */ }
        }
        catch (UnauthorizedAccessException)
        {
            try { if (File.Exists(metadataTempPath)) File.Delete(metadataTempPath); }
            catch { /* best-effort cache metadata cleanup */ }
        }
        _lastCompletedReport = entry;
    }

    private static BmesReportArtifacts? ResolvePublishedArtifacts(string token, DateTimeOffset now)
    {
        if (!IsValidToken(token)) return null;

        string root = Path.GetFullPath(ExportRoot);
        string directory = Path.GetFullPath(Path.Combine(root, token));
        if (!IsPathWithinRoot(root, directory)) return null;

        string metadataPath = Path.Combine(directory, ReportCacheMetadataFileName);
        string legacyHtmlPath = Path.Combine(directory, ReportFileName);
        if (!File.Exists(legacyHtmlPath)) return null;

        try
        {
            CompletedReportCacheEntry? entry = File.Exists(metadataPath)
                ? ReadCompletedReportEntry(metadataPath)
                : _lastCompletedReport is { } memoryEntry &&
                  string.Equals(memoryEntry.Token, token, StringComparison.Ordinal)
                    ? memoryEntry
                    : null;
            if (entry is null ||
                !string.Equals(entry.Token, token, StringComparison.Ordinal) ||
                !IsWithinCompleteReportTtl(entry.CreatedAtUtc, now) ||
                !entry.HasLegacyHtml)
            {
                return null;
            }

            bool isCurrentContract = IsCurrentReportContract(
                entry.Version,
                entry.SchemaVersion,
                entry.CalculationVersion);
            string reportJsonPath = Path.Combine(directory, ReportJsonFileName);
            string? currentReportJsonPath = isCurrentContract && entry.HasReportJson && File.Exists(reportJsonPath)
                ? reportJsonPath
                : null;

            return new BmesReportArtifacts(
                token,
                legacyHtmlPath,
                currentReportJsonPath,
                metadataPath,
                entry.CacheKey,
                entry.CreatedAtUtc,
                isCurrentContract);
        }
        catch (IOException)
        {
            return null;
        }
        catch (UnauthorizedAccessException)
        {
            return null;
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private static CompletedReportCacheEntry? ReadCompletedReportEntry(string metadataPath) =>
        JsonSerializer.Deserialize<CompletedReportCacheEntry>(File.ReadAllText(metadataPath, Encoding.UTF8));

    private sealed record CompletedReportCacheEntry(
        int Version,
        string CacheKey,
        string Token,
        DateTimeOffset CreatedAtUtc,
        string SchemaVersion = "",
        string CalculationVersion = "",
        bool HasLegacyHtml = false,
        bool HasReportJson = false);

    private static readonly (string Key, bool WeeklyReport, bool AllColumns, bool TrendCard)[] FCostVariants =
    {
        ("fcost",            false, false, true),
        ("fcost-all",        false, true,  true),
        ("fcost-weekly",     true,  false, false),
        ("fcost-weekly-all", true,  true,  false),
    };

    private static async Task<string> RenderTabAsync<TComponent>(
        HtmlRenderer renderer, Dictionary<string, object?> parameters, Action<string>? log = null)
        where TComponent : IComponent =>
        await renderer.Dispatcher.InvokeAsync(async () =>
        {
            var sw = Stopwatch.StartNew();
            var output = await renderer.RenderComponentAsync<TComponent>(
                ParameterView.FromDictionary(parameters));
            string html = output.ToHtmlString();
            log?.Invoke($"Tab {typeof(TComponent).Name} rendered in {sw.ElapsedMilliseconds:N0} ms ({html.Length:N0} chars)");
            return html;
        });

    private static string PlaceholderBody(string key) =>
        $"<div style=\"padding:40px;color:#64748b;font-size:14px;\">'{key}' 탭 HTML 생성은 준비 중입니다.</div>";

    private sealed record KpiPeriodColumn(
        int ColumnIndex,
        string Key,
        string Header,
        DateTime SortDate,
        bool IsMonth);

    private static string? GetKpiMonthGroup(KpiPeriodColumn period) =>
        period.SortDate == DateTime.MaxValue
            ? null
            : "M:" + period.SortDate.ToString("yyyyMM", CultureInfo.InvariantCulture);

    private enum KpiValueKind
    {
        None,
        TotalRate,
        TotalRateAchievement,
        TotalCost,
        CorePartsRate,
        CorePartsRateAchievement,
        CorePartsCost,
        MainDefectAveragePpm,
        MainDefectAchievement,
        IpgDefectAveragePpm,
        IpgDefectAchievement,
    }

    private sealed record KpiLineDefinition(
        string Label,
        KpiValueKind ValueKind = KpiValueKind.None);

    private sealed record KpiDefinition(
        string Name,
        string Type,
        string Baseline,
        string Target,
        double? TargetValue,
        IReadOnlyList<KpiLineDefinition> Lines);

    private static readonly KpiDefinition[] KpiDefinitions =
    [
        new(
            "초과투입 재료비(Main)",
            "법인",
            "1.79%",
            "1%",
            1.0,
            [
                new("실적", KpiValueKind.TotalRate),
                new("달성율", KpiValueKind.TotalRateAchievement),
                new("초과 투입 재료비", KpiValueKind.TotalCost),
                new("매출대비비중"),
            ]),
        new(
            "초과투입 재료비(내재화)",
            "법인",
            "0.95%",
            "0.76%",
            0.76,
            [
                new("실적", KpiValueKind.CorePartsRate),
                new("달성율", KpiValueKind.CorePartsRateAchievement),
                new("초과 투입 재료비", KpiValueKind.CorePartsCost),
                new("매출대비비중"),
            ]),
        new(
            "Main 공정불량 개선율",
            "법인",
            "73,934 ppm",
            "40,000 ppm",
            40_000,
            [
                new("실적", KpiValueKind.MainDefectAveragePpm),
                new("달성율", KpiValueKind.MainDefectAchievement),
            ]),
        new(
            "IPG 공정불량 개선율",
            "법인",
            "1,403 ppm",
            "1,000 ppm",
            1_000,
            [
                new("실적", KpiValueKind.IpgDefectAveragePpm),
                new("달성율", KpiValueKind.IpgDefectAchievement),
            ]),
    ];

    private static string BuildKpiBody(
        BmesFCostCalculationSnapshot? snapshot,
        FCostCorePartsKpiSnapshot? corePartsKpi,
        IpgDefectKpiSnapshot? ipgKpi,
        DateTime reportEnd)
    {
        if (snapshot is null)
        {
            return "<div class=\"kpi-empty\">KPI 데이터를 생성할 수 없습니다.</div>";
        }

        List<KpiPeriodColumn> periods = BuildKpiPeriodColumns(snapshot.Report);
        Dictionary<int, BmesFCostKpiPeriodValue> valuesByColumn =
            snapshot.KpiPeriods.ToDictionary(value => value.ColumnIndex);
        Dictionary<string, FCostCorePartsKpiPeriod> corePartsByPeriod =
            BuildCorePartsKpiPeriodMap(corePartsKpi);
        Dictionary<string, IpgDefectKpiPeriod> ipgByPeriod =
            BuildIpgKpiPeriodMap(ipgKpi, reportEnd.Year);
        string yearLabel = (reportEnd.Year % 100).ToString("00", CultureInfo.InvariantCulture);

        var html = new StringBuilder();
        html.Append("""
            <section class="kpi-report">
              <div class="kpi-report-title">KPI</div>
              <div class="kpi-report-note">F-COST Total, 내재화 Total 및 IPG 공정 평균을 월·주별 KPI 값으로 표시합니다.</div>
              <div class="kpi-report-wrap">
                <table class="kpi-report-table">
                  <thead><tr>
                    <th class="kpi-pin kpi-name">KPI</th>
                    <th class="kpi-pin kpi-type">KPI 종류</th>
                    <th class="kpi-pin kpi-baseline">기준실적</th>
            """);
        html.Append("<th class=\"kpi-pin kpi-target\">")
            .Append(yearLabel)
            .Append("년 목표</th><th class=\"kpi-pin kpi-ytd\">")
            .Append(yearLabel)
            .Append("년 실적</th><th class=\"kpi-pin kpi-line\">구분</th>");

        foreach (KpiPeriodColumn period in periods)
        {
            string? monthGroup = GetKpiMonthGroup(period);
            if (period.IsMonth && monthGroup is not null)
            {
                string encodedHeader = System.Net.WebUtility.HtmlEncode(period.Header);
                html.Append("<th class=\"kpi-period kpi-month\" data-kpi-month=\"")
                    .Append(System.Net.WebUtility.HtmlEncode(monthGroup))
                    .Append("\"><span class=\"kpi-month-head\"><span>")
                    .Append(encodedHeader)
                    .Append("</span><button type=\"button\" class=\"kpi-month-toggle\" data-kpi-month-toggle=\"")
                    .Append(System.Net.WebUtility.HtmlEncode(monthGroup))
                    .Append("\" data-kpi-month-label=\"")
                    .Append(encodedHeader)
                    .Append("\" aria-expanded=\"true\" aria-label=\"")
                    .Append(encodedHeader)
                    .Append(" 주차 접기\" title=\"")
                    .Append(encodedHeader)
                    .Append(" 주차 접기\">−</button></span></th>");
                continue;
            }

            html.Append(period.IsMonth
                    ? "<th class=\"kpi-period kpi-month\""
                    : "<th class=\"kpi-period kpi-week\"");
            if (monthGroup is not null)
            {
                html.Append(" data-kpi-week-month=\"")
                    .Append(System.Net.WebUtility.HtmlEncode(monthGroup))
                    .Append("\"");
            }

            html.Append(">")
                .Append(System.Net.WebUtility.HtmlEncode(period.Header))
                .Append("</th>");
        }

        html.Append("</tr></thead><tbody>");

        foreach (KpiDefinition definition in KpiDefinitions)
        {
            int lineCount = definition.Lines.Count;
            for (int lineIndex = 0; lineIndex < lineCount; lineIndex++)
            {
                KpiLineDefinition line = definition.Lines[lineIndex];
                html.Append("<tr");
                if (lineIndex == 0)
                    html.Append(" class=\"kpi-metric");
                else if (lineIndex == lineCount - 1)
                    html.Append(" class=\"kpi-metric-end");
                if (lineIndex == 0 && lineIndex == lineCount - 1)
                    html.Append(" kpi-metric-end");
                if (lineIndex == 0 || lineIndex == lineCount - 1)
                    html.Append("\"");
                html.Append(">");

                if (lineIndex == 0)
                {
                    string rowSpan = lineCount.ToString(CultureInfo.InvariantCulture);
                    html.Append("<th class=\"kpi-pin kpi-name kpi-group-cell\" rowspan=\"")
                        .Append(rowSpan)
                        .Append("\" scope=\"rowgroup\">")
                        .Append(System.Net.WebUtility.HtmlEncode(definition.Name))
                        .Append("</th><td class=\"kpi-pin kpi-type kpi-group-cell\" rowspan=\"")
                        .Append(rowSpan)
                        .Append("\">")
                        .Append(System.Net.WebUtility.HtmlEncode(definition.Type))
                        .Append("</td>");
                    AppendKpiMergedValueCell(html, "kpi-baseline", definition.Baseline, rowSpan);
                    AppendKpiMergedValueCell(html, "kpi-target", definition.Target, rowSpan);
                }

                string annualValue = FormatKpiAnnualValue(definition, line, ipgKpi);
                html.Append(annualValue == "-"
                        ? "<td class=\"kpi-pin kpi-ytd kpi-pending\">"
                        : "<td class=\"kpi-pin kpi-ytd\">")
                    .Append(System.Net.WebUtility.HtmlEncode(annualValue))
                    .Append("</td>")
                    .Append("<th class=\"kpi-pin kpi-line\" scope=\"row\">")
                    .Append(System.Net.WebUtility.HtmlEncode(line.Label))
                    .Append("</th>");

                foreach (KpiPeriodColumn period in periods)
                {
                    valuesByColumn.TryGetValue(period.ColumnIndex, out var periodValue);
                    corePartsByPeriod.TryGetValue(period.Key, out var corePartsValue);
                    ipgByPeriod.TryGetValue(period.Key, out var ipgValue);
                    string value = FormatKpiPeriodValue(
                        definition,
                        line,
                        periodValue,
                        corePartsValue,
                        ipgValue);
                    html.Append(value == "-"
                            ? "<td class=\"kpi-value kpi-pending\""
                            : "<td class=\"kpi-value\"");
                    string? monthGroup = GetKpiMonthGroup(period);
                    if (!period.IsMonth && monthGroup is not null)
                    {
                        html.Append(" data-kpi-week-month=\"")
                            .Append(System.Net.WebUtility.HtmlEncode(monthGroup))
                            .Append("\"");
                    }

                    html.Append(">")
                        .Append(System.Net.WebUtility.HtmlEncode(value))
                        .Append("</td>");
                }

                html.Append("</tr>");
            }
        }

        if (KpiDefinitions.Length == 0)
        {
            int columnCount = Math.Max(6, periods.Count + 6);
            html.Append("<tr><td class=\"kpi-empty\" colspan=\"")
                .Append(columnCount)
                .Append("\">표시할 KPI 데이터가 없습니다.</td></tr>");
        }

        html.Append("</tbody></table></div></section>");
        return html.ToString();
    }

    private static void AppendKpiMergedValueCell(
        StringBuilder html,
        string cssClass,
        string value,
        string rowSpan)
    {
        html.Append("<td class=\"kpi-pin ")
            .Append(cssClass)
            .Append(" kpi-group-cell");
        if (value == "-")
            html.Append(" kpi-pending");
        html.Append("\" rowspan=\"")
            .Append(rowSpan)
            .Append("\">")
            .Append(System.Net.WebUtility.HtmlEncode(value))
            .Append("</td>");
    }

    private static string FormatKpiPeriodValue(
        KpiDefinition definition,
        KpiLineDefinition line,
        BmesFCostKpiPeriodValue? value,
        FCostCorePartsKpiPeriod? corePartsValue,
        IpgDefectKpiPeriod? ipgValue)
    {
        return line.ValueKind switch
        {
            KpiValueKind.TotalRate =>
                value is null ? "-" : FormatKpiPercent(value.TotalRate, 2),
            KpiValueKind.TotalRateAchievement =>
                value is null
                    ? "-"
                    : FormatKpiAchievement(definition.TargetValue, value.TotalRate, 0),
            KpiValueKind.TotalCost =>
                value is null ? "-" : FormatKpiCurrency(value.TotalCost),
            KpiValueKind.CorePartsRate =>
                FormatKpiNullablePercent(corePartsValue?.TotalRatePercent, 2),
            KpiValueKind.CorePartsRateAchievement =>
                FormatKpiAchievement(
                    definition.TargetValue,
                    corePartsValue?.TotalRatePercent ?? 0,
                    0),
            KpiValueKind.CorePartsCost =>
                FormatKpiNullableCurrency(corePartsValue?.TotalCostUsd),
            KpiValueKind.MainDefectAveragePpm =>
                value is null ? "-" : FormatKpiPpm(value.MainDefectAveragePpm),
            KpiValueKind.MainDefectAchievement =>
                value is null
                    ? "-"
                    : FormatKpiAchievement(
                        definition.TargetValue,
                        value.MainDefectAveragePpm,
                        1),
            KpiValueKind.IpgDefectAveragePpm =>
                FormatKpiNullablePpm(ipgValue?.AveragePpm),
            KpiValueKind.IpgDefectAchievement =>
                FormatKpiAchievement(
                    definition.TargetValue,
                    ipgValue?.AveragePpm ?? 0,
                    1),
            _ => "-",
        };
    }

    private static string FormatKpiAnnualValue(
        KpiDefinition definition,
        KpiLineDefinition line,
        IpgDefectKpiSnapshot? ipgKpi) =>
        line.ValueKind switch
        {
            KpiValueKind.IpgDefectAveragePpm =>
                FormatKpiNullablePpm(ipgKpi?.AnnualAveragePpm),
            KpiValueKind.IpgDefectAchievement =>
                FormatKpiAchievement(
                    definition.TargetValue,
                    ipgKpi?.AnnualAveragePpm ?? 0,
                    1),
            _ => "-",
        };

    private static Dictionary<string, FCostCorePartsKpiPeriod> BuildCorePartsKpiPeriodMap(
        FCostCorePartsKpiSnapshot? snapshot)
    {
        var result = new Dictionary<string, FCostCorePartsKpiPeriod>(StringComparer.Ordinal);
        if (snapshot is null)
            return result;

        foreach (FCostCorePartsKpiPeriod period in snapshot.Periods)
        {
            string? key = CorePartsKpiPeriodKey(period);
            if (key is not null)
                result.TryAdd(key, period);
        }
        return result;
    }

    private static string? CorePartsKpiPeriodKey(FCostCorePartsKpiPeriod period)
    {
        string prefix;
        if (period.Kind.Equals("Week", StringComparison.OrdinalIgnoreCase))
            prefix = "W:";
        else if (period.Kind.Equals("Month", StringComparison.OrdinalIgnoreCase))
            prefix = "M:";
        else
            return null;

        foreach (string candidate in new[] { period.PDate, period.Code })
        {
            string digits = new(candidate.Where(char.IsDigit).ToArray());
            if (digits.Length >= 6)
                return prefix + digits[..6];
        }
        return null;
    }

    private static Dictionary<string, IpgDefectKpiPeriod> BuildIpgKpiPeriodMap(
        IpgDefectKpiSnapshot? snapshot,
        int fallbackYear)
    {
        var result = new Dictionary<string, IpgDefectKpiPeriod>(StringComparer.Ordinal);
        if (snapshot is null)
            return result;

        foreach (IpgDefectKpiPeriod period in snapshot.Periods)
        {
            string? key = IpgKpiPeriodKey(period, fallbackYear);
            if (key is not null)
                result.TryAdd(key, period);
        }
        return result;
    }

    private static string? IpgKpiPeriodKey(
        IpgDefectKpiPeriod period,
        int fallbackYear)
    {
        if (period.Kind.Equals("Week", StringComparison.OrdinalIgnoreCase))
            return IpgWeekKpiPeriodKey(period.Header, fallbackYear);
        if (period.Kind.Equals("Month", StringComparison.OrdinalIgnoreCase))
            return IpgMonthKpiPeriodKey(period.Header, fallbackYear);
        return null;
    }

    private static string? IpgWeekKpiPeriodKey(string header, int fallbackYear)
    {
        string normalized = (header ?? string.Empty).Trim().ToUpperInvariant();
        int markerIndex = normalized.IndexOf('W');
        if (markerIndex < 0)
            return null;

        string yearDigits = new(normalized[..markerIndex].Where(char.IsDigit).ToArray());
        string weekDigits = new(normalized[(markerIndex + 1)..].Where(char.IsDigit).ToArray());
        if (!int.TryParse(
                weekDigits,
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out int week) ||
            week is < 1 or > 54)
        {
            return null;
        }

        int year = fallbackYear;
        if (yearDigits.Length >= 4)
        {
            int.TryParse(
                yearDigits[^4..],
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out year);
        }
        else if (yearDigits.Length == 2 &&
                 int.TryParse(
                     yearDigits,
                     NumberStyles.None,
                     CultureInfo.InvariantCulture,
                     out int shortYear))
        {
            year = 2000 + shortYear;
        }

        return year is >= 1 and <= 9999
            ? $"W:{year:D4}{week:D2}"
            : null;
    }

    private static string? IpgMonthKpiPeriodKey(string header, int fallbackYear)
    {
        string digits = new((header ?? string.Empty).Where(char.IsDigit).ToArray());
        int year = fallbackYear;
        int month;

        if (digits.Length >= 6)
        {
            if (!int.TryParse(
                    digits.AsSpan(0, 4),
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out year) ||
                !int.TryParse(
                    digits.AsSpan(4, 2),
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out month))
            {
                return null;
            }
        }
        else if (digits.Length == 4)
        {
            if (!int.TryParse(
                    digits.AsSpan(0, 2),
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out int shortYear) ||
                !int.TryParse(
                    digits.AsSpan(2, 2),
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out month))
            {
                return null;
            }
            year = 2000 + shortYear;
        }
        else
        {
            return null;
        }

        return year is >= 1 and <= 9999 && month is >= 1 and <= 12
            ? $"M:{year:D4}{month:D2}"
            : null;
    }

    private static string FormatKpiPercent(double value, int decimals) =>
        value > 0
            ? value.ToString("F" + decimals.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture) + "%"
            : "-";

    private static string FormatKpiCurrency(double value) =>
        value > 0
            ? "$ " + ((long)Math.Round(value)).ToString("N0", CultureInfo.InvariantCulture)
            : "-";

    private static string FormatKpiNullablePercent(double? value, int decimals) =>
        value is >= 0
            ? value.Value.ToString(
                "F" + decimals.ToString(CultureInfo.InvariantCulture),
                CultureInfo.InvariantCulture) + "%"
            : "-";

    private static string FormatKpiNullableCurrency(double? value) =>
        value is >= 0
            ? "$ " + ((long)Math.Round(value.Value)).ToString(
                "N0",
                CultureInfo.InvariantCulture)
            : "-";

    private static string FormatKpiPpm(double value) =>
        value > 0
            ? ((long)Math.Round(value)).ToString("N0", CultureInfo.InvariantCulture) + " ppm"
            : "-";

    private static string FormatKpiNullablePpm(double? value) =>
        value is >= 0
            ? ((long)Math.Round(value.Value)).ToString(
                "N0",
                CultureInfo.InvariantCulture) + " ppm"
            : "-";

    private static string FormatKpiAchievement(double? target, double actual, int decimals) =>
        target is > 0 && actual > 0
            ? (target.Value / actual * 100.0).ToString(
                "F" + decimals.ToString(CultureInfo.InvariantCulture),
                CultureInfo.InvariantCulture) + "%"
            : "-";

    private static List<KpiPeriodColumn> BuildKpiPeriodColumns(
        FCostReport report)
    {
        var result = new List<KpiPeriodColumn>();

        for (int columnIndex = 0; columnIndex < report.Columns.Count; columnIndex++)
        {
            FCostColumnMeta column = report.Columns[columnIndex];
            if (column.Kind == FCostPeriodKind.Month)
            {
                string key = "M:" + KpiPeriodDigits(column);
                bool parsed = TryParseKpiMonth(key, out int year, out int monthNumber);
                result.Add(new KpiPeriodColumn(
                    columnIndex,
                    key,
                    parsed
                        ? monthNumber.ToString(CultureInfo.InvariantCulture) + "월"
                        : column.Header,
                    parsed ? new DateTime(year, monthNumber, 1) : DateTime.MaxValue,
                    IsMonth: true));
                continue;
            }

            if (column.Kind == FCostPeriodKind.Week)
            {
                string key = "W:" + KpiPeriodDigits(column);
                bool parsed = TryParseKpiWeek(key, out int year, out int weekNumber);
                DateTime sortDate = DateTime.MaxValue;
                if (parsed)
                {
                    DateTime firstDate = FirstDateOfKpiWeek(year, weekNumber);
                    // Use the middle of the week for month placement. A week beginning
                    // at the end of June but mostly belonging to July follows 7월.
                    sortDate = firstDate < DateTime.MaxValue.AddDays(-3)
                        ? firstDate.AddDays(3)
                        : firstDate;
                }

                result.Add(new KpiPeriodColumn(
                    columnIndex,
                    key,
                    parsed
                        ? "W" + weekNumber.ToString(CultureInfo.InvariantCulture)
                        : column.Header,
                    sortDate,
                    IsMonth: false));
            }
        }

        bool spansYears = result
            .Where(column => column.SortDate != DateTime.MaxValue)
            .Select(column => column.SortDate.Year)
            .Distinct()
            .Skip(1)
            .Any();

        return result
            .OrderBy(column => column.SortDate)
            .ThenByDescending(column => column.IsMonth)
            .ThenBy(column => column.Key, StringComparer.Ordinal)
            .Select(column =>
            {
                if (!spansYears || column.SortDate == DateTime.MaxValue)
                    return column;

                string prefix = (column.SortDate.Year % 100)
                    .ToString("00", CultureInfo.InvariantCulture);
                return column with { Header = prefix + "." + column.Header };
            })
            .ToList();
    }

    private static string KpiPeriodDigits(FCostColumnMeta column)
    {
        foreach (string candidate in new[] { column.PDate, column.Code })
        {
            string digits = new(candidate.Where(char.IsDigit).ToArray());
            if (digits.Length >= 6)
                return digits[..6];
        }

        return column.Index.ToString(CultureInfo.InvariantCulture);
    }

    private static bool TryParseKpiMonth(string key, out int year, out int month)
    {
        string raw = key.StartsWith("M:", StringComparison.Ordinal) ? key[2..] : key;
        year = 0;
        month = 0;
        return raw.Length >= 6 &&
               int.TryParse(raw.AsSpan(0, 4), NumberStyles.None, CultureInfo.InvariantCulture, out year) &&
               int.TryParse(raw.AsSpan(4, 2), NumberStyles.None, CultureInfo.InvariantCulture, out month) &&
               year is >= 1 and <= 9999 &&
               month is >= 1 and <= 12;
    }

    private static bool TryParseKpiWeek(string key, out int year, out int week)
    {
        string raw = key.StartsWith("W:", StringComparison.Ordinal) ? key[2..] : key;
        year = 0;
        week = 0;
        return raw.Length >= 6 &&
               int.TryParse(raw.AsSpan(0, 4), NumberStyles.None, CultureInfo.InvariantCulture, out year) &&
               int.TryParse(raw.AsSpan(4, 2), NumberStyles.None, CultureInfo.InvariantCulture, out week) &&
               year is >= 1 and <= 9999 &&
               week is >= 1 and <= 54;
    }

    private static DateTime FirstDateOfKpiWeek(int year, int week)
    {
        var calendar = CultureInfo.InvariantCulture.Calendar;
        var date = new DateTime(year, 1, 1);
        var end = new DateTime(year, 12, 31);

        for (; date <= end; date = date.AddDays(1))
        {
            int candidate = calendar.GetWeekOfYear(
                date,
                CalendarWeekRule.FirstDay,
                DayOfWeek.Monday);
            if (candidate == week)
                return date;
        }

        return DateTime.MaxValue;
    }

    /// <summary>
    /// Assemble the tab menu, the inert per-tab &lt;template&gt; bodies, and the tab-switch
    /// script into one self-contained page.
    /// </summary>
    private static string BuildCombinedHtml(IReadOnlyDictionary<string, string> bodies, double ppmDefault)
    {
        var tabButtons = new StringBuilder();
        var templates = new StringBuilder();
        foreach ((string key, string label) in Tabs)
        {
            string encLabel = System.Net.WebUtility.HtmlEncode(label);
            tabButtons.Append(
                $"<button type=\"button\" class=\"rpt-tab\" data-tab=\"{key}\">{encLabel}</button>");
            string body = bodies.TryGetValue(key, out string? b) ? b : PlaceholderBody(key);
            templates.Append($"<template data-tab=\"{key}\">").Append(body).Append("</template>");
        }

        return $$"""
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>BMES Report</title>
<link rel="stylesheet" href="/bootstrap/bootstrap.min.css">
<link rel="stylesheet" href="/app.css?v=2">
<style>
  body { margin: 0; padding: 12px 16px 24px; background: #fff; }
  .rpt-tabs { display: flex; flex-wrap: wrap; gap: 4px; border-bottom: 1px solid #d7dee8; margin-bottom: 14px; }
  .rpt-tab { border: 1px solid transparent; border-bottom: 0; background: transparent; color: #475569;
             padding: 9px 18px; font-size: 13px; font-weight: 600; border-radius: 6px 6px 0 0; cursor: pointer; }
  .rpt-tab:hover { background: #f8fafc; color: #0f172a; }
  .rpt-tab.active { background: #fff; border-color: #d7dee8; color: #0f172a; margin-bottom: -1px; }
  /* Interactive-only controls (copy / move / remove / reason toggles) are inert in a static
     snapshot — hide them. Keep the viewer chrome: top tabs, per-model sub-tabs, the toolbar
     reset, and export toggles. */
  .report-export button:not(.rpt-tab):not(.rpt-subtab):not(.export-toggle):not(#tb-reset) { display: none !important; }
  .report-export .form-check-inline { display: none !important; }

  /* Report-content sizing — mirrors BmesReportPage's own overrides so numbers render
     as compactly as the original interactive view (the .bmes-report-tab-content root
     is also what app.js's bmesReportTableSizer measures column widths against). */
  .bmes-report-tab-content .pivot-wrap,
  .bmes-report-tab-content .weekly-form-wrap { overflow-x: auto !important; text-align: left; }
  .bmes-report-tab-content .pivot-table.text-fit-table {
    table-layout: auto !important; width: max-content !important;
    min-width: 0 !important; max-width: none !important; margin-left: 0; margin-right: auto;
  }
  .bmes-report-tab-content .pivot-table.text-fit-table th:not(.label-th):not(.sep-th),
  .bmes-report-tab-content .pivot-table.text-fit-table td:not(.label-td):not(.group-name-td):not(.sep-td):not(.toggle-cell):not(.row-hide-cell) {
    padding-left: 5px !important; padding-right: 5px !important;
  }
  .bmes-report-tab-content .pivot-table.text-fit-table td.bmes-report-number-cell {
    box-sizing: border-box;
    width: var(--bmes-report-cell-width, auto) !important;
    min-width: var(--bmes-report-cell-width, max-content) !important;
    max-width: var(--bmes-report-cell-width, none) !important;
    white-space: nowrap !important; overflow: visible !important; text-overflow: clip !important;
  }
  /* Shrink the main (black) PPM value to match the delta line. */
  .bmes-report-tab-content .pivot-table td .ppm-cell-value { font-size: 9px; }

  /* KPI: fixed master-data columns + one chronological month/week axis. */
  .kpi-report { padding: 2px 0 24px; }
  .kpi-report-title { margin-bottom: 10px; color: #0f172a; font-size: 16px; font-weight: 700; }
  .kpi-report-note { margin: -4px 0 10px; color: #64748b; font-size: 11px; }
  .kpi-report-wrap { max-width: 100%; overflow-x: auto; }
  .kpi-report-table { width: max-content; min-width: 100%; border-collapse: separate;
                      border-spacing: 0; color: #0f172a; font-size: 11px;
                      --kpi-name-w: 180px; --kpi-type-w: 72px; --kpi-base-w: 86px;
                      --kpi-target-w: 86px; --kpi-ytd-w: 86px; --kpi-line-w: 78px; }
  .kpi-report-table th, .kpi-report-table td {
    box-sizing: border-box; height: 32px; border-right: 1px solid #94a3b8;
    border-bottom: 1px solid #94a3b8; padding: 6px 9px; background: #fff;
    text-align: center; white-space: nowrap; vertical-align: middle;
  }
  .kpi-report-table thead tr > :first-child,
  .kpi-report-table tbody tr.kpi-metric > :first-child { border-left: 1px solid #94a3b8; }
  .kpi-report-table thead th { position: relative; z-index: 3; border-top: 1px solid #94a3b8;
                               background: #dbeafe; font-weight: 700; }
  .kpi-report-table .kpi-pin { position: sticky; z-index: 2; background: #fff; }
  .kpi-report-table thead .kpi-pin { z-index: 5; background: #bfdbfe; }
  .kpi-report-table .kpi-name { left: 0; width: var(--kpi-name-w); min-width: var(--kpi-name-w);
                                max-width: var(--kpi-name-w); font-weight: 700; text-align: left;
                                white-space: normal; }
  .kpi-report-table thead .kpi-name { text-align: center; }
  .kpi-report-table .kpi-type { left: var(--kpi-name-w); width: var(--kpi-type-w);
                                min-width: var(--kpi-type-w); }
  .kpi-report-table .kpi-baseline { left: calc(var(--kpi-name-w) + var(--kpi-type-w));
                                    width: var(--kpi-base-w); min-width: var(--kpi-base-w); }
  .kpi-report-table .kpi-target { left: calc(var(--kpi-name-w) + var(--kpi-type-w) + var(--kpi-base-w));
                                  width: var(--kpi-target-w); min-width: var(--kpi-target-w); }
  .kpi-report-table .kpi-ytd { left: calc(var(--kpi-name-w) + var(--kpi-type-w) + var(--kpi-base-w) + var(--kpi-target-w));
                               width: var(--kpi-ytd-w); min-width: var(--kpi-ytd-w); }
  .kpi-report-table .kpi-line { left: calc(var(--kpi-name-w) + var(--kpi-type-w) + var(--kpi-base-w) + var(--kpi-target-w) + var(--kpi-ytd-w));
                                width: var(--kpi-line-w); min-width: var(--kpi-line-w); font-weight: 700; }
  .kpi-report-table .kpi-period { min-width: 82px; }
  .kpi-report-table thead .kpi-month { background: #bfdbfe; }
  .kpi-report-table thead .kpi-week { background: #e2e8f0; }
  .kpi-month-head { display: inline-flex; align-items: center; justify-content: center; gap: 4px; }
  .kpi-month-toggle { display: inline-flex; width: 16px; height: 16px; align-items: center;
                      justify-content: center; padding: 0; border: 1px solid #64748b;
                      border-radius: 3px; background: #eff6ff; color: #1e3a8a;
                      font: 700 11px/14px system-ui, -apple-system, "Segoe UI", sans-serif;
                      cursor: pointer; }
  .kpi-month-toggle:hover { background: #dbeafe; }
  .kpi-month-toggle:focus-visible { outline: 2px solid #2563eb; outline-offset: 1px; }
  .kpi-report-table .kpi-value { min-width: 82px; text-align: right;
                                 font-variant-numeric: tabular-nums; }
  .kpi-report-table .kpi-metric-end > *,
  .kpi-report-table .kpi-group-cell { border-bottom: 2px solid #64748b; }
  .kpi-report-table tbody .kpi-total > *,
  .kpi-report-table tbody .kpi-total + .kpi-achievement > * { background: #f8fafc; font-weight: 700; }
  .kpi-report-table .kpi-pending { color: #94a3b8; }
  .kpi-empty { padding: 32px; color: #64748b; font-size: 13px; text-align: center; }

  /* ── Viewer toolbar ──────────────────────────────────────────────────────────
     Live display settings for the static snapshot. Everything here is applied by
     JS to #rpt-host only, so the toolbar itself never rescales. */
  .rpt-toolbar { position: sticky; top: 0; z-index: 30; display: flex; flex-wrap: wrap;
                 align-items: center; gap: 6px 14px; padding: 8px 10px; margin-bottom: 10px;
                 background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px;
                 font: 12px/1.2 system-ui, -apple-system, "Segoe UI", sans-serif; color: #334155; }
  .rpt-toolbar label { display: inline-flex; align-items: center; gap: 5px; font-weight: 600; white-space: nowrap; }
  .rpt-toolbar input, .rpt-toolbar select { font: inherit; font-weight: 400; padding: 2px 5px;
                 border: 1px solid #cbd5e1; border-radius: 4px; background: #fff; color: #0f172a; }
  .rpt-toolbar input[type=number] { width: 62px; }
  .rpt-toolbar select { max-width: 150px; }
  .rpt-toolbar .tb-sep { width: 1px; align-self: stretch; background: #dbe3ec; }
  .rpt-toolbar button { font: inherit; font-weight: 600; padding: 3px 10px; cursor: pointer;
                 border: 1px solid #cbd5e1; border-radius: 4px; background: #fff; color: #334155; }
  .rpt-toolbar button:hover { background: #eef2f7; }
  /* The toolbar is chrome, not report content — never print it. */
  @media print { .rpt-toolbar, .rpt-tabs, .kpi-month-toggle { display: none !important; } }

  /* Sub-tabs inside a tab (Daily: Summary + one per model). Same memory trick as the
     top-level tabs — only the selected section is mounted. */
  .rpt-subtabs { display: flex; flex-wrap: wrap; gap: 4px; margin: 0 0 10px;
                 padding-bottom: 6px; border-bottom: 1px solid #e2e8f0; }
  .rpt-subtab { border: 1px solid #dbe3ec; background: #fff; color: #475569;
                padding: 4px 12px; font: 600 12px/1.2 system-ui, -apple-system, "Segoe UI", sans-serif;
                border-radius: 999px; cursor: pointer; }
  .rpt-subtab:hover { background: #f1f5f9; color: #0f172a; }
  .rpt-subtab.active { background: #0f172a; border-color: #0f172a; color: #fff; }

  /* Typography knobs. zoom scales the hard-coded px sizes in the report markup
     (9px/10px cells etc.) proportionally, which a font-size override cannot do
     without flattening those deliberate differences. */
  #rpt-host { zoom: var(--tb-zoom, 1);
              font-family: var(--tb-font, inherit);
              font-weight: var(--tb-weight, inherit); }
  #rpt-host th, #rpt-host td, #rpt-host .card-header, #rpt-host .card-body {
              font-family: inherit; font-weight: inherit; }
</style>
</head>
<body class="report-export">
<div class="rpt-tabs">{{tabButtons}}</div>
<div class="rpt-toolbar">
  <label>폰트
    <select id="tb-font">
      <option value="">기본</option>
      <option value="'Malgun Gothic','맑은 고딕',sans-serif" selected>맑은 고딕</option>
      <option value="'Noto Sans KR',sans-serif">Noto Sans KR</option>
      <option value="'Segoe UI',system-ui,sans-serif">Segoe UI</option>
      <option value="'Nanum Gothic',sans-serif">나눔고딕</option>
      <option value="Consolas,'D2Coding',monospace">고정폭</option>
    </select>
  </label>
  <label>크기 <input type="number" id="tb-zoom" min="50" max="250" step="5" value="90">%</label>
  <label>굵기
    <select id="tb-weight">
      <option value="">기본</option>
      <option value="300">가늘게</option>
      <option value="400" selected>보통</option>
      <option value="600">약간 굵게</option>
      <option value="700">굵게</option>
    </select>
  </label>
  <span class="tb-sep"></span>
  <label>Date <input type="number" id="tb-date" min="0" max="366" step="1" placeholder="전체" value="7"> 일</label>
  <label>Week <input type="number" id="tb-week" min="0" max="104" step="1" placeholder="전체" value="4"> 주</label>
  <label>Month <input type="number" id="tb-month" min="0" max="36" step="1" placeholder="전체" value="3"> 개월</label>
  <span class="tb-sep"></span>
  <label>Min PPM <input type="number" id="tb-ppm" min="0" max="1000000" step="100" value="{{ppmDefault}}"></label>
  <button type="button" id="tb-reset">초기화</button>
</div>
<div id="rpt-host" class="bmes-report-tab-content"></div>
{{templates}}
<script src="/js/chart.umd.min.js?v=1"></script>
<script src="/js/app.js?v=79"></script>
<script>
(function () {
  var host = document.getElementById('rpt-host');
  var tabs = Array.prototype.slice.call(document.querySelectorAll('.rpt-tab'));

  function initCharts(scope) {
    scope.querySelectorAll('script[type="application/json"][data-chart]').forEach(function (s) {
      try {
        var cfg = JSON.parse(s.textContent || '{}');
        if (cfg.canvasId && window.ngRateGroupChart && window.ngRateGroupChart.render) {
          window.ngRateGroupChart.render(cfg.canvasId, cfg.labels, cfg.series);
        }
      } catch (e) { /* ignore malformed chart payloads */ }
    });
  }

  // ── Viewer settings ────────────────────────────────────────────────────────
  // A static snapshot carries every column and every reason row; the toolbar only
  // decides what stays visible, so all of this is pure show/hide plus a re-total.
  var STORE = 'bmesReportViewerSettings';
  var el = {
    font:   document.getElementById('tb-font'),
    zoom:   document.getElementById('tb-zoom'),
    weight: document.getElementById('tb-weight'),
    date:   document.getElementById('tb-date'),
    week:   document.getElementById('tb-week'),
    month:  document.getElementById('tb-month'),
    ppm:    document.getElementById('tb-ppm')
  };
  // Mirrors the markup's default attributes so 초기화 restores the same starting view.
  var DEFAULTS = {
    font: "'Malgun Gothic','맑은 고딕',sans-serif", zoom: '90', weight: '400',
    date: '7', week: '4', month: '3', ppm: '{{ppmDefault}}'
  };

  function num(input, fallback) {
    var v = parseFloat(input.value);
    return isNaN(v) ? fallback : v;
  }

  function applyTypography() {
    host.style.setProperty('--tb-font',   el.font.value   || 'inherit');
    host.style.setProperty('--tb-weight', el.weight.value || 'inherit');
    host.style.setProperty('--tb-zoom',   String(num(el.zoom, 100) / 100));
  }

  function isSep(c) {
    return c.classList.contains('sep-td') || c.classList.contains('sep-th');
  }

  /// Period blocks always sit at the right edge of every row in the order
  /// date → week → month, so walking cells right-to-left (skipping the thin
  /// separator cells) identifies them without tagging all ~75 emit sites.
  /// The per-tab <script data-periods> payload supplies the block sizes.
  function applyPeriods(scope) {
    var meta = scope.querySelector('script[data-periods]');
    if (!meta) return;
    var m;
    try { m = JSON.parse(meta.textContent || '{}'); } catch (e) { return; }

    var lim = {
      date:  num(el.date,  Infinity),
      week:  num(el.week,  Infinity),
      month: num(el.month, Infinity)
    };
    // Right-to-left: month is the rightmost block.
    var blocks = [['month', m.month | 0], ['week', m.week | 0], ['date', m.date | 0]]
                   .filter(function (b) { return b[1] > 0; });

    scope.querySelectorAll('table').forEach(function (t) {
      Array.prototype.forEach.call(t.rows, function (r) {
        // The block-header row carries colspans instead of one cell per period.
        var blk = r.querySelector('.blk-date, .blk-week, .blk-month');
        if (blk) {
          [['date', '.blk-date'], ['week', '.blk-week'], ['month', '.blk-month']].forEach(function (p) {
            var th = r.querySelector(p[1]);
            if (!th) return;
            var keep = Math.max(0, Math.min(lim[p[0]], m[p[0]] | 0));
            th.colSpan = Math.max(1, keep);
            th.style.display = keep > 0 ? '' : 'none';
          });
          return;
        }
        var cells = Array.prototype.filter.call(r.cells, function (c) { return !isSep(c); });
        var idx = cells.length;
        blocks.forEach(function (b) {
          var total = b[1], keep = Math.max(0, Math.min(lim[b[0]], total)), start = idx - total;
          if (start < 0) { idx = 0; return; }
          for (var k = 0; k < total; k++) cells[start + k].style.display = (k < keep) ? '' : 'none';
          idx = start;
        });
      });
    });
  }

  function parseNum(text) {
    var t = (text || '').replace(/[,\s]/g, '');
    var v = parseFloat(t);
    return isNaN(v) ? 0 : v;
  }

  /// Reason detail rows are exported unfiltered and tagged with their reference-date
  /// PPM, so the threshold is a viewer setting here. Section totals are summed from
  /// the details that survive the filter, matching the interactive page's behaviour.
  function applyPpm(scope) {
    var threshold = num(el.ppm, 0);
    scope.querySelectorAll('[data-ppm-label]').forEach(function (n) {
      n.textContent = n.getAttribute('data-ppm-label')
        .replace('{v}', Math.round(threshold).toLocaleString('en-US'));
    });

    // Only the trailing period cells hold summable numbers; everything to their left
    // is a label. The period meta gives that count exactly, so no cell sniffing.
    var periodCount = 0;
    var meta = scope.querySelector('script[data-periods]');
    if (meta) {
      try {
        var m = JSON.parse(meta.textContent || '{}');
        periodCount = (m.date | 0) + (m.week | 0) + (m.month | 0);
      } catch (e) { periodCount = 0; }
    }

    scope.querySelectorAll('table').forEach(function (t) {
      var rows = Array.prototype.slice.call(t.rows);
      if (!t.querySelector('tr[data-ppm]')) return;

      var section = null;                       // current total-row and its details
      function settle() {
        if (!section) return;
        var shown = section.details.filter(function (d) { return d.style.display !== 'none'; });
        section.total.style.display = shown.length ? '' : 'none';
        if (!shown.length || !periodCount) { section = null; return; }
        // Sum each period column across visible details, aligning from the right.
        var tc = Array.prototype.filter.call(section.total.cells, function (c) { return !isSep(c); });
        for (var back = 1; back <= Math.min(periodCount, tc.length); back++) {
          var cell = tc[tc.length - back];
          var sum = 0, seen = false;
          shown.forEach(function (d) {
            var dc = Array.prototype.filter.call(d.cells, function (c) { return !isSep(c); });
            var src = dc[dc.length - back];
            if (!src) return;
            seen = true;
            sum += parseNum(src.textContent);
          });
          if (seen) cell.textContent = sum > 0 ? Math.round(sum).toLocaleString('en-US') : '-';
        }
        section = null;
      }

      rows.forEach(function (r) {
        if (r.classList.contains('total-row')) { settle(); section = { total: r, details: [] }; return; }
        var p = r.getAttribute('data-ppm');
        if (p === null) return;
        r.style.display = (parseFloat(p) >= threshold) ? '' : 'none';
        if (section) section.details.push(r);
      });
      settle();
    });
  }

  var kpiMonthExpanded = Object.create(null);

  function setKpiMonthExpanded(scope, monthKey, expanded) {
    scope.querySelectorAll('[data-kpi-week-month]').forEach(function (cell) {
      if (cell.getAttribute('data-kpi-week-month') === monthKey)
        cell.hidden = !expanded;
    });

    scope.querySelectorAll('[data-kpi-month-toggle]').forEach(function (button) {
      if (button.getAttribute('data-kpi-month-toggle') !== monthKey) return;
      var label = button.getAttribute('data-kpi-month-label') || '';
      button.textContent = expanded ? '−' : '+';
      button.setAttribute('aria-expanded', expanded ? 'true' : 'false');
      button.setAttribute('aria-label', label + (expanded ? ' 주차 접기' : ' 주차 펼치기'));
      button.setAttribute('title', label + (expanded ? ' 주차 접기' : ' 주차 펼치기'));
    });
  }

  function applyKpiMonthState(scope) {
    scope.querySelectorAll('[data-kpi-month-toggle]').forEach(function (button) {
      var monthKey = button.getAttribute('data-kpi-month-toggle');
      if (!monthKey) return;
      setKpiMonthExpanded(scope, monthKey, kpiMonthExpanded[monthKey] !== false);
    });
  }

  function applyAll() {
    applyTypography();
    applyPeriods(host);
    applyPpm(host);
    applyKpiMonthState(host);
  }

  function save() {
    try {
      localStorage.setItem(STORE, JSON.stringify({
        font: el.font.value, zoom: el.zoom.value, weight: el.weight.value,
        date: el.date.value, week: el.week.value, month: el.month.value, ppm: el.ppm.value
      }));
    } catch (e) { /* private mode / quota — settings just do not persist */ }
  }

  function load() {
    try {
      var s = JSON.parse(localStorage.getItem(STORE) || '{}');
      Object.keys(el).forEach(function (k) { if (s[k] !== undefined) el[k].value = s[k]; });
    } catch (e) { /* ignore malformed stored settings */ }
  }

  Object.keys(el).forEach(function (k) {
    el[k].addEventListener('input', function () { applyAll(); save(); });
    el[k].addEventListener('change', function () { applyAll(); save(); });
  });

  document.getElementById('tb-reset').addEventListener('click', function () {
    Object.keys(el).forEach(function (k) { el[k].value = DEFAULTS[k]; });
    applyAll(); save();
  });

  host.addEventListener('click', function (event) {
    var button = event.target.closest
      ? event.target.closest('[data-kpi-month-toggle]')
      : null;
    if (!button || !host.contains(button)) return;

    var monthKey = button.getAttribute('data-kpi-month-toggle');
    if (!monthKey) return;
    var expanded = button.getAttribute('aria-expanded') !== 'false';
    kpiMonthExpanded[monthKey] = !expanded;
    setKpiMonthExpanded(host, monthKey, !expanded);
    sizeTables();
  });

  function sizeTables() {
    if (window.bmesReportTableSizer && window.bmesReportTableSizer.start) {
      try { window.bmesReportTableSizer.start(); } catch (e) { }
    }
  }

  /// A tab whose body carries [data-daily-section] blocks (Daily: the Summary card plus
  /// one block per model) is split into sub-tabs. Each block is detached into its own
  /// template so only the selected one is ever laid out — the same reason the top-level
  /// tabs exist. Tabs without those markers render unchanged.
  function initSubTabs(scope) {
    var secs = Array.prototype.slice.call(scope.querySelectorAll('[data-daily-section]'));
    if (secs.length < 2) return;

    var bar = document.createElement('div');
    bar.className = 'rpt-subtabs';
    var subHost = document.createElement('div');

    var items = secs.map(function (s) {
      var tpl = document.createElement('template');
      tpl.content.appendChild(s);              // detaches s from the live DOM
      return { label: s.getAttribute('data-daily-label') || '(이름 없음)', tpl: tpl };
    });

    var buttons = items.map(function (item, i) {
      var b = document.createElement('button');
      b.type = 'button';
      b.className = 'rpt-subtab';
      b.textContent = item.label;
      b.addEventListener('click', function () { showSub(i); });
      bar.appendChild(b);
      return b;
    });

    function showSub(i) {
      buttons.forEach(function (b, j) { b.classList.toggle('active', i === j); });
      subHost.innerHTML = '';                  // drop previous section → free its DOM
      subHost.appendChild(items[i].tpl.content.cloneNode(true));
      applyPeriods(host);
      applyPpm(host);
      initCharts(subHost);
      sizeTables();
    }

    scope.appendChild(bar);
    scope.appendChild(subHost);
    showSub(0);
  }

  function show(key) {
    tabs.forEach(function (t) { t.classList.toggle('active', t.dataset.tab === key); });
    host.innerHTML = '';                       // drop previous tab body → free its DOM
    var tpl = document.querySelector('template[data-tab="' + key + '"]');
    if (tpl) host.appendChild(tpl.content.cloneNode(true));
    initSubTabs(host);
    initCharts(host);
    applyAll();                                // re-apply viewer settings to the new body
    sizeTables();
  }

  tabs.forEach(function (t) { t.addEventListener('click', function () { show(t.dataset.tab); }); });
  load();
  if (tabs.length) show(tabs[0].dataset.tab);
})();
</script>
</body>
</html>
""";
    }

    /// <summary>Remove previously generated token folders, keeping the run that just
    /// finished (best-effort).</summary>
    private static void CleanupOldTokens(string keepToken)
    {
        try
        {
            if (!Directory.Exists(ExportRoot)) return;
            foreach (string sub in Directory.GetDirectories(ExportRoot))
            {
                if (string.Equals(Path.GetFileName(sub), keepToken, StringComparison.OrdinalIgnoreCase))
                    continue;
                try { Directory.Delete(sub, recursive: true); }
                catch { /* a folder may be momentarily locked; ignore */ }
            }
        }
        catch { /* best-effort cleanup */ }
    }
}

/// <summary>
/// Files belonging to one TTL-valid, metadata-published report token. ReportJsonPath is
/// null when JSON is absent or belongs to an unsupported cache/schema generation, so the
/// host can safely fall back to the legacy HTML without exposing the JSON as a static file.
/// </summary>
public sealed record BmesReportArtifacts(
    string Token,
    string LegacyHtmlPath,
    string? ReportJsonPath,
    string MetadataPath,
    string CacheKey,
    DateTimeOffset CreatedAtUtc,
    bool IsCurrentContract);
