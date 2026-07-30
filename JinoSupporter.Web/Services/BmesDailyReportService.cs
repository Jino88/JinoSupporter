using System.Text.Json;

namespace JinoSupporter.Web.Services;

/// <summary>A single defect driver behind one model's most recent week, with how it
/// moved against the week before.</summary>
/// <param name="ProcessType">BMES process class — MAIN / SUB / FUNCTION / VISUAL.</param>
/// <param name="IsNew">True when the cause had no defects at all in the previous week.</param>
public sealed record BmesDailyReportCause(
    string NgName,
    string ProcessName,
    string ProcessType,
    double Ppm,
    double PreviousPpm,
    double DeltaPpm,
    double DeltaPercent,
    bool   IsNew);

/// <summary>One analysis unit's NG-rate trend over the report window, plus what drove
/// its most recent week.
///
/// The unit is the sub group — two levels below a model group in the
/// Group → Model → Sub Group → Shift Group tree.</summary>
/// <param name="ModelName">The sub group's own name, e.g. <c>BRS-161016S08ZZ_C2</c>.</param>
/// <param name="GroupName">Top-level model group, e.g. <c>SPK</c>.</param>
/// <param name="ParentModelName">The model (MAKTX) this sub group hangs off.</param>
/// <param name="PreviousPpm">The baseline <c>RecentPpm</c> is measured against: the most
/// recent earlier period this unit was actually produced in — not necessarily the period
/// immediately before. <see cref="BaselineStepsBack"/> says how far back it sits.</param>
/// <param name="TrendPpm">PPM per period column, aligned to
/// <see cref="BmesDailyReportSnapshot.WeekHeaders"/> (week rows) or
/// <see cref="BmesDailyReportSnapshot.DayHeaders"/> (day rows) — newest first, matching
/// the dashboard's left-to-right reading order.</param>
/// <param name="BaselineHeader">Header of the period <see cref="PreviousPpm"/> came from,
/// empty when the unit has no earlier production in the window.</param>
/// <param name="BaselineStepsBack">1 = the period immediately before the latest one,
/// 2 = one further back, and so on. 0 means no baseline was found.</param>
public sealed record BmesDailyReportModelRow(
    string                     ModelName,
    string                     GroupName,
    string                     ParentModelName,
    double                     RecentPpm,
    double                     PreviousPpm,
    double                     DeltaPpm,
    double                     DeltaPercent,
    List<double>               TrendPpm,
    List<BmesDailyReportCause> TopCauses,
    long                       RecentNgQty,
    string                     BaselineHeader,
    int                        BaselineStepsBack);

/// <summary>Cached dashboard payload. Serialized as JSON into the app settings table.</summary>
public sealed record BmesDailyReportSnapshot
{
    /// <summary>Shape marker. A snapshot written by an older layout deserializes with
    /// missing/null members, so <see cref="BmesDailyReportService.Load"/> discards
    /// anything that does not match the current version. Bump this whenever the
    /// records in this file change shape.</summary>
    public int SchemaVersion { get; init; }

    public DateTime GeneratedAt { get; init; }
    public DateTime StartDate   { get; init; }
    public DateTime EndDate     { get; init; }

    /// <summary>Week column headers, newest first — the dashboard reads left to right
    /// with the most recent week on the left.</summary>
    public List<string> WeekHeaders { get; init; } = [];

    /// <summary>Headers of the two weeks the overall figures compare: the latest week and
    /// the most recent earlier week that carries data.</summary>
    public string RecentWeekHeader   { get; init; } = string.Empty;
    public string PreviousWeekHeader { get; init; } = string.Empty;

    /// <summary>Overall PPM (all selected models) for those two weeks.</summary>
    public double OverallRecentPpm   { get; init; }
    public double OverallPreviousPpm { get; init; }

    /// <summary>Day column headers, newest first — the day-over-day half of the dashboard,
    /// capped at <see cref="BmesDailyReportService.DailyTrendDays"/> days.</summary>
    public List<string> DayHeaders { get; init; } = [];

    /// <summary>The same two headers as <see cref="RecentWeekHeader"/> /
    /// <see cref="PreviousWeekHeader"/>, one granularity down.</summary>
    public string RecentDayHeader   { get; init; } = string.Empty;
    public string PreviousDayHeader { get; init; } = string.Empty;

    public double OverallRecentDayPpm   { get; init; }
    public double OverallPreviousDayPpm { get; init; }

    /// <summary>Number of models that had usable data in the window.</summary>
    public int ModelCount { get; init; }

    /// <summary>Models whose recent week is worse than their own baseline, worst first.</summary>
    public List<BmesDailyReportModelRow> Worsened { get; init; } = [];

    /// <summary>Models that improved most, best first.</summary>
    public List<BmesDailyReportModelRow> Improved { get; init; } = [];

    /// <summary>The same two rankings computed day over day instead of week over week.</summary>
    public List<BmesDailyReportModelRow> WorsenedDaily { get; init; } = [];
    public List<BmesDailyReportModelRow> ImprovedDaily { get; init; } = [];

    /// <summary>Non-null when the refresh failed; the dashboard renders this instead of numbers.</summary>
    public string? Error { get; init; }
}

/// <summary>
/// Backs the DAILY REPORT dashboard: the same BMES NG-rate data the Report page
/// pulls, fixed to [2 months ago → yesterday] and reduced to a per-model
/// deterioration ranking with each model's top defect drivers.
///
/// Reads never block. Callers get whatever snapshot is stored; when it is stale a
/// single shared background refresh starts and <see cref="SnapshotChanged"/> fires
/// once new numbers land. The heavy lifting is already cached one layer down —
/// <see cref="NgRateService.GetOrFetchAsync"/> only hits BMES for the last few
/// days and reads the monthly DB cache for the rest of the window.
/// </summary>
public sealed class BmesDailyReportService
{
    /// <summary>Bump whenever the snapshot changes shape *or* meaning — see
    /// <see cref="BmesDailyReportSnapshot.SchemaVersion"/>. 1 = per-model-group,
    /// 2 = per-model rows with cause breakdown, 3 = week ordering corrected (2 read the
    /// period columns backwards and reported the oldest week as "recent"),
    /// 4 = baseline is W-1 instead of a multi-week average, causes carry week-over-week
    /// movement, 5 = models not produced in the latest week are excluded,
    /// 6 = rows are sub groups (level 2) instead of models (level 1),
    /// 7 = causes carry their process type (MAIN / SUB / …),
    /// 8 = the baseline skips periods the unit was not produced in (and rows say how far
    /// back it landed), plus a day-over-day ranking beside the weekly one,
    /// 9 = periods with no production at all (Sundays, holidays) no longer anchor a
    /// ranking — 8 anchored the daily half on an empty Sunday and reported nothing.</summary>
    private const int    SnapshotSchemaVersion = 9;

    private const string SnapshotSettingKey  = "BmesDailyReport:Snapshot";
    private const string ReportSelectionKey  = "BmesReport:SelectionState";
    private const int    WorsenedTopN        = 10;
    private const int    ImprovedTopN        = 5;
    private const int    TopCauseCount       = 5;

    /// <summary>How far the day-over-day half looks back: the length of its sparkline and
    /// the limit on the walk that skips days a model was not produced in. A model with no
    /// production in the last month has nothing useful to compare against anyway.</summary>
    internal const int   DailyTrendDays      = 30;

    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    private readonly IServiceScopeFactory _scopeFactory;
    private readonly WebRepository        _repo;
    private readonly AppActivityLogger    _activity;

    private readonly object _gate = new();
    private Task?           _refreshTask;
    private List<string>    _logs = [];

    public BmesDailyReportService(
        IServiceScopeFactory scopeFactory,
        WebRepository        repo,
        AppActivityLogger    activity)
    {
        _scopeFactory = scopeFactory;
        _repo         = repo;
        _activity     = activity;
    }

    /// <summary>Raised on the refreshing thread once a new snapshot is stored.</summary>
    public event Action? SnapshotChanged;

    public bool IsRefreshing
    {
        get { lock (_gate) return _refreshTask is not null; }
    }

    public IReadOnlyList<string> Logs
    {
        get { lock (_gate) return _logs.ToList(); }
    }

    /// <summary>Two months back through yesterday — BMES data for today is still moving.</summary>
    public static DateTime WindowStart => DateTime.Today.AddMonths(-2);
    public static DateTime WindowEnd   => DateTime.Today.AddDays(-1);

    public BmesDailyReportSnapshot? Load()
    {
        string? json = _repo.GetSetting(SnapshotSettingKey);
        if (string.IsNullOrWhiteSpace(json)) return null;

        BmesDailyReportSnapshot? snapshot;
        try
        {
            snapshot = JsonSerializer.Deserialize<BmesDailyReportSnapshot>(json, JsonOptions);
        }
        catch
        {
            return null;
        }

        // A snapshot from an older layout would deserialize with null collections and
        // NRE at render time. Treat it as no cache at all — a refresh will replace it.
        if (snapshot is null || snapshot.SchemaVersion != SnapshotSchemaVersion) return null;

        return snapshot with
        {
            WeekHeaders   = snapshot.WeekHeaders ?? [],
            DayHeaders    = snapshot.DayHeaders ?? [],
            Worsened      = Normalize(snapshot.Worsened),
            Improved      = Normalize(snapshot.Improved),
            WorsenedDaily = Normalize(snapshot.WorsenedDaily),
            ImprovedDaily = Normalize(snapshot.ImprovedDaily),
        };
    }

    /// <summary>Belt-and-braces against a hand-edited or partially written payload:
    /// the render path treats these collections as never-null.</summary>
    private static List<BmesDailyReportModelRow> Normalize(List<BmesDailyReportModelRow>? rows) =>
        (rows ?? [])
        .Select(r => r with
        {
            TrendPpm  = r.TrendPpm ?? [],
            TopCauses = r.TopCauses ?? [],
        })
        .ToList();

    /// <summary>A snapshot goes stale when it no longer runs through yesterday, or
    /// when it was generated on an earlier day (a failed run still counts as
    /// generated so a broken BMES connection can't spin the refresh all day).</summary>
    public static bool IsStale(BmesDailyReportSnapshot? snapshot) =>
        snapshot is null ||
        snapshot.EndDate.Date     != WindowEnd.Date ||
        snapshot.GeneratedAt.Date != DateTime.Today;

    /// <summary>Start a refresh if one is warranted and none is already running.
    /// Returns immediately — watch <see cref="SnapshotChanged"/> for the result.</summary>
    public void EnsureFresh(bool force = false)
    {
        if (!force && !IsStale(Load())) return;

        lock (_gate)
        {
            if (_refreshTask is not null) return;
            _logs = [];
            _refreshTask = Task.Run(RefreshAsync);
        }
    }

    private async Task RefreshAsync()
    {
        DateTime start = WindowStart;
        DateTime end   = WindowEnd;

        try
        {
            _activity.Log("BmesDailyReport", $"Refresh start {start:yyyy-MM-dd}~{end:yyyy-MM-dd}");
            BmesDailyReportSnapshot snapshot = await BuildAsync(start, end);
            Store(snapshot);
            _activity.Log("BmesDailyReport",
                $"Refresh done groups={snapshot.ModelCount} worsened={snapshot.Worsened.Count}");
        }
        catch (Exception ex)
        {
            AddLog("[ERROR] " + ex.Message);
            _activity.Log("BmesDailyReport", "Refresh failed: " + ex);
            Store(new BmesDailyReportSnapshot
            {
                GeneratedAt = DateTime.Now,
                StartDate   = start,
                EndDate     = end,
                Error       = ex.Message,
            });
        }
        finally
        {
            lock (_gate) _refreshTask = null;
            SnapshotChanged?.Invoke();
        }
    }

    private void Store(BmesDailyReportSnapshot snapshot) =>
        _repo.SetSetting(
            SnapshotSettingKey,
            JsonSerializer.Serialize(snapshot with { SchemaVersion = SnapshotSchemaVersion }, JsonOptions));

    private async Task<BmesDailyReportSnapshot> BuildAsync(DateTime start, DateTime end)
    {
        using IServiceScope scope = _scopeFactory.CreateScope();
        IServiceProvider sp = scope.ServiceProvider;

        var settings  = sp.GetRequiredService<NgRateSettingsService>();
        var ngRate    = sp.GetRequiredService<NgRateService>();
        var reportSvc = sp.GetRequiredService<NgRateReportService>();

        if (!settings.IsCredentialsConfigured)
            throw new InvalidOperationException("BMES credentials are not configured. Add an account under Setting.");

        List<ModelGroupRecord> groups = LoadSelectedGroups();
        if (groups.Count == 0)
            throw new InvalidOperationException("No model groups are selected. Configure groups under Model Group first.");

        var progress = new Progress<string>(AddLog);

        string? dbPath = await ngRate.GetOrFetchAsync(start, end, progress);
        if (string.IsNullOrEmpty(dbPath))
            throw new InvalidOperationException("Could not fetch data from BMES.");

        HierReports hier = await HierReportBuilder.BuildAsync(
            reportSvc, dbPath, groups, progress, start, end);

        if (hier.ByGroup is null)
            throw new InvalidOperationException("The report contains no aggregated group data.");

        return Summarize(hier, start, end);
    }

    private static BmesDailyReportSnapshot Summarize(HierReports hier, DateTime start, DateTime end)
    {
        NgRateReportService.NgRateReport report = hier.ByGroup!;
        HierPpmLookup lookup = HierPpmLookup.From(hier);

        NgRateReportService.SummaryPivotRow? total = report.Summary.FirstOrDefault(r => r.IsTotal);

        // The report emits period columns newest-first (ExtractCols sorts descending),
        // which is exactly the dashboard's reading order: leftmost = most recent period.
        // Index 0 is therefore the latest week (resp. day) — once the empty columns in
        // front of it are dropped, see TrimToLatestWithData.
        List<NgRateReportService.PeriodColumn> weekCols = TrimToLatestWithData(total, report.WeekCols);
        List<NgRateReportService.PeriodColumn> dayCols  =
            TrimToLatestWithData(total, report.DateCols.Take(DailyTrendDays).ToList());

        // The analysis unit is the sub group — two levels under a model group in the
        // Group → Model → Sub Group → Shift Group tree.
        //
        // The key must be the one HierReportBuilder registered in its Sub1 mapping,
        // so build it with the same helper; a sub with no name is the model's default
        // bucket and is labelled with the model name instead.
        List<SubGroupUnit> units = hier.Groups
            .SelectMany(g => g.MidGroups
                .Where(m => !string.IsNullOrWhiteSpace(m.Material))
                .SelectMany(m => m.SubGroups
                    .Where(s => s.AllLineShifts.Any())
                    .Select(s => new SubGroupUnit(
                        Key:    ModelGroupPickerHelpers.SubGroupKeyOf(g.Name, m.Material, s),
                        Name:   string.IsNullOrWhiteSpace(s.Name) ? m.Material : s.Name,
                        Group:  g.Name,
                        Parent: m.Material))))
            .DistinctBy(u => u.Key)
            .ToList();

        if (weekCols.Count == 0 || units.Count == 0)
        {
            return new BmesDailyReportSnapshot
            {
                GeneratedAt = DateTime.Now,
                StartDate   = start,
                EndDate     = end,
                ModelCount  = units.Count,
                Error       = "No defect data was aggregated for this period.",
            };
        }

        // One pass over the report's raw map, so every "was this unit produced in that
        // period" question below — and there is one per unit per candidate baseline — is a
        // dictionary hit rather than another scan of the whole map.
        ProductionIndex production = ProductionIndex.Build(report);

        List<BmesDailyReportModelRow> weekRows = BuildRows(lookup, production, units, weekCols);
        List<BmesDailyReportModelRow> dayRows  = BuildRows(lookup, production, units, dayCols);

        (double weekPreviousPpm, string weekPreviousHeader) = OverallBaseline(total, weekCols);
        (double dayPreviousPpm,  string dayPreviousHeader)  = OverallBaseline(total, dayCols);

        return new BmesDailyReportSnapshot
        {
            GeneratedAt        = DateTime.Now,
            StartDate          = start,
            EndDate            = end,
            WeekHeaders        = weekCols.Select(c => c.Header).ToList(),
            RecentWeekHeader   = weekCols[0].Header,
            PreviousWeekHeader = weekPreviousHeader,
            OverallRecentPpm   = total?.Ppm.GetValueOrDefault(weekCols[0].Key) ?? 0,
            OverallPreviousPpm = weekPreviousPpm,
            DayHeaders         = dayCols.Select(c => c.Header).ToList(),
            RecentDayHeader    = dayCols.Count > 0 ? dayCols[0].Header : string.Empty,
            PreviousDayHeader  = dayPreviousHeader,
            OverallRecentDayPpm   = dayCols.Count > 0 ? total?.Ppm.GetValueOrDefault(dayCols[0].Key) ?? 0 : 0,
            OverallPreviousDayPpm = dayPreviousPpm,
            ModelCount         = weekRows.Count,
            Worsened      = Rank(weekRows, worse: true),
            Improved      = Rank(weekRows, worse: false),
            WorsenedDaily = Rank(dayRows,  worse: true),
            ImprovedDaily = Rank(dayRows,  worse: false),
        };
    }

    /// <summary>Rank one granularity's rows into the two lists the dashboard shows.</summary>
    private static List<BmesDailyReportModelRow> Rank(List<BmesDailyReportModelRow> rows, bool worse) =>
        worse
            ? rows.Where(r => r.DeltaPpm > 0).OrderByDescending(r => r.DeltaPpm).Take(WorsenedTopN).ToList()
            : rows.Where(r => r.DeltaPpm < 0).OrderBy(r => r.DeltaPpm).Take(ImprovedTopN).ToList();

    /// <summary>Build one ranking pass over a single granularity — weeks or days; the two
    /// differ only in which period columns come in.
    ///
    /// The baseline is not simply the column before the latest one. A model that was not
    /// produced then reads as 0 PPM, and comparing against that 0 reports a jump from
    /// nothing as if the model had got worse. So the search walks back until it finds a
    /// period the model actually ran in, and the row records how far it had to go for the
    /// dashboard to label ("2 weeks back", "3 days back").</summary>
    private static List<BmesDailyReportModelRow> BuildRows(
        HierPpmLookup                          lookup,
        ProductionIndex                        production,
        List<SubGroupUnit>                     units,
        List<NgRateReportService.PeriodColumn> cols)
    {
        var rows = new List<BmesDailyReportModelRow>();
        if (cols.Count == 0) return rows;

        foreach (SubGroupUnit unit in units)
        {
            List<double> trend = cols.Select(c => lookup.Sub1(unit.Key, c.Key)).ToList();
            double recent = trend[0];

            // No PPM in the latest period means this sub group was not produced then. It
            // is not "improved to zero" — there is nothing to report on, and leaving it in
            // would fill the Improved list with lines that only mean "production stopped".
            if (recent <= 0) continue;

            int baseline = -1;
            for (int i = 1; i < cols.Count; i++)
            {
                if (!production.Produced(unit.Key, cols[i].Key)) continue;
                baseline = i;
                break;
            }

            double previous = baseline >= 0 ? trend[baseline] : 0;
            double delta = recent - previous;
            double deltaPercent = previous > 0
                ? delta / previous * 100
                : (recent > 0 ? 100 : 0);

            (List<BmesDailyReportCause> causes, long ngQty) = AnalyzeCauses(
                production, unit.Key, cols[0].Key, baseline >= 0 ? cols[baseline].Key : null);

            rows.Add(new BmesDailyReportModelRow(
                unit.Name, unit.Group, unit.Parent, recent, previous, delta, deltaPercent,
                trend, causes, ngQty,
                BaselineHeader:    baseline >= 0 ? cols[baseline].Header : string.Empty,
                BaselineStepsBack: baseline >= 0 ? baseline : 0));
        }

        return rows;
    }

    /// <summary>Drop the leading columns that carry no data at all, so index 0 is the
    /// latest period something was actually produced in.
    ///
    /// BMES emits a period column as soon as it has *rows* for it, and a Sunday or a
    /// holiday still produces one: the day-over-day ranking anchored on it compared every
    /// model against an empty day, found no model with a defect rate to report, and came
    /// out completely empty. Weeks go through the same trim for consistency; there the
    /// leading column normally has data anyway.</summary>
    private static List<NgRateReportService.PeriodColumn> TrimToLatestWithData(
        NgRateReportService.SummaryPivotRow?   total,
        List<NgRateReportService.PeriodColumn> cols)
    {
        // No TOTAL row to test against: leave the columns alone rather than blanking the
        // whole dashboard on a report shape this method cannot judge.
        if (total is null) return cols;

        for (int i = 0; i < cols.Count; i++)
        {
            if ((total?.Ppm.GetValueOrDefault(cols[i].Key) ?? 0) > 0)
                return i == 0 ? cols : cols.Skip(i).ToList();
        }

        return [];
    }

    /// <summary>The overall (all-models) baseline for one granularity: the most recent
    /// earlier column the TOTAL row actually carries a number in.</summary>
    private static (double Ppm, string Header) OverallBaseline(
        NgRateReportService.SummaryPivotRow?   total,
        List<NgRateReportService.PeriodColumn> cols)
    {
        for (int i = 1; i < cols.Count; i++)
        {
            double ppm = total?.Ppm.GetValueOrDefault(cols[i].Key) ?? 0;
            if (ppm > 0) return (ppm, cols[i].Header);
        }

        return (0, string.Empty);
    }

    /// <summary>One row's identity: the Sub1 key the report aggregates under, plus the
    /// labels needed to show where it sits in the tree.</summary>
    private sealed record SubGroupUnit(string Key, string Name, string Group, string Parent);

    /// <summary>Break one sub group's latest period down into its top defect drivers.
    ///
    /// Reads <see cref="NgRateReportService.NgRateReport.GroupRawIn"/> — raw
    /// (input, ng) counts keyed at (ProcessType, ProcessName, NgName, Group,
    /// Period) — rather than the report's Top-10 NG table, because that table is
    /// the global top 10 across the whole selection and would miss a driver that
    /// matters only for this one model.</summary>
    /// <returns>The biggest causes of the latest period — ranked by its PPM, each carrying
    /// its move against the row's baseline period — plus the model's defect count.</returns>
    private static (List<BmesDailyReportCause> Causes, long NgQty) AnalyzeCauses(
        ProductionIndex production,
        string  modelKey,
        string  recentKey,
        string? previousKey)
    {
        Dictionary<CauseKey, (double Ppm, double NgQty)> recent =
            CausePpm(production, modelKey, recentKey);

        Dictionary<CauseKey, (double Ppm, double NgQty)> previous =
            previousKey is null
                ? []
                : CausePpm(production, modelKey, previousKey);

        if (recent.Count == 0) return ([], 0);

        long totalNg = (long)Math.Round(recent.Values.Sum(v => v.NgQty));

        List<BmesDailyReportCause> causes = recent
            .OrderByDescending(kv => kv.Value.Ppm)
            .Take(TopCauseCount)
            .Select(kv =>
            {
                double prevPpm = previous.GetValueOrDefault(kv.Key).Ppm;
                double delta = kv.Value.Ppm - prevPpm;
                return new BmesDailyReportCause(
                    NgName:       kv.Key.Ng,
                    ProcessName:  kv.Key.Process,
                    ProcessType:  kv.Key.ProcessType,
                    Ppm:          kv.Value.Ppm,
                    PreviousPpm:  prevPpm,
                    DeltaPpm:     delta,
                    DeltaPercent: prevPpm > 0 ? delta / prevPpm * 100 : 0,
                    IsNew:        prevPpm <= 0);
            })
            .ToList();

        return (causes, totalNg);
    }

    /// <summary>Identity of one defect driver: the defect, the process it was caught
    /// in, and that process's class (MAIN / SUB / …).</summary>
    private readonly record struct CauseKey(string Ng, string Process, string ProcessType);

    /// <summary>One unit's per-cause PPM and defect count for a single period.</summary>
    private static Dictionary<CauseKey, (double Ppm, double NgQty)> CausePpm(
        ProductionIndex production,
        string modelKey,
        string periodKey)
    {
        var result = new Dictionary<CauseKey, (double Ppm, double NgQty)>();

        foreach (RawCause raw in production.Causes(modelKey, periodKey))
        {
            if (raw.Ng <= 0) continue;

            (double ppm, double qty) prior = result.GetValueOrDefault(raw.Key);
            result[raw.Key] = (prior.ppm + raw.Ng / raw.Input * 1_000_000d, prior.qty + raw.Ng);
        }

        return result;
    }

    /// <summary>One raw (input, ng) entry of the report's finest aggregation, already
    /// bucketed by unit and period. <c>Input</c> is always &gt; 0 (see
    /// <see cref="ProductionIndex.Build"/>).</summary>
    private readonly record struct RawCause(CauseKey Key, double Input, double Ng);

    /// <summary>The report's raw (input, ng) map re-keyed by the two things this dashboard
    /// asks it about: the sub group and the period.
    ///
    /// Besides making the lookups O(1), the index answers the question a PPM value cannot:
    /// a PPM of 0 means both "produced with no defects" and "not produced at all", while a
    /// bucket exists here only when the unit actually had input in that period.</summary>
    private sealed class ProductionIndex
    {
        private static readonly RawCause[] None = [];

        private readonly Dictionary<(string Unit, string Period), List<RawCause>> _buckets;

        private ProductionIndex(Dictionary<(string, string), List<RawCause>> buckets) => _buckets = buckets;

        public static ProductionIndex Build(NgRateReportService.NgRateReport report)
        {
            var buckets = new Dictionary<(string Unit, string Period), List<RawCause>>();

            foreach (var kv in report.GroupRawIn)
            {
                (double input, double ng) = kv.Value;
                if (input <= 0) continue;

                (string, string) bucket = (kv.Key.Group, kv.Key.PeriodKey);
                if (!buckets.TryGetValue(bucket, out List<RawCause>? entries))
                    buckets[bucket] = entries = [];

                entries.Add(new RawCause(new CauseKey(kv.Key.NG, kv.Key.PN, kv.Key.PT), input, ng));
            }

            return new ProductionIndex(buckets);
        }

        /// <summary>True when the unit has any recorded input in that period — i.e. there
        /// is data to compare against, whatever its defect rate turned out to be.</summary>
        public bool Produced(string unitKey, string periodKey) =>
            _buckets.ContainsKey((unitKey, periodKey));

        public IReadOnlyList<RawCause> Causes(string unitKey, string periodKey) =>
            _buckets.GetValueOrDefault((unitKey, periodKey)) ?? (IReadOnlyList<RawCause>)None;
    }

    /// <summary>Reuse the group selection the Report page saved, so the dashboard
    /// covers exactly what the team already reports on. No saved selection → all
    /// groups (<see cref="ModelGroupPickerHelpers.ApplyPickerSelection"/> treats an
    /// empty selection as "show all").</summary>
    private List<ModelGroupRecord> LoadSelectedGroups()
    {
        List<ModelGroupRecord> allGroups = _repo.GetModelGroups();
        var groupIds   = new HashSet<long>();
        var subEntries = new HashSet<string>(StringComparer.Ordinal);

        string? json = _repo.GetSetting(ReportSelectionKey);
        if (!string.IsNullOrWhiteSpace(json))
        {
            try
            {
                SelectionState? state = JsonSerializer.Deserialize<SelectionState>(json, JsonOptions);
                var validIds = allGroups.Select(g => g.Id).ToHashSet();
                foreach (SelectionRoot root in state?.RootGroups ?? [])
                    if (validIds.Contains(root.GroupId)) groupIds.Add(root.GroupId);

                var validSubKeys = allGroups
                    .SelectMany(g => ModelGroupPickerHelpers.EnumerateChildEntries(g).Select(e => e.Key))
                    .ToHashSet(StringComparer.Ordinal);
                foreach (SelectionSubGroup sub in state?.SubGroups ?? [])
                    if (!string.IsNullOrWhiteSpace(sub.SelectionKey) && validSubKeys.Contains(sub.SelectionKey))
                        subEntries.Add(sub.SelectionKey);
            }
            catch
            {
                // Unreadable selection → fall through to "all groups".
            }
        }

        return ModelGroupPickerHelpers.ApplyPickerSelection(allGroups, groupIds, subEntries);
    }

    private void AddLog(string message)
    {
        lock (_gate)
        {
            _logs.Add($"[{DateTime.Now:HH:mm:ss}] {message}");
            if (_logs.Count > 400) _logs.RemoveRange(0, _logs.Count - 400);
        }
    }

    // Mirrors the shape BmesReportPage persists under BmesReport:SelectionState.
    private sealed record SelectionState(List<SelectionRoot> RootGroups, List<SelectionSubGroup> SubGroups);
    private sealed record SelectionRoot(long GroupId, string GroupName);
    private sealed record SelectionSubGroup(long GroupId, string SelectionKey, string GroupName);
}
