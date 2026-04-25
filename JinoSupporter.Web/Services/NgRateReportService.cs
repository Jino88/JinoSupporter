using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

/// <summary>
/// NG Rate pivot report service (same behavior as WPF DataMaker).
/// - PPM calc: compute PPM per (ProcessName, NgName, Date) combo, then sum.
/// - TOTAL = sum of PPM across all ProcessTypes.
/// - Column layout: dates (desc) | blank | weeks (desc) | blank | months (desc).
/// </summary>
public sealed class NgRateReportService(NgRateSettingsService settings)
{
    private readonly NgRateSettingsService _settings = settings;

    // ── Public types ──────────────────────────────────────────────────────────────

    public sealed record PeriodColumn(string Key, string Header);

    public sealed class SummaryPivotRow
    {
        public required string ProcessType { get; init; }
        public bool IsTotal { get; init; }
        public required Dictionary<string, double> Ppm { get; init; }
    }

    public sealed class GroupPivotRow
    {
        public required string GroupName { get; init; }
        public required Dictionary<string, double> Ppm { get; init; }
    }

    public sealed class ProcessPivotRow
    {
        public int    Rank        { get; init; }
        public required string ProcessType { get; init; }
        public required string ProcessName { get; init; }
        public double TotalPpm    { get; init; }
        public required Dictionary<string, double> Ppm { get; init; }
        public List<GroupPivotRow> Groups { get; init; } = [];
    }

    public sealed class NgPivotRow
    {
        public int    Rank        { get; init; }
        public required string NgName      { get; init; }
        public required string ProcessName { get; init; }
        public required string ProcessType { get; init; }
        public double TotalPpm    { get; init; }
        public required Dictionary<string, double> Ppm { get; init; }
        public List<GroupPivotRow> Groups { get; init; } = [];
    }

    public sealed class InputQtyRow
    {
        public int    No          { get; init; }
        public string ProcessType { get; init; } = string.Empty;
        public string ProcessName { get; init; } = string.Empty;
        public string GroupName   { get; init; } = string.Empty;
        public Dictionary<string, long> Qty { get; init; } = [];
    }

    public sealed class NgQtyRow
    {
        public int    No          { get; init; }
        public string ProcessType { get; init; } = string.Empty;
        public string ProcessName { get; init; } = string.Empty;
        public string NgName      { get; init; } = string.Empty;
        public string GroupName   { get; init; } = string.Empty;
        public Dictionary<string, long> Qty { get; init; } = [];
    }

    public sealed class ReasonRow
    {
        public string Reason      { get; init; } = string.Empty;
        public int    No          { get; init; }
        public bool   IsTotal     { get; init; }
        public string ProcessType { get; init; } = string.Empty;
        public string ProcessName { get; init; } = string.Empty;
        public string NgName      { get; init; } = string.Empty;
        public Dictionary<string, double> Ppm    { get; init; } = [];
        public List<GroupPivotRow>        Groups { get; init; } = [];
    }

    /// <summary>Detail row keyed by (LineShift, ProcessType, ProcessName, NgName) with per-period PPM dicts.</summary>
    public sealed class LineShiftNgDetail
    {
        public required string LineShift   { get; init; }
        public required string ProcessType { get; init; }
        public required string ProcessName { get; init; }
        public required string NgName      { get; init; }
        public required Dictionary<string, double> DatePpm  { get; init; }
        public required Dictionary<string, double> WeekPpm  { get; init; }
        public required Dictionary<string, double> MonthPpm { get; init; }
    }

    /// <summary>Per-(group, period) raw Input/NG quantities — used to compute weighted sumNg/sumInput PPM.</summary>
    public sealed class GroupRawQtyReport
    {
        public List<PeriodColumn> DateCols  { get; init; } = [];
        public List<PeriodColumn> WeekCols  { get; init; } = [];
        public List<PeriodColumn> MonthCols { get; init; } = [];
        /// <summary>group → (period → (sumInput, sumNg))</summary>
        public Dictionary<string, Dictionary<string, (double Input, double Ng)>> Date  { get; init; } = new();
        public Dictionary<string, Dictionary<string, (double Input, double Ng)>> Week  { get; init; } = new();
        public Dictionary<string, Dictionary<string, (double Input, double Ng)>> Month { get; init; } = new();
    }

    public sealed class LineShiftNgReport
    {
        public List<PeriodColumn>       DateCols  { get; init; } = [];
        public List<PeriodColumn>       WeekCols  { get; init; } = [];
        public List<PeriodColumn>       MonthCols { get; init; } = [];
        public List<LineShiftNgDetail>  Details   { get; init; } = [];
    }

    public sealed class NgRateReport
    {
        public List<PeriodColumn>    DateCols     { get; init; } = [];
        public List<PeriodColumn>    WeekCols     { get; init; } = [];
        public List<PeriodColumn>    MonthCols    { get; init; } = [];
        public List<SummaryPivotRow> Summary      { get; init; } = [];
        public List<ProcessPivotRow> Top10Process { get; init; } = [];
        public List<NgPivotRow>      Top10Ng      { get; init; } = [];
        public List<InputQtyRow>     InputQtyRows { get; init; } = [];
        public List<NgQtyRow>        NgQtyRows    { get; init; } = [];
        public List<ReasonRow>       ReasonRows   { get; init; } = [];
        public string   DbPath      { get; init; } = string.Empty;
        public int      TotalRows   { get; init; }
        public DateTime GeneratedAt { get; init; } = DateTime.Now;
        public List<string> ModelFilter  { get; init; } = [];
        public List<string> GroupNames   { get; init; } = [];
        /// <summary>GroupName → Summary by ProcessType (PPM)</summary>
        public Dictionary<string, List<SummaryPivotRow>> GroupSummary { get; init; } = new();

        /// <summary>
        /// Per-group raw input/NG counts keyed at the finest granularity we aggregate
        /// at: (ProcessType, ProcessName, NgName, Group, PeriodKey) → (Input, Ng).
        /// Exposed so the UI can correctly combine multiple selected groups using
        /// PPM = Σngs / Σinputs · 1M per (PT/PN/NG/period) instead of averaging or
        /// summing pre-computed per-group PPMs (both of which misreport the true
        /// merged defect rate when input sizes differ).
        /// </summary>
        public Dictionary<(string PT, string PN, string NG, string Group, string PeriodKey), (double I, double N)>
            GroupRawIn { get; init; } = new();
    }

    // ── Internal aggregation unit ─────────────────────────────────────────────────

    /// <summary>PPM keyed by (ProcessType, ProcessName, NgName, PeriodKey).</summary>
    private sealed record ComboAgg(
        string ProcessType, string ProcessName, string NgName,
        string PeriodKey, double Ppm);

    // ── Internal row type ─────────────────────────────────────────────────────────

    private sealed class OrgRow
    {
        public string MaterialName { get; set; } = "";
        public string LineShift    { get; set; } = "";
        public string ProcessCode  { get; set; } = "";
        public string ProcessName  { get; set; } = "";
        public string NgName       { get; set; } = "";
        public string ProcessType  { get; set; } = "";
        public string ProductDate  { get; set; } = "";
        public double QtyInput     { get; set; }
        public double QtyNg        { get; set; }
    }

    // ── Public API ────────────────────────────────────────────────────────────────

    public string? FindMostRecentDb()
    {
        string dir = _settings.DbSaveDirectory;
        if (!Directory.Exists(dir)) return null;
        return Directory.GetFiles(dir, "*.db")
            .Where(f => !Path.GetFileName(f).Equals("ngrate_settings.db",
                            StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(File.GetLastWriteTime)
            .FirstOrDefault();
    }

    /// <summary>
    /// Report generator. All callers (By Model, By Group) provide a pre-built
    /// (LineShift → GroupName) mapping sourced from the ModelGroups DB.
    /// </summary>
    public Task<NgRateReport> GenerateReportAsync(
        string dbPath,
        IReadOnlyDictionary<string, string> lineShiftToGroup,
        IReadOnlyList<string> groupNames,
        IProgress<string>? progress = null)
        => Task.Run(() =>
        {
            var modelLineShifts = lineShiftToGroup.Keys.ToList();
            var mapping         = new Dictionary<string, string>(lineShiftToGroup, StringComparer.Ordinal);
            var groups          = groupNames.ToList();
            return GenerateReportCore(dbPath, modelLineShifts, groups, mapping, progress);
        });

    /// <summary>
    /// Multi-valued overload: a single LineShift can belong to several groups at once
    /// (e.g. a root group, its named sub-groups, and the LineShift-as-its-own-group leaf).
    /// Each membership contributes a separate row to <c>GroupSummary</c>, so the Summary
    /// drill-down under every ProcessType exposes the full tree.
    /// </summary>
    public Task<NgRateReport> GenerateReportAsync(
        string dbPath,
        IReadOnlyDictionary<string, IReadOnlyList<string>> lineShiftToGroups,
        IReadOnlyList<string> groupNames,
        IProgress<string>? progress = null)
        => Task.Run(() =>
        {
            var modelLineShifts = lineShiftToGroups.Keys.ToList();
            var multi = lineShiftToGroups.ToDictionary(
                kv => kv.Key,
                kv => kv.Value.ToList(),
                StringComparer.Ordinal);
            // Single-valued projection used by legacy call sites — first membership wins.
            var primary = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var kv in multi)
                if (kv.Value.Count > 0) primary[kv.Key] = kv.Value[0];
            return GenerateReportCore(
                dbPath, modelLineShifts, groupNames.ToList(), primary, progress, multi);
        });

    // ── Report build ──────────────────────────────────────────────────────────────

    private NgRateReport GenerateReportCore(
        string dbPath,
        List<string> modelLineShifts,
        List<string> groupNames,
        Dictionary<string, string> lineShiftToGroup,
        IProgress<string>? progress,
        Dictionary<string, List<string>>? lineShiftToGroupsMulti = null)
    {
        // If caller didn't supply multi-mapping (e.g. the By-Group overload from
        // ModelGroups DB, which is inherently single-group), promote the single-valued
        // dict to a singleton-list form so all downstream helpers can use one path.
        var lsToGroups = lineShiftToGroupsMulti
            ?? lineShiftToGroup.ToDictionary(
                kv => kv.Key, kv => new List<string> { kv.Value }, StringComparer.Ordinal);

        var modelSet = modelLineShifts.Count > 0
            ? new HashSet<string>(modelLineShifts, StringComparer.Ordinal)
            : null;

        // 2. ProcessType lookup
        progress?.Report("Loading ProcessType lookup…");
        var ptLookup = LoadProcessTypeLookup(dbPath);

        // 3. Load OrginalTable
        progress?.Report("Loading OrginalTable…");
        var orgRows = LoadOrginalRows(dbPath, modelSet);
        progress?.Report($"  {orgRows.Count:N0} rows loaded.");
        if (orgRows.Count == 0)
            return new NgRateReport { DbPath = dbPath, ModelFilter = modelLineShifts, GroupNames = groupNames };

        // 4. Map ProcessType
        foreach (var r in orgRows)
        {
            if (!string.IsNullOrEmpty(r.ProcessType)) continue;
            if (ptLookup.TryGetValue((r.MaterialName, r.ProcessCode, r.ProcessName), out string? pt))
                r.ProcessType = pt;
            else if (ptLookup.TryGetValue((string.Empty, r.ProcessCode, r.ProcessName), out string? pt2))
                r.ProcessType = pt2;
        }

        // 5. Compute PPM per (combo, period) — WPF approach: sum of combo PPMs
        progress?.Report("Pre-aggregating by period…");
        var byDate  = BuildComboPpm(orgRows, r => r.ProductDate);
        var byWeek  = BuildComboPpm(orgRows, r => GetWeekKey(r.ProductDate));
        var byMonth = BuildComboPpm(orgRows, r => GetMonthKey(r.ProductDate));

        // 6. Extract period columns
        var dateCols  = ExtractCols(byDate,  k => FormatDateHeader(k));
        var weekCols  = ExtractCols(byWeek,  k => "W" + int.Parse(k[4..]));
        var monthCols = ExtractCols(byMonth, k => "M" + int.Parse(k[4..]));
        progress?.Report($"  {dateCols.Count} dates · {weekCols.Count} weeks · {monthCols.Count} months");

        // 7. Build aggregation lookup
        // (ProcessType, PeriodKey) → sumPpm
        var datePtMap  = BuildPtMap(byDate);
        var weekPtMap  = BuildPtMap(byWeek);
        var monthPtMap = BuildPtMap(byMonth);

        // total per period
        var dateTotalMap  = BuildTotalMap(byDate);
        var weekTotalMap  = BuildTotalMap(byWeek);
        var monthTotalMap = BuildTotalMap(byMonth);

        // (ProcessType, ProcessName, PeriodKey) → sumPpm  ← for Process Top 10 NG
        var dateProcessMap  = BuildProcessMap(byDate);
        var weekProcessMap  = BuildProcessMap(byWeek);
        var monthProcessMap = BuildProcessMap(byMonth);

        // (ProcessType, ProcessName, NgName, PeriodKey) → Ppm  ← for Worst 10 NG
        var dateNgMap  = BuildNgMap(byDate);
        var weekNgMap  = BuildNgMap(byWeek);
        var monthNgMap = BuildNgMap(byMonth);

        // 8. Summary
        progress?.Report("Computing Summary…");
        var summary = ComputeSummary(
            byDate, byWeek, byMonth,
            datePtMap, weekPtMap, monthPtMap,
            dateTotalMap, weekTotalMap, monthTotalMap,
            dateCols, weekCols, monthCols);

        // 9. Process 10 NG
        progress?.Report("Computing Top 10 Processes…");
        var dateGrpProc  = BuildGroupProcessRaw(orgRows, lsToGroups, r => r.ProductDate);
        var weekGrpProc  = BuildGroupProcessRaw(orgRows, lsToGroups, r => GetWeekKey(r.ProductDate));
        var monthGrpProc = BuildGroupProcessRaw(orgRows, lsToGroups, r => GetMonthKey(r.ProductDate));
        var top10Process = ComputeTop10Process(
            byDate, byWeek, byMonth,
            dateProcessMap, weekProcessMap, monthProcessMap,
            dateGrpProc, weekGrpProc, monthGrpProc,
            dateCols, weekCols, monthCols);

        // 10. Worst 10 NG
        progress?.Report("Computing Worst 10 NG Names…");
        var dateGrpNg  = BuildGroupNgRaw(orgRows, lsToGroups, r => r.ProductDate);
        var weekGrpNg  = BuildGroupNgRaw(orgRows, lsToGroups, r => GetWeekKey(r.ProductDate));
        var monthGrpNg = BuildGroupNgRaw(orgRows, lsToGroups, r => GetMonthKey(r.ProductDate));
        var top10Ng = ComputeTop10Ng(
            byDate, byWeek, byMonth,
            dateNgMap, weekNgMap, monthNgMap,
            dateGrpNg, weekGrpNg, monthGrpNg,
            dateCols, weekCols, monthCols);

        // 11. Per-group input/NG quantities & Reason report
        progress?.Report("Computing group detail rows…");
        var inputQtyRows = BuildInputQtyRows(orgRows, lsToGroups, dateCols);
        var ngQtyRows    = BuildNgQtyRows(orgRows, lsToGroups, dateCols);
        var reasonLookup = LoadReasonLookup(dbPath);
        var reasonRows   = BuildReasonRows(orgRows, reasonLookup, lsToGroups, dateCols, weekCols, monthCols);

        // 12. Per-group summary (each group's defect rate shown on model click)
        var groupSummary = groupNames.Count > 0
            ? ComputeGroupSummary(groupNames, dateGrpProc, weekGrpProc, monthGrpProc, dateCols, weekCols, monthCols)
            : new Dictionary<string, List<SummaryPivotRow>>();

        // 13. Per-group raw Input/NG at finest granularity — merged across period kinds.
        // Used by the UI to correctly combine multiple selected chips via
        // PPM = Σng / Σinput · 1M (not by summing or averaging pre-computed group PPMs).
        var groupRawIn = new Dictionary<(string PT, string PN, string NG, string G, string PeriodKey), (double I, double N)>();
        foreach (var kv in dateGrpNg)
            groupRawIn[(kv.Key.Item1, kv.Key.Item2, kv.Key.Item3, kv.Key.Item4, kv.Key.Item5)] = kv.Value;
        foreach (var kv in weekGrpNg)
            groupRawIn[(kv.Key.Item1, kv.Key.Item2, kv.Key.Item3, kv.Key.Item4, kv.Key.Item5)] = kv.Value;
        foreach (var kv in monthGrpNg)
            groupRawIn[(kv.Key.Item1, kv.Key.Item2, kv.Key.Item3, kv.Key.Item4, kv.Key.Item5)] = kv.Value;

        progress?.Report("Report complete.");
        return new NgRateReport
        {
            DateCols     = dateCols,
            WeekCols     = weekCols,
            MonthCols    = monthCols,
            Summary      = summary,
            Top10Process = top10Process,
            Top10Ng      = top10Ng,
            InputQtyRows = inputQtyRows,
            NgQtyRows    = ngQtyRows,
            ReasonRows   = reasonRows,
            DbPath       = dbPath,
            TotalRows    = orgRows.Count,
            ModelFilter  = modelLineShifts,
            GroupNames   = groupNames,
            GroupSummary = groupSummary,
            GroupRawIn   = groupRawIn,
        };
    }

    // ── Per-group weighted (sumNg/sumInput) raw aggregation ──────────────────────

    public Task<GroupRawQtyReport> ComputeGroupRawQtyAsync(
        string dbPath,
        IReadOnlyDictionary<string, string> lineShiftToGroup,
        IProgress<string>? progress = null)
        => Task.Run(() => ComputeGroupRawQty(dbPath, lineShiftToGroup, progress));

    private GroupRawQtyReport ComputeGroupRawQty(
        string dbPath,
        IReadOnlyDictionary<string, string> lineShiftToGroup,
        IProgress<string>? progress)
    {
        if (lineShiftToGroup.Count == 0) return new GroupRawQtyReport();

        var modelSet = new HashSet<string>(lineShiftToGroup.Keys, StringComparer.Ordinal);
        progress?.Report("Loading OrginalTable (group raw qty)…");
        var orgRows = LoadOrginalRows(dbPath, modelSet);
        if (orgRows.Count == 0) return new GroupRawQtyReport();

        static Dictionary<string, Dictionary<string, (double Input, double Ng)>> BuildByPeriod(
            List<OrgRow> rows,
            IReadOnlyDictionary<string, string> lsToGroup,
            Func<OrgRow, string> getKey)
        {
            // Input은 같은 (LS, PT, PN, period) 안에서 NgName마다 중복되어 있을 수 있으므로
            // NgName 차원을 먼저 접어 Input은 MAX로 한 번만 카운트하고 Ng는 합산.
            var collapsed = rows
                .Where(r => !string.IsNullOrEmpty(getKey(r)))
                .Where(r => !string.IsNullOrEmpty(r.LineShift) && lsToGroup.ContainsKey(r.LineShift))
                .GroupBy(r => (r.LineShift, r.ProcessType, r.ProcessName, K: getKey(r)))
                .Select(g => (
                    LineShift: g.Key.LineShift,
                    K:         g.Key.K,
                    Input:     g.Max(r => r.QtyInput),
                    Ng:        g.Sum(r => r.QtyNg)))
                .ToList();

            var outer = new Dictionary<string, Dictionary<string, (double, double)>>(StringComparer.Ordinal);
            foreach (var x in collapsed)
            {
                string grp = lsToGroup[x.LineShift];
                if (!outer.TryGetValue(grp, out var inner))
                {
                    inner = new Dictionary<string, (double, double)>(StringComparer.Ordinal);
                    outer[grp] = inner;
                }
                var (i, n) = inner.GetValueOrDefault(x.K);
                inner[x.K] = (i + x.Input, n + x.Ng);
            }
            return outer;
        }

        var date  = BuildByPeriod(orgRows, lineShiftToGroup, r => r.ProductDate);
        var week  = BuildByPeriod(orgRows, lineShiftToGroup, r => GetWeekKey(r.ProductDate));
        var month = BuildByPeriod(orgRows, lineShiftToGroup, r => GetMonthKey(r.ProductDate));

        var dateKeys  = date .Values.SelectMany(d => d.Keys).ToHashSet(StringComparer.Ordinal);
        var weekKeys  = week .Values.SelectMany(d => d.Keys).ToHashSet(StringComparer.Ordinal);
        var monthKeys = month.Values.SelectMany(d => d.Keys).ToHashSet(StringComparer.Ordinal);

        var dateCols  = dateKeys .OrderByDescending(k => k).Select(k => new PeriodColumn(k, FormatDateHeader(k))).ToList();
        var weekCols  = weekKeys .OrderByDescending(k => k).Select(k => new PeriodColumn(k, "W" + int.Parse(k[4..]))).ToList();
        var monthCols = monthKeys.OrderByDescending(k => k).Select(k => new PeriodColumn(k, "M" + int.Parse(k[4..]))).ToList();

        return new GroupRawQtyReport
        {
            DateCols  = dateCols,
            WeekCols  = weekCols,
            MonthCols = monthCols,
            Date      = date,
            Week      = week,
            Month     = month,
        };
    }

    // ── Per-LineShift × Process × NG detail aggregation ──────────────────────────

    /// <summary>
    /// Returns per-(LineShift, ProcessType, ProcessName, NgName) PPM broken down by
    /// date / week / month. Used by the "By Group" view's top-10 per group table.
    /// </summary>
    public Task<LineShiftNgReport> ComputeLineShiftNgDetailsAsync(
        string dbPath,
        IReadOnlyCollection<string> lineShifts,
        IProgress<string>? progress = null)
        => Task.Run(() => ComputeLineShiftNgDetails(dbPath, lineShifts, progress));

    /// <summary>
    /// Same as ComputeLineShiftNgDetailsAsync but aggregates LineShifts per supplied mapping.
    /// Returned `LineShiftNgDetail.LineShift` holds the group key (e.g. "대그룹::Material").
    /// PPMs are combo-merged (weighted) per (group, PT, PN, NgName, period).
    /// </summary>
    public Task<LineShiftNgReport> ComputeGroupedNgDetailsAsync(
        string dbPath,
        IReadOnlyDictionary<string, string> lineShiftToGroup,
        IProgress<string>? progress = null)
        => Task.Run(() => ComputeGroupedNgDetails(dbPath, lineShiftToGroup, progress));

    private LineShiftNgReport ComputeGroupedNgDetails(
        string dbPath,
        IReadOnlyDictionary<string, string> lineShiftToGroup,
        IProgress<string>? progress)
    {
        if (lineShiftToGroup.Count == 0) return new LineShiftNgReport();

        var modelSet = new HashSet<string>(lineShiftToGroup.Keys, StringComparer.Ordinal);
        progress?.Report("Loading ProcessType lookup (grouped NG detail)…");
        var ptLookup = LoadProcessTypeLookup(dbPath);

        progress?.Report("Loading OrginalTable (grouped NG detail)…");
        var orgRows = LoadOrginalRows(dbPath, modelSet);
        if (orgRows.Count == 0) return new LineShiftNgReport();

        foreach (var r in orgRows)
        {
            if (!string.IsNullOrEmpty(r.ProcessType)) continue;
            if (ptLookup.TryGetValue((r.MaterialName, r.ProcessCode, r.ProcessName), out string? pt))
                r.ProcessType = pt;
            else if (ptLookup.TryGetValue((string.Empty, r.ProcessCode, r.ProcessName), out string? pt2))
                r.ProcessType = pt2;
        }

        progress?.Report("Aggregating by (Group, Process, NG)…");
        var details = new List<LineShiftNgDetail>();
        foreach (var g in orgRows
                     .Where(r => lineShiftToGroup.ContainsKey(r.LineShift))
                     .GroupBy(r => (Grp: lineShiftToGroup[r.LineShift],
                                    r.ProcessType, r.ProcessName, r.NgName)))
        {
            var datePpm  = g.GroupBy(r => r.ProductDate)
                .Where(x => !string.IsNullOrEmpty(x.Key))
                .ToDictionary(x => x.Key, x => CalcPpm(x.Sum(r => r.QtyInput), x.Sum(r => r.QtyNg)));
            var weekPpm  = g.GroupBy(r => GetWeekKey(r.ProductDate))
                .Where(x => !string.IsNullOrEmpty(x.Key))
                .ToDictionary(x => x.Key, x => CalcPpm(x.Sum(r => r.QtyInput), x.Sum(r => r.QtyNg)));
            var monthPpm = g.GroupBy(r => GetMonthKey(r.ProductDate))
                .Where(x => !string.IsNullOrEmpty(x.Key))
                .ToDictionary(x => x.Key, x => CalcPpm(x.Sum(r => r.QtyInput), x.Sum(r => r.QtyNg)));

            details.Add(new LineShiftNgDetail
            {
                LineShift   = g.Key.Grp,      // now holds the group key
                ProcessType = g.Key.ProcessType,
                ProcessName = g.Key.ProcessName,
                NgName      = g.Key.NgName,
                DatePpm     = datePpm,
                WeekPpm     = weekPpm,
                MonthPpm    = monthPpm,
            });
        }

        var dateKeys  = details.SelectMany(d => d.DatePpm.Keys).ToHashSet(StringComparer.Ordinal);
        var weekKeys  = details.SelectMany(d => d.WeekPpm.Keys).ToHashSet(StringComparer.Ordinal);
        var monthKeys = details.SelectMany(d => d.MonthPpm.Keys).ToHashSet(StringComparer.Ordinal);

        var dateCols  = dateKeys .OrderByDescending(k => k).Select(k => new PeriodColumn(k, FormatDateHeader(k))).ToList();
        var weekCols  = weekKeys .OrderByDescending(k => k).Select(k => new PeriodColumn(k, "W" + int.Parse(k[4..]))).ToList();
        var monthCols = monthKeys.OrderByDescending(k => k).Select(k => new PeriodColumn(k, "M" + int.Parse(k[4..]))).ToList();

        progress?.Report($"  {details.Count:N0} (Group×Process×NG) combos.");
        return new LineShiftNgReport
        {
            DateCols = dateCols, WeekCols = weekCols, MonthCols = monthCols, Details = details,
        };
    }

    private LineShiftNgReport ComputeLineShiftNgDetails(
        string dbPath,
        IReadOnlyCollection<string> lineShifts,
        IProgress<string>? progress)
    {
        var modelSet = lineShifts.Count > 0
            ? new HashSet<string>(lineShifts, StringComparer.Ordinal)
            : null;

        progress?.Report("Loading ProcessType lookup (LS×NG detail)…");
        var ptLookup = LoadProcessTypeLookup(dbPath);

        progress?.Report("Loading OrginalTable (LS×NG detail)…");
        var orgRows = LoadOrginalRows(dbPath, modelSet);
        if (orgRows.Count == 0) return new LineShiftNgReport();

        foreach (var r in orgRows)
        {
            if (!string.IsNullOrEmpty(r.ProcessType)) continue;
            if (ptLookup.TryGetValue((r.MaterialName, r.ProcessCode, r.ProcessName), out string? pt))
                r.ProcessType = pt;
            else if (ptLookup.TryGetValue((string.Empty, r.ProcessCode, r.ProcessName), out string? pt2))
                r.ProcessType = pt2;
        }

        progress?.Report("Aggregating by (LineShift, Process, NG)…");
        var details = new List<LineShiftNgDetail>();
        foreach (var g in orgRows
                     .Where(r => !string.IsNullOrEmpty(r.LineShift))
                     .GroupBy(r => (r.LineShift, r.ProcessType, r.ProcessName, r.NgName)))
        {
            var datePpm  = g.GroupBy(r => r.ProductDate)
                .Where(x => !string.IsNullOrEmpty(x.Key))
                .ToDictionary(x => x.Key, x => CalcPpm(x.Sum(r => r.QtyInput), x.Sum(r => r.QtyNg)));
            var weekPpm  = g.GroupBy(r => GetWeekKey(r.ProductDate))
                .Where(x => !string.IsNullOrEmpty(x.Key))
                .ToDictionary(x => x.Key, x => CalcPpm(x.Sum(r => r.QtyInput), x.Sum(r => r.QtyNg)));
            var monthPpm = g.GroupBy(r => GetMonthKey(r.ProductDate))
                .Where(x => !string.IsNullOrEmpty(x.Key))
                .ToDictionary(x => x.Key, x => CalcPpm(x.Sum(r => r.QtyInput), x.Sum(r => r.QtyNg)));

            details.Add(new LineShiftNgDetail
            {
                LineShift   = g.Key.LineShift,
                ProcessType = g.Key.ProcessType,
                ProcessName = g.Key.ProcessName,
                NgName      = g.Key.NgName,
                DatePpm     = datePpm,
                WeekPpm     = weekPpm,
                MonthPpm    = monthPpm,
            });
        }

        // Collect distinct period keys across all details.
        var dateKeys  = details.SelectMany(d => d.DatePpm.Keys).ToHashSet(StringComparer.Ordinal);
        var weekKeys  = details.SelectMany(d => d.WeekPpm.Keys).ToHashSet(StringComparer.Ordinal);
        var monthKeys = details.SelectMany(d => d.MonthPpm.Keys).ToHashSet(StringComparer.Ordinal);

        var dateCols  = dateKeys .OrderByDescending(k => k).Select(k => new PeriodColumn(k, FormatDateHeader(k))).ToList();
        var weekCols  = weekKeys .OrderByDescending(k => k).Select(k => new PeriodColumn(k, "W" + int.Parse(k[4..]))).ToList();
        var monthCols = monthKeys.OrderByDescending(k => k).Select(k => new PeriodColumn(k, "M" + int.Parse(k[4..]))).ToList();

        progress?.Report($"  {details.Count:N0} (LS×Process×NG) combos.");
        return new LineShiftNgReport
        {
            DateCols  = dateCols,
            WeekCols  = weekCols,
            MonthCols = monthCols,
            Details   = details,
        };
    }

    // ── Core aggregation builder ──────────────────────────────────────────────────

    /// <summary>
    /// Compute PPM per (ProcessType, ProcessName, NgName, PeriodKey).
    /// Same as WPF: PPM of each combo = QTYNG / QTYINPUT * 1,000,000.
    /// </summary>
    private static List<ComboAgg> BuildComboPpm(List<OrgRow> rows, Func<OrgRow, string> getKey)
    {
        return rows
            .GroupBy(r => (r.ProcessType, r.ProcessName, r.NgName, Key: getKey(r)))
            .Where(g => !string.IsNullOrEmpty(g.Key.Key))
            .Select(g => new ComboAgg(
                g.Key.ProcessType,
                g.Key.ProcessName,
                g.Key.NgName,
                g.Key.Key,
                CalcPpm(g.Sum(r => r.QtyInput), g.Sum(r => r.QtyNg))))
            .ToList();
    }

    private static List<PeriodColumn> ExtractCols(List<ComboAgg> agg, Func<string, string> headerOf)
    {
        return agg
            .Select(x => x.PeriodKey)
            .ToHashSet()
            .OrderByDescending(k => k)
            .Select(k => new PeriodColumn(k, headerOf(k)))
            .ToList();
    }

    // (ProcessType, PeriodKey) → Σ Ppm
    private static Dictionary<(string, string), double> BuildPtMap(List<ComboAgg> agg)
        => agg
           .GroupBy(x => (x.ProcessType, x.PeriodKey))
           .ToDictionary(g => g.Key, g => g.Sum(x => x.Ppm));

    // PeriodKey → Σ Ppm (TOTAL)
    private static Dictionary<string, double> BuildTotalMap(List<ComboAgg> agg)
        => agg
           .GroupBy(x => x.PeriodKey)
           .ToDictionary(g => g.Key, g => g.Sum(x => x.Ppm));

    // (ProcessType, ProcessName, PeriodKey) → Σ Ppm
    private static Dictionary<(string, string, string), double> BuildProcessMap(List<ComboAgg> agg)
        => agg
           .GroupBy(x => (x.ProcessType, x.ProcessName, x.PeriodKey))
           .ToDictionary(g => g.Key, g => g.Sum(x => x.Ppm));

    // (ProcessType, ProcessName, NgName, PeriodKey) → Ppm
    private static Dictionary<(string, string, string, string), double> BuildNgMap(List<ComboAgg> agg)
    {
        var dict = new Dictionary<(string, string, string, string), double>();
        foreach (var x in agg)
            dict[(x.ProcessType, x.ProcessName, x.NgName, x.PeriodKey)] = x.Ppm;
        return dict;
    }

    // ── Summary ──────────────────────────────────────────────────────────────────

    private static List<SummaryPivotRow> ComputeSummary(
        List<ComboAgg> byDate, List<ComboAgg> byWeek, List<ComboAgg> byMonth,
        Dictionary<(string, string), double> datePtMap,
        Dictionary<(string, string), double> weekPtMap,
        Dictionary<(string, string), double> monthPtMap,
        Dictionary<string, double> dateTotalMap,
        Dictionary<string, double> weekTotalMap,
        Dictionary<string, double> monthTotalMap,
        List<PeriodColumn> dateCols,
        List<PeriodColumn> weekCols,
        List<PeriodColumn> monthCols)
    {
        var result = new List<SummaryPivotRow>();

        // TOTAL
        result.Add(new SummaryPivotRow
        {
            ProcessType = "TOTAL",
            IsTotal     = true,
            Ppm         = BuildPeriodDict(dateCols, weekCols, monthCols,
                d => dateTotalMap.GetValueOrDefault(d),
                w => weekTotalMap.GetValueOrDefault(w),
                m => monthTotalMap.GetValueOrDefault(m)),
        });

        // By ProcessType (WPF order: SUB → MAIN → FUNCTION → VISUAL → others)
        var types = byDate
            .Select(x => x.ProcessType)
            .Concat(byWeek.Select(x => x.ProcessType))
            .Concat(byMonth.Select(x => x.ProcessType))
            .Where(t => !string.IsNullOrEmpty(t))
            .ToHashSet()
            .OrderBy(ProcessTypeOrder)
            .ThenBy(t => t);

        foreach (string pt in types)
        {
            result.Add(new SummaryPivotRow
            {
                ProcessType = pt,
                IsTotal     = false,
                Ppm         = BuildPeriodDict(dateCols, weekCols, monthCols,
                    d => datePtMap.GetValueOrDefault((pt, d)),
                    w => weekPtMap.GetValueOrDefault((pt, w)),
                    m => monthPtMap.GetValueOrDefault((pt, m))),
            });
        }

        return result;
    }

    // ── Top 10 Process ────────────────────────────────────────────────────────────

    private static List<ProcessPivotRow> ComputeTop10Process(
        List<ComboAgg> byDate, List<ComboAgg> byWeek, List<ComboAgg> byMonth,
        Dictionary<(string, string, string), double> dateProcMap,
        Dictionary<(string, string, string), double> weekProcMap,
        Dictionary<(string, string, string), double> monthProcMap,
        Dictionary<(string, string, string, string), double> dateGrpProc,
        Dictionary<(string, string, string, string), double> weekGrpProc,
        Dictionary<(string, string, string, string), double> monthGrpProc,
        List<PeriodColumn> dateCols,
        List<PeriodColumn> weekCols,
        List<PeriodColumn> monthCols)
    {
        // Sort desc by prior week (weekCols[1]); if no weeks, use the most recent
        var prevWeekKeyProc = weekCols.Count >= 2 ? weekCols[1].Key : weekCols.FirstOrDefault()?.Key ?? string.Empty;

        var allProcPpm = byDate
            .Concat(byWeek)
            .Concat(byMonth)
            .GroupBy(x => (x.ProcessType, x.ProcessName))
            .Select(g => (g.Key.ProcessType, g.Key.ProcessName,
                          Total: weekProcMap.GetValueOrDefault((g.Key.ProcessType, g.Key.ProcessName, prevWeekKeyProc))))
            .OrderByDescending(x => x.Total)
            .Take(10)
            .ToList();

        return allProcPpm.Select((x, i) =>
        {
            var groupNames = dateGrpProc.Keys
                .Where(k => k.Item1 == x.ProcessType && k.Item2 == x.ProcessName).Select(k => k.Item3)
                .Concat(weekGrpProc.Keys.Where(k => k.Item1 == x.ProcessType && k.Item2 == x.ProcessName).Select(k => k.Item3))
                .Concat(monthGrpProc.Keys.Where(k => k.Item1 == x.ProcessType && k.Item2 == x.ProcessName).Select(k => k.Item3))
                .Distinct().OrderBy(g => g).ToList();

            return new ProcessPivotRow
            {
                Rank        = i + 1,
                ProcessType = x.ProcessType,
                ProcessName = x.ProcessName,
                TotalPpm    = x.Total,
                Ppm         = BuildPeriodDict(dateCols, weekCols, monthCols,
                    d => dateProcMap.GetValueOrDefault((x.ProcessType, x.ProcessName, d)),
                    w => weekProcMap.GetValueOrDefault((x.ProcessType, x.ProcessName, w)),
                    m => monthProcMap.GetValueOrDefault((x.ProcessType, x.ProcessName, m))),
                Groups      = groupNames.Select(g => new GroupPivotRow
                {
                    GroupName = g,
                    Ppm       = BuildPeriodDict(dateCols, weekCols, monthCols,
                        d => dateGrpProc.GetValueOrDefault((x.ProcessType, x.ProcessName, g, d)),
                        w => weekGrpProc.GetValueOrDefault((x.ProcessType, x.ProcessName, g, w)),
                        m => monthGrpProc.GetValueOrDefault((x.ProcessType, x.ProcessName, g, m))),
                }).ToList(),
            };
        }).ToList();
    }

    // ── Worst 10 NG ───────────────────────────────────────────────────────────────

    private static List<NgPivotRow> ComputeTop10Ng(
        List<ComboAgg> byDate, List<ComboAgg> byWeek, List<ComboAgg> byMonth,
        Dictionary<(string, string, string, string), double> dateNgMap,
        Dictionary<(string, string, string, string), double> weekNgMap,
        Dictionary<(string, string, string, string), double> monthNgMap,
        Dictionary<(string, string, string, string, string), (double I, double N)> dateGrpNg,
        Dictionary<(string, string, string, string, string), (double I, double N)> weekGrpNg,
        Dictionary<(string, string, string, string, string), (double I, double N)> monthGrpNg,
        List<PeriodColumn> dateCols,
        List<PeriodColumn> weekCols,
        List<PeriodColumn> monthCols)
    {
        // Sort desc by prior week (weekCols[1]); if no weeks, use the most recent
        var prevWeekKeyNg = weekCols.Count >= 2 ? weekCols[1].Key : weekCols.FirstOrDefault()?.Key ?? string.Empty;

        var allNgPpm = byDate
            .Concat(byWeek)
            .Concat(byMonth)
            .GroupBy(x => (x.ProcessType, x.ProcessName, x.NgName))
            .Select(g => (g.Key.ProcessType, g.Key.ProcessName, g.Key.NgName,
                          Total: weekNgMap.GetValueOrDefault((g.Key.ProcessType, g.Key.ProcessName, g.Key.NgName, prevWeekKeyNg))))
            .OrderByDescending(x => x.Total)
            .Take(10)
            .ToList();

        return allNgPpm.Select((x, i) =>
        {
            var groupNames = dateGrpNg.Keys
                .Where(k => k.Item1 == x.ProcessType && k.Item2 == x.ProcessName && k.Item3 == x.NgName).Select(k => k.Item4)
                .Concat(weekGrpNg.Keys.Where(k => k.Item1 == x.ProcessType && k.Item2 == x.ProcessName && k.Item3 == x.NgName).Select(k => k.Item4))
                .Concat(monthGrpNg.Keys.Where(k => k.Item1 == x.ProcessType && k.Item2 == x.ProcessName && k.Item3 == x.NgName).Select(k => k.Item4))
                .Distinct().OrderBy(g => g).ToList();

            return new NgPivotRow
            {
                Rank        = i + 1,
                NgName      = x.NgName,
                ProcessName = x.ProcessName,
                ProcessType = x.ProcessType,
                TotalPpm    = x.Total,
                Ppm         = BuildPeriodDict(dateCols, weekCols, monthCols,
                    d => dateNgMap.GetValueOrDefault((x.ProcessType, x.ProcessName, x.NgName, d)),
                    w => weekNgMap.GetValueOrDefault((x.ProcessType, x.ProcessName, x.NgName, w)),
                    m => monthNgMap.GetValueOrDefault((x.ProcessType, x.ProcessName, x.NgName, m))),
                Groups      = groupNames.Select(g => new GroupPivotRow
                {
                    GroupName = g,
                    Ppm       = BuildPeriodDict(dateCols, weekCols, monthCols,
                        d => { var v = dateGrpNg.GetValueOrDefault((x.ProcessType, x.ProcessName, x.NgName, g, d)); return CalcPpm(v.I, v.N); },
                        w => { var v = weekGrpNg.GetValueOrDefault((x.ProcessType, x.ProcessName, x.NgName, g, w)); return CalcPpm(v.I, v.N); },
                        m => { var v = monthGrpNg.GetValueOrDefault((x.ProcessType, x.ProcessName, x.NgName, g, m)); return CalcPpm(v.I, v.N); }),
                }).ToList(),
            };
        }).ToList();
    }

    // ── Pivot dictionary builder ──────────────────────────────────────────────────

    private static Dictionary<string, double> BuildPeriodDict(
        List<PeriodColumn> dateCols,
        List<PeriodColumn> weekCols,
        List<PeriodColumn> monthCols,
        Func<string, double> getDate,
        Func<string, double> getWeek,
        Func<string, double> getMonth)
    {
        var dict = new Dictionary<string, double>();
        foreach (var c in dateCols)  dict[c.Key] = getDate(c.Key);
        foreach (var c in weekCols)  dict[c.Key] = getWeek(c.Key);
        foreach (var c in monthCols) dict[c.Key] = getMonth(c.Key);
        return dict;
    }

    // ── Helpers ──────────────────────────────────────────────────────────────────

    private static double CalcPpm(double input, double ng)
    {
        if (input <= 0) return 0;
        return Math.Round((ng / input) * 1_000_000, 0);
    }

    private static int ProcessTypeOrder(string pt) => pt switch
    {
        "SUB"      => 1,
        "MAIN"     => 2,
        "FUNCTION" => 3,
        "VISUAL"   => 4,
        _          => pt.StartsWith("SUB") ? 1 : 5,
    };

    private static string GetWeekKey(string date)
    {
        if (!DateTime.TryParse(date, out var dt)) return string.Empty;
        int week = CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(
            dt, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
        return $"{dt.Year:0000}{week:00}";
    }

    private static string GetMonthKey(string date)
    {
        if (!DateTime.TryParse(date, out var dt)) return string.Empty;
        return $"{dt.Year:0000}{dt.Month:00}";
    }

    private static string FormatDateHeader(string date)
    {
        if (DateTime.TryParse(date, out var dt)) return dt.ToString("MM/dd");
        return date.Length >= 5 ? date[^5..] : date;
    }

    // ── Top-10 group raw map ──────────────────────────────────────────────────────

    // Expand each OrgRow into one entry per group it belongs to, so a LineShift that
    // sits in both a leaf group and its umbrella parent contributes to BOTH aggregates.
    private static IEnumerable<string> GroupsOf(
        Dictionary<string, List<string>> map, string lineShift)
    {
        if (map.TryGetValue(lineShift, out var list) && list.Count > 0) return list;
        return new[] { lineShift };
    }

    // Compute PPM per NgName first, then sum — matches the parent row formula
    private static Dictionary<(string, string, string, string), double> BuildGroupProcessRaw(
        List<OrgRow> rows, Dictionary<string, List<string>> lsToGroups, Func<OrgRow, string> getKey)
        => rows
            .SelectMany(r => GroupsOf(lsToGroups, r.LineShift).Select(g => (r, G: g)))
            .GroupBy(x => (x.r.ProcessType, x.r.ProcessName,
                           x.G,
                           K: getKey(x.r),
                           x.r.NgName))
            .Where(g => !string.IsNullOrEmpty(g.Key.K))
            .Select(g => (
                Key: (g.Key.ProcessType, g.Key.ProcessName, g.Key.G, g.Key.K),
                Ppm: CalcPpm(g.Sum(x => x.r.QtyInput), g.Sum(x => x.r.QtyNg))))
            .GroupBy(x => x.Key)
            .ToDictionary(g => g.Key, g => g.Sum(x => x.Ppm));

    private static Dictionary<(string, string, string, string, string), (double I, double N)> BuildGroupNgRaw(
        List<OrgRow> rows, Dictionary<string, List<string>> lsToGroups, Func<OrgRow, string> getKey)
        => rows
            .SelectMany(r => GroupsOf(lsToGroups, r.LineShift).Select(g => (r, G: g)))
            .GroupBy(x => (x.r.ProcessType, x.r.ProcessName, x.r.NgName,
                           x.G,
                           K: getKey(x.r)))
            .Where(g => !string.IsNullOrEmpty(g.Key.K))
            .ToDictionary(
                g => (g.Key.ProcessType, g.Key.ProcessName, g.Key.NgName, g.Key.G, g.Key.K),
                g => (I: g.Sum(x => x.r.QtyInput), N: g.Sum(x => x.r.QtyNg)));

    // ── Per-group summary (ProcessType × period PPM) ─────────────────────────────

    private static Dictionary<string, List<SummaryPivotRow>> ComputeGroupSummary(
        List<string> groupNames,
        Dictionary<(string, string, string, string), double> dateGrpProc,
        Dictionary<(string, string, string, string), double> weekGrpProc,
        Dictionary<(string, string, string, string), double> monthGrpProc,
        List<PeriodColumn> dateCols,
        List<PeriodColumn> weekCols,
        List<PeriodColumn> monthCols)
    {
        var result = new Dictionary<string, List<SummaryPivotRow>>(StringComparer.Ordinal);

        foreach (string grp in groupNames)
        {
            // (ProcessType, PeriodKey) → sumPpm — sum over all ProcessNames in this group
            var datePt = dateGrpProc
                .Where(kv => kv.Key.Item3 == grp)
                .GroupBy(kv => (kv.Key.Item1, kv.Key.Item4))
                .ToDictionary(g => g.Key, g => g.Sum(kv => kv.Value));

            var weekPt = weekGrpProc
                .Where(kv => kv.Key.Item3 == grp)
                .GroupBy(kv => (kv.Key.Item1, kv.Key.Item4))
                .ToDictionary(g => g.Key, g => g.Sum(kv => kv.Value));

            var monthPt = monthGrpProc
                .Where(kv => kv.Key.Item3 == grp)
                .GroupBy(kv => (kv.Key.Item1, kv.Key.Item4))
                .ToDictionary(g => g.Key, g => g.Sum(kv => kv.Value));

            var summary = new List<SummaryPivotRow>();

            var types = datePt .Keys.Select(k => k.Item1)
                .Concat(weekPt .Keys.Select(k => k.Item1))
                .Concat(monthPt.Keys.Select(k => k.Item1))
                .Where(t => !string.IsNullOrEmpty(t))
                .ToHashSet()
                .OrderBy(ProcessTypeOrder)
                .ThenBy(t => t);

            foreach (string pt in types)
            {
                summary.Add(new SummaryPivotRow
                {
                    ProcessType = pt,
                    IsTotal     = false,
                    Ppm         = BuildPeriodDict(dateCols, weekCols, monthCols,
                        d => datePt .GetValueOrDefault((pt, d)),
                        w => weekPt .GetValueOrDefault((pt, w)),
                        m => monthPt.GetValueOrDefault((pt, m))),
                });
            }

            // TOTAL = sum of PPM across ProcessType rows (insert at the front)
            var totalPpm = new Dictionary<string, double>();
            foreach (var row in summary)
                foreach (var (k, v) in row.Ppm)
                    totalPpm[k] = totalPpm.GetValueOrDefault(k) + v;

            summary.Insert(0, new SummaryPivotRow
            {
                ProcessType = "TOTAL",
                IsTotal     = true,
                Ppm         = totalPpm,
            });

            result[grp] = summary;
        }

        return result;
    }

    // ── Per-group input/NG quantities & Reason ───────────────────────────────────

    private static List<InputQtyRow> BuildInputQtyRows(
        List<OrgRow> rows,
        Dictionary<string, List<string>> lineShiftToGroups,
        List<PeriodColumn> dateCols)
    {
        if (dateCols.Count == 0) return [];
        var dateKeys = dateCols.Select(c => c.Key).ToHashSet(StringComparer.Ordinal);

        var aggDict = rows
            .Where(r => dateKeys.Contains(r.ProductDate))
            .SelectMany(r => GroupsOf(lineShiftToGroups, r.LineShift).Select(g => (r, Group: g)))
            .GroupBy(x => (x.r.ProcessType, x.r.ProcessName,
                           x.Group,
                           x.r.ProductDate))
            .ToDictionary(g => g.Key, g => (long)g.Sum(x => x.r.QtyInput));

        var procList = aggDict.Keys
            .Select(k => (k.ProcessType, k.ProcessName))
            .Distinct()
            .OrderBy(x => ProcessTypeOrder(x.ProcessType))
            .ThenBy(x => x.ProcessName)
            .ToList();

        var procNoMap = procList.Select((x, i) => (x, No: i + 1))
            .ToDictionary(x => x.x, x => x.No);

        var groupOrder = aggDict.Keys
            .Select(k => k.Group).Where(g => !string.IsNullOrEmpty(g))
            .Distinct().OrderBy(g => g).ToList();

        var result = new List<InputQtyRow>();
        foreach (var (pt, pn) in procList)
        {
            int no = procNoMap[(pt, pn)];
            foreach (string grp in groupOrder)
            {
                var qty = new Dictionary<string, long>();
                bool hasData = false;
                foreach (var col in dateCols)
                {
                    long v = aggDict.GetValueOrDefault((pt, pn, grp, col.Key));
                    qty[col.Key] = v;
                    if (v > 0) hasData = true;
                }
                if (hasData)
                    result.Add(new InputQtyRow { No = no, ProcessType = pt, ProcessName = pn, GroupName = grp, Qty = qty });
            }
        }
        return result;
    }

    private static List<NgQtyRow> BuildNgQtyRows(
        List<OrgRow> rows,
        Dictionary<string, List<string>> lineShiftToGroups,
        List<PeriodColumn> dateCols)
    {
        if (dateCols.Count == 0) return [];
        var dateKeys = dateCols.Select(c => c.Key).ToHashSet(StringComparer.Ordinal);

        var aggDict = rows
            .Where(r => dateKeys.Contains(r.ProductDate) && r.QtyNg > 0)
            .SelectMany(r => GroupsOf(lineShiftToGroups, r.LineShift).Select(g => (r, Group: g)))
            .GroupBy(x => (x.r.ProcessType, x.r.ProcessName, x.r.NgName,
                           x.Group,
                           x.r.ProductDate))
            .ToDictionary(g => g.Key, g => (long)g.Sum(x => x.r.QtyNg));

        var comboList = aggDict.Keys
            .Select(k => (k.ProcessType, k.ProcessName, k.NgName))
            .Distinct()
            .OrderBy(x => ProcessTypeOrder(x.ProcessType))
            .ThenBy(x => x.ProcessName)
            .ThenBy(x => x.NgName)
            .ToList();

        var comboNoMap = comboList.Select((x, i) => (x, No: i + 1))
            .ToDictionary(x => x.x, x => x.No);

        var groupOrder = aggDict.Keys
            .Select(k => k.Group).Where(g => !string.IsNullOrEmpty(g))
            .Distinct().OrderBy(g => g).ToList();

        var result = new List<NgQtyRow>();
        foreach (var (pt, pn, ng) in comboList)
        {
            int no = comboNoMap[(pt, pn, ng)];
            foreach (string grp in groupOrder)
            {
                var qty = new Dictionary<string, long>();
                bool hasData = false;
                foreach (var col in dateCols)
                {
                    long v = aggDict.GetValueOrDefault((pt, pn, ng, grp, col.Key));
                    qty[col.Key] = v;
                    if (v > 0) hasData = true;
                }
                if (hasData)
                    result.Add(new NgQtyRow { No = no, ProcessType = pt, ProcessName = pn, NgName = ng, GroupName = grp, Qty = qty });
            }
        }
        return result;
    }

    private static List<ReasonRow> BuildReasonRows(
        List<OrgRow> rows,
        Dictionary<string, string> reasonLookup,
        Dictionary<string, List<string>> lineShiftToGroups,
        List<PeriodColumn> dateCols,
        List<PeriodColumn> weekCols,
        List<PeriodColumn> monthCols)
    {
        if (dateCols.Count == 0 || reasonLookup.Count == 0) return [];
        var dateKeys = dateCols.Select(c => c.Key).ToHashSet(StringComparer.Ordinal);

        var tagged = rows
            .Where(r => dateKeys.Contains(r.ProductDate))
            .SelectMany(r => GroupsOf(lineShiftToGroups, r.LineShift)
                .Select(g => (r, Reason: reasonLookup.GetValueOrDefault(r.ProcessName + "\0" + r.NgName, ""),
                                 Group:  g)))
            .Where(x => !string.IsNullOrEmpty(x.Reason))
            .ToList();

        // (Reason, ProcessType, ProcessName, NgName, PeriodKey) → (SumInput, SumNg)
        var aggDate = tagged
            .GroupBy(x => (x.Reason, x.r.ProcessType, x.r.ProcessName, x.r.NgName, x.r.ProductDate))
            .ToDictionary(g => g.Key, g => (Input: g.Sum(x => x.r.QtyInput), Ng: g.Sum(x => x.r.QtyNg)));

        var aggWeek = tagged
            .GroupBy(x => (x.Reason, x.r.ProcessType, x.r.ProcessName, x.r.NgName, GetWeekKey(x.r.ProductDate)))
            .ToDictionary(g => g.Key, g => (Input: g.Sum(x => x.r.QtyInput), Ng: g.Sum(x => x.r.QtyNg)));

        var aggMonth = tagged
            .GroupBy(x => (x.Reason, x.r.ProcessType, x.r.ProcessName, x.r.NgName, GetMonthKey(x.r.ProductDate)))
            .ToDictionary(g => g.Key, g => (Input: g.Sum(x => x.r.QtyInput), Ng: g.Sum(x => x.r.QtyNg)));

        // (Reason, ProcessType, ProcessName, NgName, Group, PeriodKey) → (SumInput, SumNg)
        var grpDate = tagged
            .GroupBy(x => (x.Reason, x.r.ProcessType, x.r.ProcessName, x.r.NgName, x.Group, x.r.ProductDate))
            .ToDictionary(g => g.Key, g => (Input: g.Sum(x => x.r.QtyInput), Ng: g.Sum(x => x.r.QtyNg)));

        var grpWeek = tagged
            .GroupBy(x => (x.Reason, x.r.ProcessType, x.r.ProcessName, x.r.NgName, x.Group, GetWeekKey(x.r.ProductDate)))
            .ToDictionary(g => g.Key, g => (Input: g.Sum(x => x.r.QtyInput), Ng: g.Sum(x => x.r.QtyNg)));

        var grpMonth = tagged
            .GroupBy(x => (x.Reason, x.r.ProcessType, x.r.ProcessName, x.r.NgName, x.Group, GetMonthKey(x.r.ProductDate)))
            .ToDictionary(g => g.Key, g => (Input: g.Sum(x => x.r.QtyInput), Ng: g.Sum(x => x.r.QtyNg)));

        static double ToPpm(double input, double ng)
            => input > 0 ? Math.Round((ng / input) * 1_000_000, 0) : 0;

        // Sort Reasons desc by prior-week PPM sum (weekCols[1])
        var prevWeekKeyReason = weekCols.Count >= 2 ? weekCols[1].Key : weekCols.FirstOrDefault()?.Key ?? string.Empty;

        var reasonGroups = aggDate.Keys
            .GroupBy(k => k.Reason)
            .Select(rg => (
                Reason: rg.Key,
                TotalPpm: rg
                    .Select(k => (k.ProcessType, k.ProcessName, k.NgName))
                    .Distinct()
                    .Sum(c => ToPpm(
                        aggWeek.GetValueOrDefault((rg.Key, c.ProcessType, c.ProcessName, c.NgName, prevWeekKeyReason)).Input,
                        aggWeek.GetValueOrDefault((rg.Key, c.ProcessType, c.ProcessName, c.NgName, prevWeekKeyReason)).Ng))))
            .OrderByDescending(x => x.TotalPpm)
            .ToList();

        var result = new List<ReasonRow>();
        foreach (var (reason, _) in reasonGroups)
        {
            // Combos (ProcessType, ProcessName, NgName) for this Reason — desc by prior-week PPM
            var combos = aggDate.Keys
                .Where(k => k.Reason == reason)
                .GroupBy(k => (k.ProcessType, k.ProcessName, k.NgName))
                .Select(g => (
                    g.Key.ProcessType, g.Key.ProcessName, g.Key.NgName,
                    TotalPpm: ToPpm(
                        aggWeek.GetValueOrDefault((reason, g.Key.ProcessType, g.Key.ProcessName, g.Key.NgName, prevWeekKeyReason)).Input,
                        aggWeek.GetValueOrDefault((reason, g.Key.ProcessType, g.Key.ProcessName, g.Key.NgName, prevWeekKeyReason)).Ng)))
                .OrderByDescending(x => x.TotalPpm)
                .Take(10)
                .ToList();

            // Total row
            result.Add(new ReasonRow
            {
                Reason  = reason,
                IsTotal = true,
                Ppm     = BuildPeriodDict(dateCols, weekCols, monthCols,
                    d => combos.Sum(c => ToPpm(aggDate.GetValueOrDefault((reason, c.ProcessType, c.ProcessName, c.NgName, d)).Input,
                                              aggDate.GetValueOrDefault((reason, c.ProcessType, c.ProcessName, c.NgName, d)).Ng)),
                    w => combos.Sum(c => ToPpm(aggWeek.GetValueOrDefault((reason, c.ProcessType, c.ProcessName, c.NgName, w)).Input,
                                              aggWeek.GetValueOrDefault((reason, c.ProcessType, c.ProcessName, c.NgName, w)).Ng)),
                    m => combos.Sum(c => ToPpm(aggMonth.GetValueOrDefault((reason, c.ProcessType, c.ProcessName, c.NgName, m)).Input,
                                              aggMonth.GetValueOrDefault((reason, c.ProcessType, c.ProcessName, c.NgName, m)).Ng))),
            });

            // Individual rows
            for (int i = 0; i < combos.Count; i++)
            {
                var (pt, pn, ng, _) = combos[i];

                var groupNames = grpDate.Keys
                    .Where(k => k.Item1 == reason && k.Item2 == pt && k.Item3 == pn && k.Item4 == ng)
                    .Select(k => k.Item5)
                    .Concat(grpWeek.Keys.Where(k => k.Item1 == reason && k.Item2 == pt && k.Item3 == pn && k.Item4 == ng).Select(k => k.Item5))
                    .Concat(grpMonth.Keys.Where(k => k.Item1 == reason && k.Item2 == pt && k.Item3 == pn && k.Item4 == ng).Select(k => k.Item5))
                    .Distinct().OrderBy(g => g).ToList();

                var groups = groupNames.Select(grp => new GroupPivotRow
                {
                    GroupName = grp,
                    Ppm       = BuildPeriodDict(dateCols, weekCols, monthCols,
                        d => ToPpm(grpDate.GetValueOrDefault((reason, pt, pn, ng, grp, d)).Input,  grpDate.GetValueOrDefault((reason, pt, pn, ng, grp, d)).Ng),
                        w => ToPpm(grpWeek.GetValueOrDefault((reason, pt, pn, ng, grp, w)).Input,  grpWeek.GetValueOrDefault((reason, pt, pn, ng, grp, w)).Ng),
                        m => ToPpm(grpMonth.GetValueOrDefault((reason, pt, pn, ng, grp, m)).Input, grpMonth.GetValueOrDefault((reason, pt, pn, ng, grp, m)).Ng)),
                }).ToList();

                result.Add(new ReasonRow
                {
                    Reason      = reason,
                    No          = i + 1,
                    ProcessType = pt,
                    ProcessName = pn,
                    NgName      = ng,
                    Ppm         = BuildPeriodDict(dateCols, weekCols, monthCols,
                        d => ToPpm(aggDate.GetValueOrDefault((reason, pt, pn, ng, d)).Input,  aggDate.GetValueOrDefault((reason, pt, pn, ng, d)).Ng),
                        w => ToPpm(aggWeek.GetValueOrDefault((reason, pt, pn, ng, w)).Input,  aggWeek.GetValueOrDefault((reason, pt, pn, ng, w)).Ng),
                        m => ToPpm(aggMonth.GetValueOrDefault((reason, pt, pn, ng, m)).Input, aggMonth.GetValueOrDefault((reason, pt, pn, ng, m)).Ng)),
                    Groups      = groups,
                });
            }
        }
        return result;
    }

    // ── SQLite ────────────────────────────────────────────────────────────────────

    private static Dictionary<string, string> LoadReasonLookup(string dbPath)
    {
        var lookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        using var conn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
        conn.Open();
        if (!TableExists(conn, "Reason")) return lookup;
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT [processName], [NgName], [Reason] FROM [Reason]";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            string pn     = ReadStr(r, 0);
            string ngName = ReadStr(r, 1);
            string reason = ReadStr(r, 2);
            if (!string.IsNullOrEmpty(pn) && !string.IsNullOrEmpty(ngName))
                lookup.TryAdd(pn + "\0" + ngName, reason);
        }
        return lookup;
    }

    private static List<OrgRow> LoadOrginalRows(string dbPath, HashSet<string>? modelFilter)
    {
        var rows = new List<OrgRow>();
        using var conn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
        conn.Open();
        if (!TableExists(conn, "OrginalTable")) return rows;

        bool hasLineShift = ColumnExists(conn, "OrginalTable", "LineShift");
        string lineShiftExpr = hasLineShift
            ? "[LineShift]"
            : "([MATERIALNAME] || '_' || [PRODUCTION_LINE])";

        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"""
            SELECT [MATERIALNAME], {lineShiftExpr}, [PROCESSCODE], [PROCESSNAME],
                   [NGNAME], [QTYINPUT], [QTYNG], [PRODUCT_DATE]
            FROM [OrginalTable]
            """;
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            string lineShift = ReadStr(r, 1);
            if (modelFilter is not null && !modelFilter.Contains(lineShift)) continue;
            rows.Add(new OrgRow
            {
                MaterialName = ReadStr(r, 0),
                LineShift    = lineShift,
                ProcessCode  = ReadStr(r, 2),
                ProcessName  = NormalizeText(ReadStr(r, 3)),
                NgName       = NormalizeText(ReadStr(r, 4)),
                QtyInput     = ReadDouble(r, 5),
                QtyNg        = ReadDouble(r, 6),
                ProductDate  = ReadStr(r, 7),
            });
        }
        return rows;
    }

    private static Dictionary<(string, string, string), string> LoadProcessTypeLookup(string dbPath)
    {
        var lookup = new Dictionary<(string, string, string), string>();
        using var conn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
        conn.Open();
        if (!TableExists(conn, "RoutingTable")) return lookup;
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT [모델명], [ProcessCode], [ProcessName], [ProcessType] FROM [RoutingTable]";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            string mat  = ReadStr(r, 0), code = ReadStr(r, 1),
                   name = NormalizeText(ReadStr(r, 2)), pt = ReadStr(r, 3);
            lookup.TryAdd((mat, code, name), pt);
            lookup.TryAdd((string.Empty, code, name), pt);
        }
        return lookup;
    }

    private static bool TableExists(SqliteConnection conn, string t)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@t";
        cmd.Parameters.AddWithValue("@t", t);
        return (long)cmd.ExecuteScalar()! > 0;
    }

    private static bool ColumnExists(SqliteConnection conn, string table, string col)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"PRAGMA table_info([{table}])";
        using var r = cmd.ExecuteReader();
        while (r.Read())
            if (r.GetString(1).Equals(col, StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    private static string NormalizeText(string input)
    {
        if (string.IsNullOrWhiteSpace(input)) return string.Empty;
        input = input.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
        input = input.Replace("\u2018", "'").Replace("\u2019", "'")
                     .Replace("\u201c", "\"").Replace("\u201d", "\"");
        input = input.Replace("'", " ").Replace("\"", " ").Replace("~", " ");
        input = input.Replace("[", "").Replace("]", "_").Replace("+", " ");
        input = Regex.Replace(input, @"\s{2,}", " ");
        return input.Trim();
    }

    private static string ReadStr(SqliteDataReader r, int i)
        => r.IsDBNull(i) ? string.Empty : r.GetValue(i).ToString()!;

    private static double ReadDouble(SqliteDataReader r, int i)
    {
        if (r.IsDBNull(i)) return 0;
        var val = r.GetValue(i);
        if (val is double d) return d;
        if (val is long   l) return l;
        double.TryParse(val?.ToString(), NumberStyles.Any,
                        CultureInfo.InvariantCulture, out double res);
        return res;
    }
}
