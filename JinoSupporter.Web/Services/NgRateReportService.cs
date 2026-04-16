using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

/// <summary>
/// NG Rate 피벗 리포트 서비스 (WPF DataMaker 방식 동일)
/// ─ PPM 계산: 각 (ProcessName, NgName, Date) 조합별 PPM을 먼저 구한 후 합산
/// ─ TOTAL = 모든 ProcessType PPM 합산
/// ─ 열 구조: 날짜(내림차순) | 빈열 | 주차(내림차순) | 빈열 | 월(내림차순)
/// </summary>
public sealed class NgRateReportService(NgRateSettingsService settings)
{
    private readonly NgRateSettingsService _settings = settings;

    // ── 공개 타입 ─────────────────────────────────────────────────────────────────

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
    }

    // ── 내부 집계 단위 ────────────────────────────────────────────────────────────

    /// <summary>(ProcessType, ProcessName, NgName, PeriodKey) 단위의 PPM</summary>
    private sealed record ComboAgg(
        string ProcessType, string ProcessName, string NgName,
        string PeriodKey, double Ppm);

    // ── JSON 역직렬화용 ───────────────────────────────────────────────────────────

    private sealed class ModelGroupData
    {
        public int          Index     { get; set; }
        public string       GroupName { get; set; } = string.Empty;
        public List<string> ModelList { get; set; } = [];
    }

    /// <summary>새 포맷: { ProductGroup, Groups: [...] }</summary>
    private sealed class ModelFileData
    {
        public string              ProductGroup { get; set; } = "ETC";
        public List<ModelGroupData> Groups      { get; set; } = [];
    }

    // ── 내부 행 타입 ──────────────────────────────────────────────────────────────

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

    // ── 공개 API ──────────────────────────────────────────────────────────────────

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

    public (List<string> ModelList, List<string> GroupNames, Dictionary<string, string> LineShiftToGroup) LoadJsonInfo(IEnumerable<string> jsonPaths)
    {
        var models           = new HashSet<string>(StringComparer.Ordinal);
        var groupNames       = new List<string>();
        var lineShiftToGroup = new Dictionary<string, string>(StringComparer.Ordinal);
        var opts             = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };

        foreach (string path in jsonPaths)
        {
            if (!File.Exists(path)) continue;
            try
            {
                string json = File.ReadAllText(path);
                List<ModelGroupData>? groups = null;

                // 1) 새 포맷: { "ProductGroup": "...", "Groups": [...] }
                try
                {
                    var file = JsonSerializer.Deserialize<ModelFileData>(json, opts);
                    if (file?.Groups is { Count: > 0 }) groups = file.Groups;
                }
                catch { }

                // 2) 구 포맷: [ { "GroupName": ..., "ModelList": [...] }, ... ]
                if (groups is null || groups.Count == 0)
                {
                    try { groups = JsonSerializer.Deserialize<List<ModelGroupData>>(json, opts); } catch { }
                }

                // 3) 단일 객체 포맷 (레거시)
                if (groups is null || groups.Count == 0)
                {
                    try
                    {
                        var single = JsonSerializer.Deserialize<ModelGroupData>(json, opts);
                        if (single is not null) groups = [single];
                    }
                    catch { }
                }

                if (groups is null) continue;
                foreach (var g in groups)
                {
                    foreach (string m in g.ModelList)
                    {
                        if (!string.IsNullOrWhiteSpace(m))
                        {
                            string trimmed = m.Trim();
                            models.Add(trimmed);
                            if (!string.IsNullOrWhiteSpace(g.GroupName))
                                lineShiftToGroup[trimmed] = g.GroupName.Trim();
                        }
                    }
                    if (!string.IsNullOrWhiteSpace(g.GroupName))
                        groupNames.Add(g.GroupName.Trim());
                }
            }
            catch { }
        }
        return (models.ToList(), groupNames, lineShiftToGroup);
    }

    // 하위 호환용
    public List<string> LoadModelListFromJsonFiles(IEnumerable<string> jsonPaths)
        => LoadJsonInfo(jsonPaths).ModelList;

    public Task<NgRateReport> GenerateReportAsync(
        string dbPath,
        IEnumerable<string> selectedJsonPaths,
        IProgress<string>? progress = null)
        => Task.Run(() => GenerateReport(dbPath, selectedJsonPaths, progress));

    // ── 리포트 생성 ───────────────────────────────────────────────────────────────

    private NgRateReport GenerateReport(
        string dbPath,
        IEnumerable<string> selectedJsonPaths,
        IProgress<string>? progress)
    {
        // 1. JSON 파싱
        progress?.Report("Loading model definitions from JSON…");
        var jsonList = selectedJsonPaths.ToList();
        var (modelLineShifts, groupNames, lineShiftToGroup) = jsonList.Count > 0
            ? LoadJsonInfo(jsonList)
            : (new List<string>(), new List<string>(), new Dictionary<string, string>());

        var modelSet = modelLineShifts.Count > 0
            ? new HashSet<string>(modelLineShifts, StringComparer.Ordinal)
            : null;

        // 2. ProcessType lookup
        progress?.Report("Loading ProcessType lookup…");
        var ptLookup = LoadProcessTypeLookup(dbPath);

        // 3. OrginalTable 로드
        progress?.Report("Loading OrginalTable…");
        var orgRows = LoadOrginalRows(dbPath, modelSet);
        progress?.Report($"  {orgRows.Count:N0} rows loaded.");
        if (orgRows.Count == 0)
            return new NgRateReport { DbPath = dbPath, ModelFilter = modelLineShifts, GroupNames = groupNames };

        // 4. ProcessType 매핑
        foreach (var r in orgRows)
        {
            if (!string.IsNullOrEmpty(r.ProcessType)) continue;
            if (ptLookup.TryGetValue((r.MaterialName, r.ProcessCode, r.ProcessName), out string? pt))
                r.ProcessType = pt;
            else if (ptLookup.TryGetValue((string.Empty, r.ProcessCode, r.ProcessName), out string? pt2))
                r.ProcessType = pt2;
        }

        // 5. 기간별 조합 PPM 계산 (WPF 방식: 조합 단위 PPM 합산)
        progress?.Report("Pre-aggregating by period…");
        var byDate  = BuildComboPpm(orgRows, r => r.ProductDate);
        var byWeek  = BuildComboPpm(orgRows, r => GetWeekKey(r.ProductDate));
        var byMonth = BuildComboPpm(orgRows, r => GetMonthKey(r.ProductDate));

        // 6. 기간 컬럼 추출
        var dateCols  = ExtractCols(byDate,  k => FormatDateHeader(k));
        var weekCols  = ExtractCols(byWeek,  k => "W" + int.Parse(k[4..]));
        var monthCols = ExtractCols(byMonth, k => "M" + int.Parse(k[4..]));
        progress?.Report($"  {dateCols.Count} dates · {weekCols.Count} weeks · {monthCols.Count} months");

        // 7. 집계 Lookup 빌드
        // (ProcessType, PeriodKey) → sumPpm
        var datePtMap  = BuildPtMap(byDate);
        var weekPtMap  = BuildPtMap(byWeek);
        var monthPtMap = BuildPtMap(byMonth);

        // total per period
        var dateTotalMap  = BuildTotalMap(byDate);
        var weekTotalMap  = BuildTotalMap(byWeek);
        var monthTotalMap = BuildTotalMap(byMonth);

        // (ProcessType, ProcessName, PeriodKey) → sumPpm  ← Process 10 NG용
        var dateProcessMap  = BuildProcessMap(byDate);
        var weekProcessMap  = BuildProcessMap(byWeek);
        var monthProcessMap = BuildProcessMap(byMonth);

        // (ProcessType, ProcessName, NgName, PeriodKey) → Ppm  ← Worst 10 NG용
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
        var dateGrpProc  = BuildGroupProcessRaw(orgRows, lineShiftToGroup, r => r.ProductDate);
        var weekGrpProc  = BuildGroupProcessRaw(orgRows, lineShiftToGroup, r => GetWeekKey(r.ProductDate));
        var monthGrpProc = BuildGroupProcessRaw(orgRows, lineShiftToGroup, r => GetMonthKey(r.ProductDate));
        var top10Process = ComputeTop10Process(
            byDate, byWeek, byMonth,
            dateProcessMap, weekProcessMap, monthProcessMap,
            dateGrpProc, weekGrpProc, monthGrpProc,
            dateCols, weekCols, monthCols);

        // 10. Worst 10 NG
        progress?.Report("Computing Worst 10 NG Names…");
        var dateGrpNg  = BuildGroupNgRaw(orgRows, lineShiftToGroup, r => r.ProductDate);
        var weekGrpNg  = BuildGroupNgRaw(orgRows, lineShiftToGroup, r => GetWeekKey(r.ProductDate));
        var monthGrpNg = BuildGroupNgRaw(orgRows, lineShiftToGroup, r => GetMonthKey(r.ProductDate));
        var top10Ng = ComputeTop10Ng(
            byDate, byWeek, byMonth,
            dateNgMap, weekNgMap, monthNgMap,
            dateGrpNg, weekGrpNg, monthGrpNg,
            dateCols, weekCols, monthCols);

        // 11. 그룹별 투입/NG 수량 & Reason 리포트
        progress?.Report("Computing group detail rows…");
        var inputQtyRows = BuildInputQtyRows(orgRows, lineShiftToGroup, dateCols);
        var ngQtyRows    = BuildNgQtyRows(orgRows, lineShiftToGroup, dateCols);
        var reasonLookup = LoadReasonLookup(dbPath);
        var reasonRows   = BuildReasonRows(orgRows, reasonLookup, lineShiftToGroup, dateCols, weekCols, monthCols);

        // 12. 그룹별 Summary (모델명 클릭 시 각 그룹의 불량률)
        var groupSummary = groupNames.Count > 0
            ? ComputeGroupSummary(groupNames, dateGrpProc, weekGrpProc, monthGrpProc, dateCols, weekCols, monthCols)
            : new Dictionary<string, List<SummaryPivotRow>>();

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
        };
    }

    // ── 핵심 집계 빌더 ────────────────────────────────────────────────────────────

    /// <summary>
    /// (ProcessType, ProcessName, NgName, PeriodKey) 단위로 PPM 계산
    /// WPF와 동일: 각 조합의 PPM = QTYNG/QTYINPUT * 1,000,000
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

        // ProcessType별 (WPF 순서: SUB→MAIN→FUNCTION→VISUAL→기타)
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
        // 직전 주(weekCols[1]) 기준 내림차순 정렬, 주차 없으면 최신 주
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
        // 직전 주(weekCols[1]) 기준 내림차순 정렬, 주차 없으면 최신 주
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

    // ── Pivot 딕셔너리 빌더 ───────────────────────────────────────────────────────

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

    // ── 헬퍼 ─────────────────────────────────────────────────────────────────────

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

    // ── Top10 그룹 Raw 맵 ─────────────────────────────────────────────────────────

    // NgName별 PPM을 먼저 구한 뒤 합산 → 부모 행과 동일한 수식
    private static Dictionary<(string, string, string, string), double> BuildGroupProcessRaw(
        List<OrgRow> rows, Dictionary<string, string> lsToGroup, Func<OrgRow, string> getKey)
        => rows
            .GroupBy(r => (r.ProcessType, r.ProcessName,
                           G: lsToGroup.GetValueOrDefault(r.LineShift, r.LineShift),
                           K: getKey(r),
                           r.NgName))
            .Where(g => !string.IsNullOrEmpty(g.Key.K))
            .Select(g => (
                Key: (g.Key.ProcessType, g.Key.ProcessName, g.Key.G, g.Key.K),
                Ppm: CalcPpm(g.Sum(r => r.QtyInput), g.Sum(r => r.QtyNg))))
            .GroupBy(x => x.Key)
            .ToDictionary(g => g.Key, g => g.Sum(x => x.Ppm));

    private static Dictionary<(string, string, string, string, string), (double I, double N)> BuildGroupNgRaw(
        List<OrgRow> rows, Dictionary<string, string> lsToGroup, Func<OrgRow, string> getKey)
        => rows
            .GroupBy(r => (r.ProcessType, r.ProcessName, r.NgName,
                           G: lsToGroup.GetValueOrDefault(r.LineShift, r.LineShift),
                           K: getKey(r)))
            .Where(g => !string.IsNullOrEmpty(g.Key.K))
            .ToDictionary(
                g => (g.Key.ProcessType, g.Key.ProcessName, g.Key.NgName, g.Key.G, g.Key.K),
                g => (I: g.Sum(r => r.QtyInput), N: g.Sum(r => r.QtyNg)));

    // ── 그룹별 Summary (ProcessType × 기간 PPM) ──────────────────────────────────

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
            // (ProcessType, PeriodKey) → sumPpm — 해당 그룹의 모든 ProcessName 합산
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

            // TOTAL = ProcessType 행들의 PPM 합산 (순서상 맨 앞에 삽입)
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

    // ── 그룹별 투입/NG 수량 & Reason ─────────────────────────────────────────────

    private static List<InputQtyRow> BuildInputQtyRows(
        List<OrgRow> rows,
        Dictionary<string, string> lineShiftToGroup,
        List<PeriodColumn> dateCols)
    {
        if (dateCols.Count == 0) return [];
        var dateKeys = dateCols.Select(c => c.Key).ToHashSet(StringComparer.Ordinal);

        var aggDict = rows
            .Where(r => dateKeys.Contains(r.ProductDate))
            .GroupBy(r => (r.ProcessType, r.ProcessName,
                           Group: lineShiftToGroup.GetValueOrDefault(r.LineShift, r.LineShift),
                           r.ProductDate))
            .ToDictionary(g => g.Key, g => (long)g.Sum(r => r.QtyInput));

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
        Dictionary<string, string> lineShiftToGroup,
        List<PeriodColumn> dateCols)
    {
        if (dateCols.Count == 0) return [];
        var dateKeys = dateCols.Select(c => c.Key).ToHashSet(StringComparer.Ordinal);

        var aggDict = rows
            .Where(r => dateKeys.Contains(r.ProductDate) && r.QtyNg > 0)
            .GroupBy(r => (r.ProcessType, r.ProcessName, r.NgName,
                           Group: lineShiftToGroup.GetValueOrDefault(r.LineShift, r.LineShift),
                           r.ProductDate))
            .ToDictionary(g => g.Key, g => (long)g.Sum(r => r.QtyNg));

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
        Dictionary<string, string> lineShiftToGroup,
        List<PeriodColumn> dateCols,
        List<PeriodColumn> weekCols,
        List<PeriodColumn> monthCols)
    {
        if (dateCols.Count == 0 || reasonLookup.Count == 0) return [];
        var dateKeys = dateCols.Select(c => c.Key).ToHashSet(StringComparer.Ordinal);

        var tagged = rows
            .Where(r => dateKeys.Contains(r.ProductDate))
            .Select(r => (r, Reason: reasonLookup.GetValueOrDefault(r.ProcessName + "\0" + r.NgName, ""),
                            Group:  lineShiftToGroup.GetValueOrDefault(r.LineShift, r.LineShift)))
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

        // 직전 주(weekCols[1]) PPM 합산으로 Reason 정렬 (높은 순)
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
            // 해당 Reason의 (ProcessType, ProcessName, NgName) 조합 — 직전 주 PPM 내림차순
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

            // Total 행
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

            // 개별 행
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
