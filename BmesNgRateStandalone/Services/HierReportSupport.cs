namespace BmesNgRateStandalone.Services;

/// <summary>Bundle of the 5 NgRate reports needed by a Group → Mid → Sub1 → Sub2
/// → LineShift hierarchy view. Pages either build all 5 in one shot via
/// <see cref="HierReportBuilder.BuildAsync"/>, or wrap their existing report
/// references via <see cref="HierPpmLookup"/> when the same data is already
/// computed for other tables.</summary>
public sealed class HierReports
{
    public NgRateReportService.NgRateReport? ByGroup { get; init; }   // LS → Group name
    public NgRateReportService.NgRateReport? ByMid   { get; init; }   // LS → "Group::Material"
    public NgRateReportService.NgRateReport? BySub1  { get; init; }   // LS → top-level sub key (subtree aggregate)
    public NgRateReportService.NgRateReport? BySub2  { get; init; }   // LS → leaf-path key (this node's own LS only)
    public NgRateReportService.NgRateReport? ByLs    { get; init; }   // LS → LS (per-LineShift bucket)
    public List<ModelGroupRecord>            Groups { get; init; } = new();
    public IReadOnlyDictionary<string, string> MidMapping { get; init; }
        = new Dictionary<string, string>(StringComparer.Ordinal);
    public IReadOnlyList<string> LineShifts { get; init; } = Array.Empty<string>();

    public bool HasData => ByGroup is not null && Groups.Count > 0;
}

/// <summary>PPM lookups across the hierarchy levels. Reads the IsTotal row of the
/// appropriate <see cref="NgRateReportService.NgRateReport"/> for any column key.
/// All lookups return 0 when the underlying report or key is missing — caller
/// can use that as the "no data" signal.</summary>
public sealed class HierPpmLookup
{
    public NgRateReportService.NgRateReport? ByGroup { get; init; }
    public NgRateReportService.NgRateReport? ByMid   { get; init; }
    public NgRateReportService.NgRateReport? BySub1  { get; init; }
    public NgRateReportService.NgRateReport? BySub2  { get; init; }
    public NgRateReportService.NgRateReport? ByLs    { get; init; }

    public static HierPpmLookup From(HierReports r) => new()
    {
        ByGroup = r.ByGroup, ByMid = r.ByMid,
        BySub1 = r.BySub1, BySub2 = r.BySub2, ByLs = r.ByLs,
    };

    public double Group(string groupName, string colKey)
        => Read(ByGroup, groupName, colKey);

    public double Mid(string groupName, string material, string colKey)
        => Read(ByMid, $"{groupName}::{material}", colKey);

    /// <summary>Mid PPM rendered as the simple average of its Sub1 PPMs (skipping
    /// zeros) so the combo-sum aggregation can't balloon Mid above any of its
    /// children. Falls back to <see cref="Mid"/> when no named Sub1 has data.</summary>
    public double MidAvgOfSubs(string groupName, MidGroupRecord mid, string colKey)
    {
        double sum = 0; int cnt = 0;
        foreach (var sub in mid.SubGroups)
        {
            if (string.IsNullOrEmpty(sub.Name)) continue;
            if (sub.AllLineShifts.Count() == 0) continue;
            string sub1Key = ModelGroupPickerHelpers.SubGroupKeyOf(groupName, mid.Material, sub);
            double v = Sub1(sub1Key, colKey);
            if (v <= 0) continue;
            sum += v; cnt++;
        }
        return cnt == 0 ? Mid(groupName, mid.Material, colKey) : sum / cnt;
    }

    /// <summary>Group PPM rendered as the simple average of its Mid PPMs (skipping
    /// zeros) — same anti-combo-sum policy as <see cref="MidAvgOfSubs"/> applied one
    /// level up so Group stays bounded by its children's values. Each Mid PPM is
    /// itself a Sub1-average, so the entire Group → Mid → Sub1 chain is consistent.
    /// Falls back to <see cref="Group"/> when no Mid in the group has data.</summary>
    public double GroupAvgOfMids(ModelGroupRecord group, string colKey)
    {
        double sum = 0; int cnt = 0;
        foreach (var mid in group.MidGroups)
        {
            if (string.IsNullOrEmpty(mid.Material)) continue;
            if (mid.LineShifts.Count == 0) continue;
            double v = MidAvgOfSubs(group.Name, mid, colKey);
            if (v <= 0) continue;
            sum += v; cnt++;
        }
        return cnt == 0 ? Group(group.Name, colKey) : sum / cnt;
    }

    public double Sub1(string sub1Key, string colKey) => Read(BySub1, sub1Key,  colKey);
    public double Sub2(string leafKey, string colKey) => Read(BySub2, leafKey,  colKey);
    public double Ls  (string lineShift, string colKey) => Read(ByLs, lineShift, colKey);

    private static double Read(NgRateReportService.NgRateReport? report, string key, string colKey)
    {
        if (report is null) return 0;
        var rows = report.GroupSummary.GetValueOrDefault(key);
        return rows?.FirstOrDefault(r => r.IsTotal)?.Ppm.GetValueOrDefault(colKey) ?? 0;
    }
}

/// <summary>One row in the Sub-level hierarchy walker. Pages render with depth-based
/// indent + parent-expansion gating. <see cref="IsLineShift"/> distinguishes virtual
/// LineShift leaves from actual SubGroup nodes for styling/expansion (LineShifts
/// have <see cref="HasChildren"/>=false).</summary>
public sealed record HierSubRow(
    string  RowId,
    string? ParentId,
    string  Display,
    int     Depth,
    bool    HasChildren,
    bool    IsLineShift,
    Func<string, double> PpmFor);

/// <summary>Builds the 5 hierarchy reports + the per-Mid sub-row list. Centralises
/// the mapping construction (LineShift → Group / Mid / Sub1 / Sub2 / LineShift) and
/// the recursive walker that emits Sub1/Sub2/… rows plus virtual LineShift leaves.</summary>
public static class HierReportBuilder
{
    private sealed class PrefixedProgress(IProgress<string> inner, string label) : IProgress<string>
    {
        public void Report(string value)
        {
            if (string.Equals(value, "Report complete.", StringComparison.Ordinal))
            {
                inner.Report($"{label} complete.");
                return;
            }

            inner.Report($"[{label}] {value}");
        }
    }

    private static IProgress<string>? WithProgressLabel(IProgress<string>? progress, string label)
        => progress is null ? null : new PrefixedProgress(progress, label);

    /// <summary>Generate all 5 hierarchy reports in one pass. Empty <paramref name="selectedGroups"/>
    /// or no LineShifts → returns an empty <see cref="HierReports"/> (HasData=false).</summary>
    public static async Task<HierReports> BuildAsync(
        NgRateReportService             svc,
        string                          dbPath,
        IReadOnlyList<ModelGroupRecord> selectedGroups,
        IProgress<string>?              progress = null,
        DateTime?                       periodStart = null,
        DateTime?                       periodEnd = null,
        bool                            includeMidDetailReport = false)
    {
        if (selectedGroups.Count == 0)
            return new HierReports();

        var mapping        = new Dictionary<string, string>(StringComparer.Ordinal);
        var midMapping     = new Dictionary<string, string>(StringComparer.Ordinal);
        var subMapping     = new Dictionary<string, string>(StringComparer.Ordinal);
        var subLeafMapping = new Dictionary<string, string>(StringComparer.Ordinal);
        var lsMapping      = new Dictionary<string, string>(StringComparer.Ordinal);
        var groupList   = new List<string>();
        var midList     = new List<string>();
        var subList     = new List<string>();
        var subLeafList = new List<string>();
        var lsList      = new List<string>();

        foreach (var g in selectedGroups)
        {
            if (!groupList.Contains(g.Name)) groupList.Add(g.Name);
            foreach (var mid in g.MidGroups)
            {
                string matKey = $"{g.Name}::{mid.Material}";
                if (!midList.Contains(matKey)) midList.Add(matKey);

                foreach (var sub in mid.SubGroups)
                {
                    string subKey = ModelGroupPickerHelpers.SubGroupKeyOf(g.Name, mid.Material, sub);
                    if (!subList.Contains(subKey)) subList.Add(subKey);
                    foreach (var ls in sub.AllLineShifts)
                    {
                        if (string.IsNullOrEmpty(ls)) continue;
                        subMapping[ls] = subKey;
                    }
                    ModelGroupPickerHelpers.BuildSubLeafMapping(
                        g.Name, mid.Material, sub, parentPath: string.Empty,
                        subLeafMapping, subLeafList);
                }

                foreach (var ls in mid.LineShifts)
                {
                    if (string.IsNullOrEmpty(ls)) continue;
                    mapping[ls]    = g.Name;
                    midMapping[ls] = matKey;
                    lsMapping[ls]  = ls;
                    if (!lsList.Contains(ls)) lsList.Add(ls);
                }
            }
        }

        if (mapping.Count == 0) return new HierReports();

        progress?.Report("─── Building Group/Mid/Sub hierarchy reports");
        static void AddMembership(Dictionary<string, List<string>> target, string lineShift, string groupName)
        {
            if (string.IsNullOrEmpty(lineShift) || string.IsNullOrEmpty(groupName)) return;
            if (!target.TryGetValue(lineShift, out var groups))
            {
                groups = new List<string>();
                target[lineShift] = groups;
            }
            if (!groups.Contains(groupName, StringComparer.Ordinal)) groups.Add(groupName);
        }

        var allMapping = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        foreach (var kv in mapping)        AddMembership(allMapping, kv.Key, kv.Value);
        foreach (var kv in midMapping)     AddMembership(allMapping, kv.Key, kv.Value);
        foreach (var kv in subMapping)     AddMembership(allMapping, kv.Key, kv.Value);
        foreach (var kv in subLeafMapping) AddMembership(allMapping, kv.Key, kv.Value);
        foreach (var kv in lsMapping)      AddMembership(allMapping, kv.Key, kv.Value);

        var allGroupNames = groupList
            .Concat(midList)
            .Concat(subList)
            .Concat(subLeafList)
            .Concat(lsList)
            .Distinct(StringComparer.Ordinal)
            .ToList();
        var allMappingReadOnly = allMapping.ToDictionary(
            kv => kv.Key,
            kv => (IReadOnlyList<string>)kv.Value,
            StringComparer.Ordinal);

        progress?.Report(includeMidDetailReport
            ? "Starting hierarchy summary, Mid detail, and Sub Group detail reports in parallel."
            : "Starting hierarchy summary report.");
        var summaryProgress = WithProgressLabel(progress, "Hierarchy summary");
        var midProgress = WithProgressLabel(progress, "Mid detail");
        var sub1Progress = WithProgressLabel(progress, "Sub Group detail");
        var summaryTask = svc.GenerateSummaryReportAsync(
            dbPath, allMappingReadOnly, allGroupNames, summaryProgress, periodStart, periodEnd,
            weightedGroupSummary: true);
        var byMidTask = includeMidDetailReport
            ? svc.GenerateReportAsync(dbPath, midMapping, midList, midProgress, periodStart, periodEnd)
            : null;
        var bySub1Task = includeMidDetailReport && subList.Count > 0
            ? svc.GenerateReportAsync(dbPath, subMapping, subList, sub1Progress, periodStart, periodEnd)
            : null;

        var summary = await summaryTask;
        var byMid = byMidTask is not null ? await byMidTask : summary;
        var bySub1 = bySub1Task is not null
            ? await bySub1Task
            : (subList.Count > 0 ? summary : null);

        return new HierReports
        {
            ByGroup = summary,
            ByMid   = byMid,
            BySub1  = bySub1,
            BySub2  = subLeafList.Count > 0 ? summary : null,
            ByLs    = lsList.Count > 0 ? summary : null,
            Groups  = selectedGroups.ToList(),
            MidMapping = new Dictionary<string, string>(midMapping, StringComparer.Ordinal),
            LineShifts = lsList.ToList(),
        };
    }

    /// <summary>Walk <paramref name="mid"/>'s SubGroup tree and emit Sub1 / Sub2 / …
    /// rows plus virtual LineShift leaves for any sub holding direct LineShifts.
    /// Top-level subs (depth=1) read PPM from <c>BySub1</c>; nested leaves use
    /// <c>BySub2</c>; LineShift rows pull from <c>ByLs</c>. Empty-named placeholder
    /// subs collapse — their named children render at the parent's depth.</summary>
    public static List<HierSubRow> BuildSubRows(
        string groupName, MidGroupRecord mid, HierPpmLookup lookup)
    {
        var rows = new List<HierSubRow>();
        string midRowId = $"{groupName}::{mid.Material}";
        foreach (var sub in mid.SubGroups)
            EmitSub(groupName, mid.Material, sub, parentPath: string.Empty,
                    depth: 1, parentRowId: midRowId, lookup, rows);
        return rows;
    }

    private static void EmitSub(
        string groupName, string material, SubGroupRecord sub,
        string parentPath, int depth, string parentRowId,
        HierPpmLookup lookup, List<HierSubRow> rows)
    {
        string nodeName = sub.Name ?? string.Empty;
        string path = string.IsNullOrEmpty(nodeName)
            ? parentPath
            : (string.IsNullOrEmpty(parentPath) ? nodeName : parentPath + "::" + nodeName);

        string thisRowId = parentRowId;
        bool emit = !string.IsNullOrEmpty(nodeName) && sub.AllLineShifts.Count() > 0;
        if (emit)
        {
            string ownRowId;
            Func<string, double> ppmFor;
            if (depth == 1)
            {
                string sub1Key = ModelGroupPickerHelpers.SubGroupKeyOf(groupName, material, sub);
                ownRowId = "S1::" + sub1Key;
                ppmFor   = colKey => lookup.Sub1(sub1Key, colKey);
            }
            else
            {
                string leafKey = ModelGroupPickerHelpers.SubLeafKeyOf(groupName, material, path);
                ownRowId = "S" + depth + "::" + leafKey;
                ppmFor   = colKey => lookup.Sub2(leafKey, colKey);
            }
            bool hasNamedKids = sub.SubGroups.Any(c =>
                !string.IsNullOrEmpty(c.Name) && c.AllLineShifts.Count() > 0);
            bool hasDirectLs  = sub.LineShifts.Count > 0;
            bool hasChildren  = hasNamedKids || hasDirectLs;
            rows.Add(new HierSubRow(ownRowId, parentRowId, nodeName, depth,
                                    hasChildren, IsLineShift: false, ppmFor));
            thisRowId = ownRowId;

            if (hasDirectLs)
            {
                int lsDepth = depth + 1;
                foreach (var ls in sub.LineShifts)
                {
                    if (string.IsNullOrEmpty(ls)) continue;
                    string lsCaptured = ls;
                    rows.Add(new HierSubRow(
                        RowId: "LS::" + ownRowId + "::" + lsCaptured,
                        ParentId: ownRowId,
                        Display: lsCaptured,
                        Depth: lsDepth,
                        HasChildren: false,
                        IsLineShift: true,
                        PpmFor: colKey => lookup.Ls(lsCaptured, colKey)));
                }
            }
        }

        int childDepth = string.IsNullOrEmpty(nodeName) ? depth : depth + 1;
        foreach (var child in sub.SubGroups)
            EmitSub(groupName, material, child, path, childDepth, thisRowId, lookup, rows);
    }

    /// <summary>True iff every Sub-level ancestor in <paramref name="row"/>'s chain
    /// is in <paramref name="expanded"/>. The HashSet is keyed by RowId — same
    /// shape <see cref="BuildSubRows"/> emits.</summary>
    public static bool IsRowVisible(
        HierSubRow                              row,
        IReadOnlyDictionary<string, HierSubRow> byId,
        IReadOnlySet<string>                    expanded)
    {
        string? pid = row.ParentId;
        while (pid is not null && byId.TryGetValue(pid, out var p))
        {
            if (!expanded.Contains(pid)) return false;
            pid = p.ParentId;
        }
        return true;
    }
}
