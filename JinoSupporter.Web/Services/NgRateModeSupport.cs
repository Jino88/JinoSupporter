namespace JinoSupporter.Web.Services;

public sealed record NgRateFlatMapping(
    Dictionary<string, string> Mapping,
    List<string> Groups);

public sealed record NgRateTreeMapping(
    Dictionary<string, IReadOnlyList<string>> Mapping,
    List<string> Groups);

public sealed record NgRateChildTreeReport(
    string Label,
    Dictionary<string, IReadOnlyList<string>> Mapping,
    List<string> Groups);

public sealed record NgRateChildFlatReport(
    string Label,
    Dictionary<string, string> Mapping,
    List<string> Groups);

public sealed record NgRateGroupMappings(
    Dictionary<string, string> GroupMapping,
    Dictionary<string, string> MidMapping,
    Dictionary<string, string> Sub1Mapping,
    Dictionary<string, string> Sub2Mapping,
    Dictionary<string, string> LineShiftMapping,
    List<string> GroupList,
    List<string> MidList,
    List<string> Sub1List,
    List<string> Sub2List,
    List<string> LineShiftList)
{
    public bool HasData => GroupMapping.Count > 0;
}

public sealed record NgRateGroupReportBundle(
    NgRateReportService.NgRateReport GroupReport,
    NgRateReportService.NgRateReport MidReport,
    NgRateReportService.NgRateReport LineShiftReport,
    NgRateReportService.NgRateReport? Sub1Report,
    NgRateReportService.NgRateReport? Sub2Report,
    NgRateReportService.LineShiftNgReport LineShiftNgReport,
    NgRateReportService.LineShiftNgReport MidNgReport,
    NgRateGroupMappings Mappings)
{
    public HierPpmLookup ToLookup() => new()
    {
        ByGroup = GroupReport,
        ByMid = MidReport,
        BySub1 = Sub1Report,
        BySub2 = Sub2Report,
        ByLs = LineShiftReport,
    };
}

public static class NgRateModeSupport
{
    public const int DefaultDateCap = 10;
    public const int DefaultWeekCap = 4;
    public const int DefaultMonthCap = 3;

    public static string ModelKey(string groupName, string material)
        => $"{groupName}::{material}";

    public static string FormatPpm(double value)
        => value <= 0 ? "-" : ((long)Math.Round(value)).ToString("N0");

    public static string PpmClass(double value)
        => value <= 0 ? "ppm-zero" : "ppm-val";

    public static int ProcessTypeRank(string processType) => processType switch
    {
        "SUB" => 1,
        "MAIN" => 2,
        "FUNCTION" => 3,
        "VISUAL" => 4,
        _ => processType.StartsWith("SUB", StringComparison.Ordinal) ? 1 : 5,
    };

    public static List<ModelGroupRecord> SelectGroupsByName(
        IEnumerable<ModelGroupRecord> groups,
        IReadOnlySet<string> selectedGroupNames)
        => groups
            .Where(g => selectedGroupNames.Contains(g.Name))
            .ToList();

    public static NgRateFlatMapping BuildFlatMapping(IEnumerable<ModelGroupRecord> selectedGroups)
    {
        var mapping = new Dictionary<string, string>(StringComparer.Ordinal);
        var groups = new List<string>();

        foreach (var group in selectedGroups)
        {
            AddDistinct(groups, group.Name);
            foreach (var mid in group.MidGroups)
            {
                foreach (var lineShift in mid.LineShifts)
                {
                    if (!string.IsNullOrEmpty(lineShift))
                        mapping[lineShift] = group.Name;
                }
            }
        }

        return new NgRateFlatMapping(mapping, groups);
    }

    public static NgRateTreeMapping BuildTreeMultiMapping(ModelGroupRecord group)
    {
        var raw = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var groups = new List<string>();

        void Attach(string lineShift, string groupName)
        {
            if (string.IsNullOrEmpty(lineShift) || string.IsNullOrEmpty(groupName)) return;

            if (!raw.TryGetValue(lineShift, out var list))
            {
                list = new List<string>();
                raw[lineShift] = list;
            }

            AddDistinct(list, groupName);
            AddDistinct(groups, groupName);
        }

        var allLineShifts = group.MidGroups
            .SelectMany(m => m.LineShifts)
            .Where(ls => !string.IsNullOrEmpty(ls))
            .Distinct(StringComparer.Ordinal)
            .ToList();
        if (allLineShifts.Count == 0)
        {
            return new NgRateTreeMapping(
                new Dictionary<string, IReadOnlyList<string>>(StringComparer.Ordinal),
                groups);
        }

        foreach (var lineShift in allLineShifts)
            Attach(lineShift, group.Name);

        bool includeMids = group.MidGroups.Count > 1
            || (group.MidGroups.Count == 1
                && !string.Equals(group.MidGroups[0].Material, group.Name, StringComparison.Ordinal));
        if (includeMids)
        {
            foreach (var mid in group.MidGroups)
            {
                if (string.IsNullOrEmpty(mid.Material)) continue;
                foreach (var lineShift in mid.LineShifts)
                    Attach(lineShift, mid.Material);
            }
        }

        foreach (var mid in group.MidGroups)
        {
            foreach (var sub in mid.SubGroups)
                AttachSubGroupTree(sub, Attach);
        }

        foreach (var mid in group.MidGroups)
        {
            foreach (var sub in mid.SubGroups)
                AttachLineShiftLeaves(sub, Attach);
        }

        return new NgRateTreeMapping(
            raw.ToDictionary(
                kv => kv.Key,
                kv => (IReadOnlyList<string>)kv.Value,
                StringComparer.Ordinal),
            groups);
    }

    public static NgRateChildTreeReport? BuildChildTreeReport(
        IReadOnlyList<ModelGroupRecord> allGroups,
        string key)
    {
        var parts = key.Split('|');
        if (parts.Length < 3) return null;
        if (!long.TryParse(parts[0], out long groupId)) return null;

        var group = allGroups.FirstOrDefault(g => g.Id == groupId);
        if (group is null) return null;

        if (parts[1] == "mid" && parts.Length == 3)
        {
            string material = parts[2];
            var mid = group.MidGroups.FirstOrDefault(m => m.Material == material);
            if (mid is null) return null;

            var virtualGroup = new ModelGroupRecord
            {
                Id = group.Id,
                Name = material,
                ProductGroup = group.ProductGroup,
                MidGroups = new List<MidGroupRecord> { mid },
            };
            var tree = BuildTreeMultiMapping(virtualGroup);
            return new NgRateChildTreeReport(material, tree.Mapping, tree.Groups);
        }

        if (parts[1] == "sub" && parts.Length == 4)
        {
            string material = parts[2];
            string path = parts[3];
            var mid = group.MidGroups.FirstOrDefault(m => m.Material == material);
            if (mid is null) return null;

            var node = ModelGroupPickerHelpers.ResolveSubByPath(mid.SubGroups, path);
            if (node is null) return null;

            var soloMid = new MidGroupRecord
            {
                Material = mid.Material,
                SubGroups = new List<SubGroupRecord> { node },
            };
            string nodeName = node.Name ?? string.Empty;
            var virtualGroup = new ModelGroupRecord
            {
                Id = group.Id,
                Name = nodeName,
                ProductGroup = group.ProductGroup,
                MidGroups = new List<MidGroupRecord> { soloMid },
            };
            var tree = BuildTreeMultiMapping(virtualGroup);
            return new NgRateChildTreeReport(nodeName, tree.Mapping, tree.Groups);
        }

        return null;
    }

    public static NgRateChildFlatReport? BuildMidChildFlatReport(
        IReadOnlyList<ModelGroupRecord> allGroups,
        string key)
    {
        var parts = key.Split('|');
        if (parts.Length != 3 || parts[1] != "mid") return null;
        if (!long.TryParse(parts[0], out long groupId)) return null;

        var group = allGroups.FirstOrDefault(g => g.Id == groupId);
        if (group is null) return null;

        string material = parts[2];
        var mid = group.MidGroups.FirstOrDefault(m => m.Material == material);
        if (mid is null) return null;

        var mapping = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var lineShift in mid.LineShifts)
        {
            if (!string.IsNullOrEmpty(lineShift))
                mapping[lineShift] = material;
        }

        return mapping.Count == 0
            ? null
            : new NgRateChildFlatReport(material, mapping, new List<string> { material });
    }

    public static NgRateGroupMappings BuildGroupMappings(IEnumerable<ModelGroupRecord> selectedGroups)
    {
        var groupMapping = new Dictionary<string, string>(StringComparer.Ordinal);
        var midMapping = new Dictionary<string, string>(StringComparer.Ordinal);
        var sub1Mapping = new Dictionary<string, string>(StringComparer.Ordinal);
        var sub2Mapping = new Dictionary<string, string>(StringComparer.Ordinal);
        var lineShiftMapping = new Dictionary<string, string>(StringComparer.Ordinal);

        var groupList = new List<string>();
        var midList = new List<string>();
        var sub1List = new List<string>();
        var sub2List = new List<string>();
        var lineShiftList = new List<string>();

        foreach (var group in selectedGroups)
        {
            AddDistinct(groupList, group.Name);
            foreach (var mid in group.MidGroups)
            {
                string midKey = ModelKey(group.Name, mid.Material);
                AddDistinct(midList, midKey);

                foreach (var sub in mid.SubGroups)
                {
                    string sub1Key = ModelGroupPickerHelpers.SubGroupKeyOf(group.Name, mid.Material, sub);
                    AddDistinct(sub1List, sub1Key);

                    foreach (var lineShift in sub.AllLineShifts)
                    {
                        if (!string.IsNullOrEmpty(lineShift))
                            sub1Mapping[lineShift] = sub1Key;
                    }

                    ModelGroupPickerHelpers.BuildSubLeafMapping(
                        group.Name, mid.Material, sub, parentPath: string.Empty,
                        sub2Mapping, sub2List);
                }

                foreach (var lineShift in mid.LineShifts)
                {
                    if (string.IsNullOrEmpty(lineShift)) continue;

                    groupMapping[lineShift] = group.Name;
                    midMapping[lineShift] = midKey;
                    lineShiftMapping[lineShift] = lineShift;
                    AddDistinct(lineShiftList, lineShift);
                }
            }
        }

        return new NgRateGroupMappings(
            groupMapping,
            midMapping,
            sub1Mapping,
            sub2Mapping,
            lineShiftMapping,
            groupList,
            midList,
            sub1List,
            sub2List,
            lineShiftList);
    }

    public static async Task<NgRateGroupReportBundle?> GenerateGroupReportBundleAsync(
        NgRateReportService svc,
        string dbPath,
        IReadOnlyList<ModelGroupRecord> selectedGroups,
        IProgress<string>? progress = null,
        DateTime? periodStart = null,
        DateTime? periodEnd = null)
    {
        var mappings = BuildGroupMappings(selectedGroups);
        if (!mappings.HasData) return null;

        var groupReport = await svc.GenerateReportAsync(
            dbPath, mappings.GroupMapping, mappings.GroupList,
            progress, periodStart, periodEnd);
        var midReport = await svc.GenerateSummaryReportAsync(
            dbPath, mappings.MidMapping, mappings.MidList,
            progress, periodStart, periodEnd);
        var lineShiftReport = await svc.GenerateSummaryReportAsync(
            dbPath, mappings.LineShiftMapping, mappings.LineShiftList,
            progress, periodStart, periodEnd);
        var sub1Report = mappings.Sub1List.Count > 0
            ? await svc.GenerateSummaryReportAsync(
                dbPath, mappings.Sub1Mapping, mappings.Sub1List,
                progress, periodStart, periodEnd)
            : null;
        var sub2Report = mappings.Sub2List.Count > 0
            ? await svc.GenerateSummaryReportAsync(
                dbPath, mappings.Sub2Mapping, mappings.Sub2List,
                progress, periodStart, periodEnd)
            : null;
        var lineShiftNgReport = await svc.ComputeLineShiftNgDetailsAsync(
            dbPath, mappings.LineShiftList, progress, periodStart, periodEnd);
        var midNgReport = await svc.ComputeGroupedNgDetailsAsync(
            dbPath, mappings.MidMapping, progress, periodStart, periodEnd);

        return new NgRateGroupReportBundle(
            groupReport,
            midReport,
            lineShiftReport,
            sub1Report,
            sub2Report,
            lineShiftNgReport,
            midNgReport,
            mappings);
    }

    private static void AttachSubGroupTree(SubGroupRecord node, Action<string, string> attach)
    {
        string nodeName = node.Name ?? string.Empty;

        if (!string.IsNullOrEmpty(nodeName))
        {
            foreach (var lineShift in node.AllLineShifts)
                attach(lineShift, nodeName);
        }

        foreach (var child in node.SubGroups)
            AttachSubGroupTree(child, attach);
    }

    private static void AttachLineShiftLeaves(SubGroupRecord node, Action<string, string> attach)
    {
        foreach (var lineShift in node.LineShifts)
            attach(lineShift, lineShift);

        foreach (var child in node.SubGroups)
            AttachLineShiftLeaves(child, attach);
    }

    private static void AddDistinct(List<string> list, string value)
    {
        if (!string.IsNullOrEmpty(value) && !list.Contains(value, StringComparer.Ordinal))
            list.Add(value);
    }
}
