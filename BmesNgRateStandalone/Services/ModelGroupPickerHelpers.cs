namespace BmesNgRateStandalone.Services;

/// <summary>One row in the Model-Groups picker. Pages render a checkbox per entry and
/// indent it by <see cref="Depth"/>. The <see cref="Key"/> is the stable identifier
/// the page tracks in its "selected" hashset.</summary>
public sealed record ModelGroupChildEntry(string Key, string Label, int LsCount, int Depth);

/// <summary>Static helpers shared by NG Rate / F-Cost group pickers. Builds a flattened
/// list of selectable rows (Mid + every named Sub) under each ModelGroup, and resolves
/// keys back to nodes when the user toggles them.</summary>
public static class ModelGroupPickerHelpers
{
    private static readonly string[] ProductGroupOrder = ["SPK", "UNIT", "MODULE", "TWS", "ETC"];

    /// <summary>Sub-group path separator - same unit-separator char WebRepository uses
    /// when persisting nested sub paths to <c>ModelGroupItems.SubGroup</c>.</summary>
    public const char SubPathSep = '';

    public static IEnumerable<string> ProductGroupsWithRows(IEnumerable<ModelGroupRecord> groups)
    {
        var groupList = groups as IReadOnlyCollection<ModelGroupRecord> ?? groups.ToList();
        return ProductGroupOrder.Where(pg => groupList.Any(g => g.ProductGroup == pg));
    }

    public static string FirstProductGroupWithRows(IEnumerable<ModelGroupRecord> groups, string fallback = "ETC")
        => ProductGroupsWithRows(groups).FirstOrDefault() ?? fallback;

    private static string BuildSubEntryLabel(string material, string path)
    {
        if (string.IsNullOrEmpty(material)) return path;
        if (string.IsNullOrEmpty(path)) return material;

        var parts = path.Split(SubPathSep);
        return parts.Length == 0 ? material : parts[^1];
    }

    /// <summary>Yield rows for the picker:
    ///   - Each Mid (depth=1) - checkbox includes the entire Mid subtree.
    ///   - Every named descendant Sub (depth=2,3,...) - checkbox includes that node's
    ///     full subtree. Container subs (no own LineShifts, only nested) ARE surfaced
    /// `_D3_L` becomes "Group::Material::_D3::L". Distinguishes the parent Sub1's
    /// Empty-named placeholder subs collapse - their children render at the parent's
    /// depth so a no-op container doesn't waste an indent level.
    /// A single Mid whose Material equals the Group name is suppressed as redundant.</summary>
    public static IEnumerable<ModelGroupChildEntry> EnumerateChildEntries(ModelGroupRecord g)
    {
        bool skipMidWhenSingle = g.MidGroups.Count == 1
            && string.Equals(g.MidGroups[0].Material, g.Name, StringComparison.Ordinal);

        foreach (var mid in g.MidGroups)
        {
            int midLs = mid.LineShifts.Count;
            if (midLs == 0) continue;
            if (string.IsNullOrEmpty(mid.Material)) continue;

            if (!skipMidWhenSingle)
            {
                yield return new ModelGroupChildEntry(
                    Key:     $"{g.Id}|mid|{mid.Material}",
                    Label:   mid.Material,
                    LsCount: midLs,
                    Depth:   1);
            }

            int subBaseDepth = skipMidWhenSingle ? 1 : 2;
            foreach (var sub in mid.SubGroups)
                foreach (var entry in WalkSubEntries(g.Id, mid, sub, parentPath: string.Empty, depth: subBaseDepth))
                    yield return entry;
        }
    }

    /// <summary>Yield only first-level Mid rows. Used by the By Section report,
    /// where each selected child becomes one report section.</summary>
    public static IEnumerable<ModelGroupChildEntry> EnumerateMidEntries(ModelGroupRecord g)
    {
        bool skipMidWhenSingle = g.MidGroups.Count == 1
            && string.Equals(g.MidGroups[0].Material, g.Name, StringComparison.Ordinal);

        foreach (var mid in g.MidGroups)
        {
            int midLs = mid.LineShifts.Count;
            if (midLs == 0) continue;
            if (skipMidWhenSingle) continue;
            if (string.IsNullOrEmpty(mid.Material)) continue;

            yield return new ModelGroupChildEntry(
                Key:     $"{g.Id}|mid|{mid.Material}",
                Label:   mid.Material,
                LsCount: midLs,
                Depth:   1);
        }
    }

    private static IEnumerable<ModelGroupChildEntry> WalkSubEntries(
        long gid, MidGroupRecord mid, SubGroupRecord sub, string parentPath, int depth)
    {
        string nodeName = sub.Name ?? string.Empty;
        string path = string.IsNullOrEmpty(nodeName)
            ? parentPath
            : (string.IsNullOrEmpty(parentPath) ? nodeName : parentPath + SubPathSep + nodeName);

        if (!string.IsNullOrEmpty(nodeName))
        {
            int subtreeLs = sub.AllLineShifts.Count();
            if (subtreeLs > 0)
            {
                yield return new ModelGroupChildEntry(
                    Key:     $"{gid}|sub|{mid.Material}|{path}",
                    Label:   BuildSubEntryLabel(mid.Material, path),
                    LsCount: subtreeLs,
                    Depth:   depth);
            }
        }

        int childDepth = string.IsNullOrEmpty(nodeName) ? depth : depth + 1;
        foreach (var child in sub.SubGroups)
            foreach (var e in WalkSubEntries(gid, mid, child, path, childDepth))
                yield return e;
    }

    /// <summary>Top-level SubGroup key - namespaced by Group + Material so different
    /// Materials with similarly-named subs don't collide. Empty Sub names get a
    /// bucket. Container Sub1s (no direct LineShifts, only nested children) are
    public static string SubGroupKeyOf(string groupName, string material, SubGroupRecord sub)
    {
        string subSegment = string.IsNullOrEmpty(sub.Name) ? "<default>" : sub.Name;
        return $"{groupName}::{material}::{subSegment}";
    }

    /// <summary>Leaf-path key - full sub path joined by "::" so a leaf like
    /// `_D3_L` becomes "Group::Material::_D3::L". Distinguishes the parent Sub1's
    /// own LineShifts from each Sub2 leaf inside it.</summary>
    public static string SubLeafKeyOf(string groupName, string material, string path)
        => $"{groupName}::{material}::{path}";

    /// <summary>Leaf-path key - full sub path joined by "::" so a leaf like
    /// mapping. Every named node that holds direct LineShifts becomes its own
    /// bucket. Container Sub1s (no direct LineShifts, only nested children) are
    /// bucket. Container Sub1s (no direct LineShifts, only nested children) are
    /// Empty-named placeholder subs collapse - their children render at the parent's
    /// a path segment.</summary>
    public static void BuildSubLeafMapping(
        string groupName, string material, SubGroupRecord sub,
        string parentPath,
        Dictionary<string, string> mapping, List<string> list)
    {
        string nodeName = sub.Name ?? string.Empty;
        string path = string.IsNullOrEmpty(nodeName)
            ? parentPath
            : (string.IsNullOrEmpty(parentPath) ? nodeName : parentPath + "::" + nodeName);

        if (!string.IsNullOrEmpty(nodeName) && sub.LineShifts.Count > 0)
        {
            string leafKey = SubLeafKeyOf(groupName, material, path);
            if (!list.Contains(leafKey)) list.Add(leafKey);
            foreach (var ls in sub.LineShifts)
                if (!string.IsNullOrEmpty(ls))
                    mapping[ls] = leafKey;
        }

        foreach (var child in sub.SubGroups)
            BuildSubLeafMapping(groupName, material, child, path, mapping, list);
    }

    /// <summary>Walk a sub-tree by SubPathSep-separated path. Returns null when any
    /// segment is missing.</summary>
    public static SubGroupRecord? ResolveSubByPath(List<SubGroupRecord> roots, string path)
    {
        if (string.IsNullOrEmpty(path)) return null;
        var segs = path.Split(SubPathSep);
        SubGroupRecord? cur = null;
        var list = roots;
        foreach (var seg in segs)
        {
            cur = list.FirstOrDefault(s => string.Equals(s.Name, seg, StringComparison.Ordinal));
            if (cur is null) return null;
            list = cur.SubGroups;
        }
        return cur;
    }

    /// <summary>Apply the user's checkbox selections to the full ModelGroup list,
    /// returning a filtered subset suitable for downstream rollups (F-Cost subgroup
    /// aggregation). Selections accumulate:
    ///   - A selected Group includes the entire Group as-is.
    ///   - A selected Mid (key "gid|mid|material") includes only that Mid (with all
    ///     its subs) inside an otherwise stripped-down copy of its Group.
    ///   - A selected Sub (key "gid|sub|material|path") includes only that sub-node
    ///     (with its own subtree) inside a stripped-down copy of its Mid + Group.
    /// When neither <paramref name="selectedGroupIds"/> nor <paramref name="selectedSubEntries"/>
    /// has anything, the original list is returned unchanged ("show all" default).</summary>
    public static List<ModelGroupRecord> ApplyPickerSelection(
        IReadOnlyList<ModelGroupRecord> allGroups,
        IReadOnlySet<long>              selectedGroupIds,
        IReadOnlySet<string>            selectedSubEntries)
    {
        if (selectedGroupIds.Count == 0 && selectedSubEntries.Count == 0)
            return allGroups.ToList();

        // Index sub-entry keys by groupId so we can iterate per-group efficiently.
        var subsByGid = new Dictionary<long, List<(string mat, string subType, string path)>>();
        foreach (string key in selectedSubEntries)
        {
            var parts = key.Split('|');
            if (parts.Length < 3) continue;
            if (!long.TryParse(parts[0], out long gid)) continue;
            string subType = parts[1];
            string mat = parts.Length > 2 ? parts[2] : string.Empty;
            string path = parts.Length > 3 ? parts[3] : string.Empty;
            if (!subsByGid.TryGetValue(gid, out var lst))
                subsByGid[gid] = lst = new();
            lst.Add((mat, subType, path));
        }

        static string BuildSelectionLabel(string material, string path)
        {
            if (string.IsNullOrEmpty(path)) return material;

            string[] segments = path.Split(SubPathSep);
            return segments.Length == 0 ? material : $"{material}_{string.Join("_", segments)}";
        }

        var result = new List<ModelGroupRecord>();
        foreach (var g in allGroups)
        {
            if (selectedGroupIds.Contains(g.Id))
            {
                result.Add(g);
                continue;
            }

            if (!subsByGid.TryGetValue(g.Id, out var subSelections)) continue;

            var wholeMidSet = subSelections
                .Where(x => string.Equals(x.subType, "mid", StringComparison.Ordinal))
                .Select(x => x.mat)
                .ToHashSet(StringComparer.Ordinal);

            var subEntries = subSelections
                .Where(x => string.Equals(x.subType, "sub", StringComparison.Ordinal))
                .ToList();

            if (wholeMidSet.Count == 0 && subEntries.Count == 0)
                continue;

            foreach (var mat in wholeMidSet)
            {
                var origMid = g.MidGroups.FirstOrDefault(m => m.Material == mat);
                if (origMid is null) continue;

                result.Add(new ModelGroupRecord
                {
                    Id = g.Id,
                    Name = mat,
                    ProductGroup = g.ProductGroup,
                    SortOrder = g.SortOrder,
                    MidGroups = new List<MidGroupRecord>
                    {
                        new MidGroupRecord
                        {
                            Material = mat,
                            SubGroups = origMid.SubGroups
                        }
                    },
                });
            }

            foreach (var (mat, _, path) in subEntries)
            {
                if (wholeMidSet.Contains(mat))
                    continue;

                var origMid = g.MidGroups.FirstOrDefault(m => m.Material == mat);
                if (origMid is null) continue;

                var node = ResolveSubByPath(origMid.SubGroups, path);
                if (node is null) continue;

                bool hasData = node.AllLineShifts.Any();
                if (!hasData) continue;

                result.Add(new ModelGroupRecord
                {
                    Id = g.Id,
                    Name = BuildSelectionLabel(mat, path),
                    ProductGroup = g.ProductGroup,
                    SortOrder = g.SortOrder,
                    MidGroups = new List<MidGroupRecord>
                    {
                        new MidGroupRecord
                        {
                            Material = mat,
                            SubGroups = new List<SubGroupRecord> { node }
                        }
                    },
                });
            }
        }

        return result;
    }
}
