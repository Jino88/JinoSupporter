using System.Globalization;

namespace PortalGNLAB.Web.Services;

public sealed record WorkingBmesSubGroup(string Path, List<string> LineShifts);

public sealed record WorkingBmesMaterial(
    string Group, string ProductGroup, string Material, List<string> Products, List<WorkingBmesSubGroup> SubGroups);

public sealed record WorkingModelCandidate(string ProductName, string ProductCode, string Source);

public sealed record WorkingBmesBomRow(
    int Level, string MaterialCode, string MaterialName, decimal? UsageQty, string UsageUnit,
    string ParentMaterialName, bool HasChildren);

public sealed record WorkingBmesBom(
    string Product, string Source, string FetchedAt, string WorkDate, List<WorkingBmesBomRow> Rows, string Error);

/// <summary>Working Model Info: the BMES model groups and a product's BOM (cache first, server on request).</summary>
public static class WorkingBmesService
{
    public const string Plant = "3200";
    public const int MaxDepth = 6;
    public const int MaxRows = 2000;
    private const string PathSeparator = " › ";

    /// <summary>Content-Type of a downloaded drawing, from its file extension (any case).</summary>
    public static string DrawingContentType(string? fileName) =>
        Path.GetExtension(fileName ?? "").ToLowerInvariant() switch
        {
            ".pdf" => "application/pdf",
            ".png" => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".tif" or ".tiff" => "image/tiff",
            _ => "application/octet-stream",
        };

    /// <summary>One entry per group material, in group SortOrder and then input order.</summary>
    public static List<WorkingBmesMaterial> BuildMaterials(IEnumerable<ModelGroupRecord> groups) =>
        groups.OrderBy(g => g.SortOrder)
            .SelectMany(g => g.MidGroups.Select(m => new WorkingBmesMaterial(
                g.Name, g.ProductGroup, (m.Material ?? "").Trim(),
                ProductsOf(m.LineShifts), SubGroupsOf(m.SubGroups, []))))
            .Where(m => m.Material.Length > 0)
            .ToList();

    /// <summary>Line-shift names without the part after the last '_', first-seen order, no duplicates.</summary>
    public static List<string> ProductsOf(IEnumerable<string> lineShifts)
    {
        var products = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string raw in lineShifts)
        {
            string name = (raw ?? "").Trim();
            int cut = name.LastIndexOf('_');
            if (cut > 0) name = name[..cut];
            if (name.Length > 0 && seen.Add(name)) products.Add(name);
        }
        return products;
    }

    private static List<WorkingBmesSubGroup> SubGroupsOf(IEnumerable<SubGroupRecord> nodes, List<string> parents)
    {
        var result = new List<WorkingBmesSubGroup>();
        foreach (SubGroupRecord node in nodes)
        {
            var chain = new List<string>(parents);
            if (!string.IsNullOrWhiteSpace(node.Name)) chain.Add(node.Name.Trim());
            if (node.LineShifts.Count > 0)
                result.Add(new WorkingBmesSubGroup(string.Join(PathSeparator, chain), node.LineShifts.ToList()));
            result.AddRange(SubGroupsOf(node.SubGroups, chain));
        }
        return result;
    }

    public const int MaxCandidates = 10;

    private static bool SameName(string? a, string? b) =>
        string.Equals((a ?? "").Trim(), (b ?? "").Trim(), StringComparison.OrdinalIgnoreCase);

    /// <summary>Names containing the text (any case), exact matches first, at most MaxCandidates.</summary>
    private static List<WorkingModelCandidate> Matching(string name, IEnumerable<(string Name, string Code)> items, string source) =>
        items.Where(i => !string.IsNullOrWhiteSpace(i.Name) && i.Name.Contains(name.Trim(), StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(i => SameName(i.Name, name))
            .Select(i => new WorkingModelCandidate(i.Name.Trim(), (i.Code ?? "").Trim(), source))
            .Take(MaxCandidates)
            .ToList();

    public static List<WorkingModelCandidate> CatalogCandidates(string name, IEnumerable<BmesBomModelCandidate> models) =>
        Matching(name, models.Select(m => (m.ProductName, m.ProductCode)), "catalog");

    /// <summary>Group materials and product names (line shifts without their last "_" part) containing the text.</summary>
    public static List<WorkingModelCandidate> GroupCandidates(string name, IEnumerable<ModelGroupRecord> groups)
    {
        var names = new List<(string, string)>();
        foreach (WorkingBmesMaterial m in BuildMaterials(groups))
        {
            names.Add((m.Material, ""));
            names.AddRange(m.Products.Select(p => (p, "")));
        }
        return Matching(name, names.DistinctBy(n => n.Item1.Trim(), StringComparer.OrdinalIgnoreCase), "group");
    }

    /// <summary>
    /// Joins the candidates of the three sources (bmes, then catalog, then group wins when a name repeats) and tells
    /// whether any of them is exactly the checked name (any case, outer spaces ignored).
    /// </summary>
    public static (bool Exact, List<WorkingModelCandidate> Candidates) MergeCandidates(
        string name, IEnumerable<WorkingModelCandidate> bmes, IEnumerable<WorkingModelCandidate> catalog, IEnumerable<WorkingModelCandidate> group)
    {
        var merged = new List<WorkingModelCandidate>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (WorkingModelCandidate c in bmes.Concat(catalog).Concat(group))
            if (!string.IsNullOrWhiteSpace(c.ProductName) && seen.Add(c.ProductName.Trim())) merged.Add(c);
        return (merged.Any(c => SameName(c.ProductName, name)), merged);
    }

    public static WorkingBmesBom Empty(string product, string error = "") => new(product, "", "", "", [], error);

    public static WorkingBmesBom ToBom(string product, string source, BmesBomCacheEntry entry) => new(
        product, source, entry.FetchedAt,
        entry.Query.WorkDate == DateTime.MinValue ? "" : entry.Query.WorkDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
        entry.Rows.Select(r => new WorkingBmesBomRow(
            r.BomLevel, r.MaterialCode, r.MaterialName, r.UsageQty, r.UsageUnit, r.ParentMaterialName, r.HasChildren)).ToList(),
        "");
}
