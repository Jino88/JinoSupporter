using System.Globalization;

namespace PortalGNLAB.Web.Services;

public sealed record WorkingBmesSubGroup(string Path, List<string> LineShifts);

public sealed record WorkingBmesMaterial(
    string Group, string ProductGroup, string Material, List<string> Products, List<WorkingBmesSubGroup> SubGroups);

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

    public static WorkingBmesBom Empty(string product, string error = "") => new(product, "", "", "", [], error);

    public static WorkingBmesBom ToBom(string product, string source, BmesBomCacheEntry entry) => new(
        product, source, entry.FetchedAt,
        entry.Query.WorkDate == DateTime.MinValue ? "" : entry.Query.WorkDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
        entry.Rows.Select(r => new WorkingBmesBomRow(
            r.BomLevel, r.MaterialCode, r.MaterialName, r.UsageQty, r.UsageUnit, r.ParentMaterialName, r.HasChildren)).ToList(),
        "");
}
