using System.Globalization;

namespace PortalGNLAB.Web.Services;

/// <summary>One basic row of the Main-line F-Cost source (MES050027 verification view): facility × plant × line ×
/// finished product × material × period, standard and actual input in VND. The amounts are SAP amount × 100,
/// the same unit dbo.E008 (FCAMT / INAMT × 100) gave the old SQL.</summary>
public sealed record BmesFcostRawSourceRow(
    string Facility,
    string Plant,
    string Line,
    string ProductCode,
    string ProductName,
    string MaterialCode,
    string MaterialName,
    string PeriodKey,
    decimal StandardVnd,
    decimal ActualVnd);

public static partial class BmesGateway
{
    /// <summary>AI_TOOL_00001 view V (MES050027 source rows). <paramref name="periodCode"/> W = 7 ISO weeks
    /// (ZPVTT "yyyy-ww"), M = 7 months ("yyyyMM"), both counted back from <paramref name="sdate"/>.</summary>
    public static Dictionary<string, object> FcostRawParameters(DateTime sdate, string periodCode, string facility) => new()
    {
        ["ZTYPE"] = "V",
        ["SDATE"] = sdate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
        ["ZGUBN"] = periodCode,
        ["FACCO"] = facility,
    };

    /// <summary>
    /// MES050027 — F-Cost source rows (standard / actual input per finished product × material × period)
    /// for the 7 weeks or months up to <paramref name="sdate"/>, from PLATFORM_API (AI_TOOL_00001, view V).
    /// An empty answer is an empty list; a failed call throws <see cref="BmesApiException"/>.
    /// Checked 2026-10-29: per line shift × material the period sums equal the old dbo.E008 SQL exactly
    /// (2026-08, 2026-09, W41, W42, GN 3200).
    /// </summary>
    public static async Task<List<BmesFcostRawSourceRow>> FetchFcostRawRowsAsync(
        BmesWebSession session,
        DateTime sdate,
        string periodCode,
        string facility,
        CancellationToken cancellationToken = default)
    {
        BmesApiResult result = await BmesApi.ExecuteAsync(
            RequireApiCredentials(session), "AI_TOOL_00001", FcostRawParameters(sdate, periodCode, facility),
            cancellationToken: cancellationToken);
        return FcostRawRowsFromResult(result);
    }

    /// <summary>Tool answer (view V) → rows. Result 2 carries the codes and numbers; result 1 is the same rows as
    /// display text (comma numbers) and is only read for the product / material names, joined by code.</summary>
    public static List<BmesFcostRawSourceRow> FcostRawRowsFromResult(BmesApiResult result)
    {
        BmesApiResultSet? data = BmesToolRows.FindSet(result, "MATNR", "ASSEM", "ZPVTT", "INAMT", "FCAMT");
        if (data is null) return [];

        var productNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var materialNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (BmesToolRows.FindSet(result, "Product Code", "Product Name", "Material Code", "Material Name") is { } names)
        {
            foreach (var row in BmesToolRows.ToRowDictionaries(names))
            {
                productNames.TryAdd(row["Product Code"].Trim(), row["Product Name"].Trim());
                materialNames.TryAdd(row["Material Code"].Trim(), row["Material Name"].Trim());
            }
        }

        var rows = new List<BmesFcostRawSourceRow>(data.Rows.Count);
        foreach (var row in BmesToolRows.ToRowDictionaries(data))
        {
            string product = row["MATNR"].Trim();
            string material = row["ASSEM"].Trim();
            rows.Add(new BmesFcostRawSourceRow(
                row.GetValueOrDefault("FACCO", "").Trim(),
                row.GetValueOrDefault("WERKS", "").Trim(),
                row.GetValueOrDefault("VERID", "").Trim(),
                product,
                productNames.GetValueOrDefault(product, string.Empty),
                material,
                materialNames.GetValueOrDefault(material, string.Empty),
                row["ZPVTT"].Trim(),
                ParseVnd(row["INAMT"]),
                ParseVnd(row["FCAMT"])));
        }
        return rows;
    }

    private static decimal ParseVnd(string text) =>
        decimal.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal value) ? value : 0m;
}
