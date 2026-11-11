using System.Globalization;
using System.Text;
using System.Text.Json;

namespace PortalGNLAB.Web.Services;

/// <summary>An item the MES071050 lookup offers for a code or name: one BOM header (plant, alternative, usage).</summary>
public sealed record BmesBomHeader(string Plant, string Alternative, string Usage, string Code, string Name);

public static partial class BmesGateway
{
    /// <summary>Base quantity the screen asks for (Condition.EMENG); MENGE comes back per this many parents.</summary>
    public const int BomBaseQty = 1000;

    /// <summary>
    /// MES071050 (BOM screen) — forward BOM of one product, the same tree the screen shows.
    /// Input: logged-in session, plant and a product name or code. SearchLOOKUP finds the BOM header
    /// (only P- / S- codes, see <see cref="PickBomHeader"/>), then SearchList explodes it.
    /// Output: the BOM rows under the product (product itself left out), or null with the reason
    /// in <c>Error</c>.
    /// </summary>
    public static async Task<(List<BmesBomMaterialCandidate>? Rows, string Error)> FetchBomAsync(
        BmesWebSession session, string plant, string product, IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (await WebLoginErrorAsync(session, progress, cancellationToken) is { } loginError)
            return (null, loginError);

        progress?.Report("Calling MES071050/SearchLOOKUP…");
        BmesBomHeader? header = PickBomHeader(await LookupBomHeadersAsync(session, plant, product, cancellationToken), product);
        if (header is null)
            return (null, $"MES071050 has no P- or S- BOM for {product} in plant {plant}.");

        progress?.Report($"Calling MES071050/SearchList ({header.Code})…");
        string list = await PostJsonAsync(session, "/MES071050/SearchList", new
        {
            Condition = new
            {
                WERKS = header.Plant, STLAL = header.Alternative, MATNR = header.Code, MATNR_TX = "",
                DIREC = "F", STLAN = 1, EMENG = BomBaseQty,
            },
        }, cancellationToken);
        return (BomRowsFromList(list), "");
    }

    /// <summary>
    /// MES071050/SearchLOOKUP — the BOM headers whose code or name contains <paramref name="text"/> (the screen's
    /// item popup; BMES answers at most 1000). Session must be logged in (<see cref="FetchBomAsync"/>) or is
    /// logged in here.
    /// </summary>
    public static async Task<List<BmesBomHeader>> LookupBomHeadersAsync(
        BmesWebSession session, string plant, string text, CancellationToken cancellationToken = default)
    {
        if (await WebLoginErrorAsync(session, null, cancellationToken) is { } loginError)
            throw new InvalidOperationException(loginError);
        string json = await PostJsonAsync(session, "/MES071050/SearchLOOKUP",
            new { Condition = new { WERKS = plant, MATNR = text } }, cancellationToken);
        return BomHeadersFromLookup(json);
    }

    /// <summary>A BMES web session for the BOM screen with the Setting login; throws when it is not configured.</summary>
    public static async Task<BmesWebSession> OpenBomSessionAsync(NgRateSettingsService settings, CancellationToken cancellationToken = default)
    {
        if (!settings.IsCredentialsConfigured)
            throw new InvalidOperationException("The BMES login is not configured in Setting.");
        var opened = await BmesWebSession.OpenAsync(settings.LoginId, settings.Password, new BmesWebSessionOptions
        {
            Timeout = TimeSpan.FromSeconds(120),
            AutomaticDecompression = System.Net.DecompressionMethods.All,
        }, cancellationToken: cancellationToken);
        return opened.Session ?? throw new InvalidOperationException("BMES login failed: " + opened.Diagnostic);
    }

    /// <summary>
    /// The P- / S- items of a lookup as (code, name), one per code, items without a name left out; at most
    /// <paramref name="max"/>.
    /// </summary>
    public static List<BmesBomHeader> BomProducts(IEnumerable<BmesBomHeader> headers, int max = 50) => headers
        .Where(h => IsBomProductCode(h.Code) && h.Name.Length > 0)
        .DistinctBy(h => h.Code, StringComparer.OrdinalIgnoreCase)
        .Take(max)
        .ToList();

    /// <summary>P- (product) or S- (semi-product) SAP code: the only items a BOM is looked up for.</summary>
    public static bool IsBomProductCode(string code) =>
        code.StartsWith("P-", StringComparison.OrdinalIgnoreCase) || code.StartsWith("S-", StringComparison.OrdinalIgnoreCase);

    private static async Task<string> PostJsonAsync(BmesWebSession session, string path, object body, CancellationToken cancellationToken)
    {
        using var content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");
        using var resp = await session.Client.PostAsync(BmesWebSession.BaseUrl + path, content, cancellationToken);
        if (!resp.IsSuccessStatusCode)
            throw new InvalidOperationException($"{path}: HTTP {(int)resp.StatusCode} {resp.StatusCode}");
        return await resp.Content.ReadAsStringAsync(cancellationToken);
    }

    /// <summary>
    /// The <c>data.contents</c> rows of a screen answer; none for "[IVDT_M_018] No result found", and an
    /// error with the BMES message for any other result false.
    /// </summary>
    private static List<JsonElement> ScreenContents(JsonDocument doc)
    {
        JsonElement root = doc.RootElement;
        if (root.TryGetProperty("result", out var ok) && ok.ValueKind == JsonValueKind.False)
        {
            string message = ReadRowText(root, "message");
            if (message.Contains("IVDT_M_018", StringComparison.Ordinal)) return [];
            throw new InvalidOperationException(message.Length > 0 ? message : "BMES answered no result.");
        }
        return root.TryGetProperty("data", out var data) && data.TryGetProperty("contents", out var c) && c.ValueKind == JsonValueKind.Array
            ? c.EnumerateArray().ToList()
            : [];
    }

    /// <summary>The BOM headers of a SearchLOOKUP answer.</summary>
    public static List<BmesBomHeader> BomHeadersFromLookup(string json)
    {
        using var doc = JsonDocument.Parse(json);
        return ScreenContents(doc)
            .Select(r => new BmesBomHeader(ReadRowText(r, "WERKS").Trim(), ReadRowText(r, "STLAL").Trim(),
                ReadRowText(r, "STLAN").Trim(), ReadRowText(r, "MATNR").Trim(), ReadRowText(r, "MAKTX").Trim()))
            .ToList();
    }

    /// <summary>
    /// The header to explode: only P- (product) and S- (semi-product) codes count — the lookup also
    /// lists Z-, T-, ... items of the same name. Exact name, else exact code, else the only P-/S-
    /// item; P- before S-, then the first alternative. Null when nothing fits.
    /// </summary>
    public static BmesBomHeader? PickBomHeader(IEnumerable<BmesBomHeader> headers, string product)
    {
        string want = product.Trim();
        var sap = headers
            .Where(h => IsBomProductCode(h.Code))
            .OrderBy(h => h.Code.StartsWith("P-", StringComparison.OrdinalIgnoreCase) ? 0 : 1)
            .ThenBy(h => h.Alternative, StringComparer.Ordinal)
            .ToList();
        return sap.FirstOrDefault(h => string.Equals(h.Name, want, StringComparison.OrdinalIgnoreCase))
            ?? sap.FirstOrDefault(h => string.Equals(h.Code, want, StringComparison.OrdinalIgnoreCase))
            ?? (sap.Select(h => h.Code).Distinct(StringComparer.OrdinalIgnoreCase).Count() == 1 ? sap[0] : null);
    }

    /// <summary>
    /// The rows of a SearchList answer as BOM rows. The answer is the tree in order with a level per row
    /// (ZLEVE 1 = the product); each row's parent is the last row one level up. Levels are shifted so the
    /// product's direct materials are level 1, like the BOM cache. MENGE is kept as BMES gives it.
    /// </summary>
    public static List<BmesBomMaterialCandidate> BomRowsFromList(string json)
    {
        using var doc = JsonDocument.Parse(json);
        var items = ScreenContents(doc)
            .Select(r => (Level: int.TryParse(ReadRowText(r, "ZLEVE"), out int l) ? l : 0,
                Code: ReadRowText(r, "MATNR").Trim(), Name: ReadRowText(r, "MATNR_TX").Trim(),
                Qty: decimal.TryParse(ReadRowText(r, "MENGE"), NumberStyles.Number, CultureInfo.InvariantCulture, out decimal q) ? q : (decimal?)null,
                Unit: ReadRowText(r, "MEINS").Trim()))
            .Where(x => x.Level > 0 && x.Code.Length > 0)
            .ToList();
        if (items.Count == 0) return [];

        var root = items[0];
        var path = new List<(string Code, string Name)> { (root.Code, root.Name) };
        var rows = new List<BmesBomMaterialCandidate>();
        for (int i = 1; i < items.Count; i++)
        {
            var x = items[i];
            int level = x.Level - root.Level;
            if (level < 1 || level > path.Count) continue; // not under the product, or a level skipped
            path.RemoveRange(level, path.Count - level);
            var parent = path[^1];
            path.Add((x.Code, x.Name));
            rows.Add(new BmesBomMaterialCandidate
            {
                ProductCode = root.Code,
                ProductName = root.Name,
                ParentMaterialCode = parent.Code,
                ParentMaterialName = parent.Name,
                MaterialCode = x.Code,
                MaterialName = x.Name,
                UsageQty = x.Qty,
                UsageUnit = x.Unit,
                BomLevel = level,
                BomPath = string.Join(">", path.Select(p => p.Code)),
                BomPathText = string.Join(" > ", path.Select(p => p.Name.Length > 0 ? p.Name : p.Code)),
                HasChildren = i + 1 < items.Count && items[i + 1].Level > x.Level,
                SourceRows = 1,
                Source = "MES071050",
            });
        }
        return rows;
    }
}
