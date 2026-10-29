using System.Text;
using System.Text.Json;

namespace PortalGNLAB.Web.Services;

public static partial class BmesGateway
{
    private const string LpaListEndpoint   = "/MES073260/SearchList";
    private const string LpaDetailEndpoint = "/MES073261/SearchList";

    /// <summary>
    /// MES073260/SearchList — LPA audit list (JSON body).
    /// Input: logged-in session and the search conditions. Output: generic rows (every property
    /// of the returned objects becomes a column, ordered by first appearance) or a failed result
    /// carrying the error text.
    /// </summary>
    public static Task<BmesLpaScrapeService.LpaResult> FetchLpaListAsync(
        BmesWebSession session, BmesLpaScrapeService.LpaQuery query, IProgress<string>? progress = null)
        => PostLpaSearchAsync(session, LpaListEndpoint, LpaListCondition(query), progress);

    /// <summary>
    /// MES073261/SearchList — LPA checklist behind one LQRNO (JSON body).
    /// Input: logged-in session and the detail conditions. Output: same generic rows as
    /// <see cref="FetchLpaListAsync"/>. Safe to call concurrently on one session.
    /// </summary>
    public static Task<BmesLpaScrapeService.LpaResult> FetchLpaDetailAsync(
        BmesWebSession session, BmesLpaScrapeService.LpaDetailQuery query, IProgress<string>? progress = null)
        => PostLpaSearchAsync(session, LpaDetailEndpoint, LpaDetailCondition(query), progress);

    /// <summary>
    /// MES073261/GetImage — one LPA result photo. Anonymous: no session, any HttpClient works.
    /// Input: the photo path (ZIMAG_TX). Output: the original bytes, or null when the request
    /// failed, the status was not success, or the answer was not an image (a deleted photo comes
    /// back as an HTML page with 200).
    /// </summary>
    public static async Task<byte[]?> FetchLpaImageAsync(HttpClient http, string path)
    {
        try
        {
            using var response = await http.GetAsync(LpaImageUrlPrefix + Uri.EscapeDataString(path));
            if (!response.IsSuccessStatusCode) return null;
            string? type = response.Content.Headers.ContentType?.MediaType;
            if (type is null || !type.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
                return null;
            return await response.Content.ReadAsByteArrayAsync();
        }
        catch { return null; }
    }

    /// <summary>Public (no login) image URL prefix; the viewer also uses it for the "Original" link.</summary>
    public const string LpaImageUrlPrefix = BmesWebSession.BaseUrl + "/MES073261/GetImage?fileName=";

    private static object LpaListCondition(BmesLpaScrapeService.LpaQuery query) => new
    {
        Condition = new
        {
            FACCO    = query.Facco,
            AUDAT_FR = query.DateFrom,
            AUDAT_TO = query.DateTo,
            LQRNO    = query.Lqrno,
            LQBNO    = query.Lqbno,
            AULOC    = query.Auloc,
            IMPLV    = query.Implv,
            DICNO    = query.Dicno,
            LASEQ    = query.Laseq,
            CHKER    = query.Chker,
            ZSTAT    = query.Zstat,
            USEYN    = query.UseYn,
        },
    };

    private static object LpaDetailCondition(BmesLpaScrapeService.LpaDetailQuery query) => new
    {
        Condition = new
        {
            LQRNO = query.Lqrno,
            FACCO = query.Facco,
            DICNO = query.Dicno,
            AULOC = query.Auloc,
            IMPLV = query.Implv,
            LCITM = query.Lcitm,
        },
    };

    /// <summary>JSON POST → generic row parse, on an already-authenticated session. Shared by
    /// both endpoints, which differ only in path and condition object.</summary>
    private static async Task<BmesLpaScrapeService.LpaResult> PostLpaSearchAsync(
        BmesWebSession session, string endpoint, object condition, IProgress<string>? progress)
    {
        string payload = JsonSerializer.Serialize(condition);

        progress?.Report($"Calling {endpoint}…");
        HttpResponseMessage response;
        try
        {
            using var content = new StringContent(payload, Encoding.UTF8, "application/json");
            response = await session.Client.PostAsync(BmesWebSession.BaseUrl + endpoint, content);
        }
        catch (Exception ex)
        {
            return BmesLpaScrapeService.LpaResult.Fail($"Request failed: {ex.Message}");
        }

        if (!response.IsSuccessStatusCode)
            return BmesLpaScrapeService.LpaResult.Fail($"HTTP {(int)response.StatusCode} {response.StatusCode}");

        string json = await response.Content.ReadAsStringAsync();
        try
        {
            var (columns, rows) = ParseLpaRows(json);
            progress?.Report($"✓ {rows.Count:N0} row(s), {columns.Count} column(s).");
            return new BmesLpaScrapeService.LpaResult(columns, rows, null);
        }
        catch (Exception ex)
        {
            return BmesLpaScrapeService.LpaResult.Fail($"Failed to parse the response: {ex.Message}");
        }
    }

    /// <summary>
    /// Pulls the row array out of the response and flattens it.
    ///
    /// The wrapper shape is not documented and differs between BMES endpoints
    /// (<c>data.contents</c>, <c>data</c>, a bare array…), so rather than guessing a path
    /// this walks the whole document and takes the largest array of objects it finds. On
    /// a result payload that is the row set by a wide margin — metadata arrays alongside
    /// it are short.
    /// </summary>
    private static (List<string> Columns, List<IReadOnlyDictionary<string, string>> Rows) ParseLpaRows(string json)
    {
        using var doc = JsonDocument.Parse(json);

        JsonElement? best = null;
        int bestCount = 0;
        FindLpaLargestObjectArray(doc.RootElement, 0, ref best, ref bestCount);

        var columns = new List<string>();
        var seen    = new HashSet<string>(StringComparer.Ordinal);
        var rows    = new List<IReadOnlyDictionary<string, string>>();
        if (best is not { } array) return (columns, rows);

        foreach (JsonElement item in array.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Object) continue;
            // Case-insensitive: callers look columns up by name and BMES casing varies.
            var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (JsonProperty prop in item.EnumerateObject())
            {
                if (seen.Add(prop.Name)) columns.Add(prop.Name);
                row[prop.Name] = StringifyLpa(prop.Value);
            }
            rows.Add(row);
        }

        return (columns, rows);
    }

    private const int LpaMaxScanDepth = 8;

    private static void FindLpaLargestObjectArray(
        JsonElement element, int depth, ref JsonElement? best, ref int bestCount)
    {
        if (depth > LpaMaxScanDepth) return;

        switch (element.ValueKind)
        {
            case JsonValueKind.Array:
                int objects = 0;
                foreach (JsonElement item in element.EnumerateArray())
                    if (item.ValueKind == JsonValueKind.Object) objects++;

                if (objects > bestCount)
                {
                    best = element;
                    bestCount = objects;
                }
                // Rows can still be nested one level deeper (e.g. grouped results).
                foreach (JsonElement item in element.EnumerateArray())
                    FindLpaLargestObjectArray(item, depth + 1, ref best, ref bestCount);
                break;

            case JsonValueKind.Object:
                foreach (JsonProperty prop in element.EnumerateObject())
                    FindLpaLargestObjectArray(prop.Value, depth + 1, ref best, ref bestCount);
                break;
        }
    }

    private static string StringifyLpa(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Null or JsonValueKind.Undefined => string.Empty,
        JsonValueKind.String => value.GetString() ?? string.Empty,
        JsonValueKind.True   => "Y",
        JsonValueKind.False  => "N",
        _ => value.ToString(),
    };
}
