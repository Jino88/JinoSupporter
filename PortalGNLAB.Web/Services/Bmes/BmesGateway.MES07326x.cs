namespace PortalGNLAB.Web.Services;

public static partial class BmesGateway
{
    private const string LpaTool = "AI_TOOL_00027";

    /// <summary>
    /// MES073260/SearchList — LPA audit list, from PLATFORM_API (AI_TOOL_00027, view C: one row per audit with
    /// OK / NG / N/A counts). Ranges longer than 31 days are fetched in 31-day windows, one call each.
    /// Output: generic rows (every column of the Tool answer, in order) or a failed result carrying the error text.
    /// </summary>
    public static async Task<BmesLpaScrapeService.LpaResult> FetchLpaListAsync(
        BmesWebSession session, BmesLpaScrapeService.LpaQuery query, IProgress<string>? progress = null)
    {
        progress?.Report("Calling AI_TOOL_00027 (LPA audit list)…");
        try
        {
            BmesApiCredentials credentials = RequireApiCredentials(session);
            DateTime from = DateTime.Parse(BmesToolRows.ToIsoDate(query.DateFrom));
            DateTime to = DateTime.Parse(BmesToolRows.ToIsoDate(query.DateTo));

            var columns = new List<string>();
            var rows = new List<IReadOnlyDictionary<string, string>>();
            foreach (var (windowStart, windowEnd) in BmesToolRows.SplitPeriod(from, to))
            {
                BmesApiResult result = await BmesApi.ExecuteAsync(credentials, LpaTool, LpaListParameters(query, windowStart, windowEnd));
                var (windowColumns, windowRows) = LpaRowsFromResult(result, "LQRNO", "AUDAT");
                foreach (string column in windowColumns)
                    if (!columns.Contains(column)) columns.Add(column);
                rows.AddRange(windowRows);
            }

            rows = FilterLpaUseYn(rows, query.UseYn);
            progress?.Report($"✓ {rows.Count:N0} row(s), {columns.Count} column(s).");
            return new BmesLpaScrapeService.LpaResult(columns, rows, null);
        }
        catch (Exception ex)
        {
            return BmesLpaScrapeService.LpaResult.Fail($"BMES API request failed: {ex.Message}");
        }
    }

    /// <summary>
    /// MES073261/SearchList — LPA checklist behind one LQRNO, from PLATFORM_API (AI_TOOL_00027, view D).
    /// Output: same generic rows as <see cref="FetchLpaListAsync"/> (ZIMAG_TX still lists the photo paths;
    /// the photos themselves come from <see cref="FetchLpaImageAsync"/>).
    /// </summary>
    public static async Task<BmesLpaScrapeService.LpaResult> FetchLpaDetailAsync(
        BmesWebSession session, BmesLpaScrapeService.LpaDetailQuery query, IProgress<string>? progress = null)
    {
        progress?.Report("Calling AI_TOOL_00027 (LPA checklist)…");
        try
        {
            BmesApiResult result = await BmesApi.ExecuteAsync(
                RequireApiCredentials(session), LpaTool, LpaDetailParameters(query));
            var (columns, rows) = LpaRowsFromResult(result, "LQRNO", "LCITM");
            progress?.Report($"✓ {rows.Count:N0} row(s), {columns.Count} column(s).");
            return new BmesLpaScrapeService.LpaResult(columns, rows, null);
        }
        catch (Exception ex)
        {
            return BmesLpaScrapeService.LpaResult.Fail($"BMES API request failed: {ex.Message}");
        }
    }

    /// <summary>View C parameters; empty conditions are left out so the Tool applies its own defaults.</summary>
    public static Dictionary<string, object> LpaListParameters(
        BmesLpaScrapeService.LpaQuery query, DateTime from, DateTime to)
    {
        var parameters = new Dictionary<string, object>
        {
            ["ZTYPE"] = "C",
            ["SDATE"] = from.ToString("yyyy-MM-dd"),
            ["EDATE"] = to.ToString("yyyy-MM-dd"),
        };
        void Add(string name, string value) { if (!string.IsNullOrWhiteSpace(value)) parameters[name] = value.Trim(); }
        Add("FACCO", query.Facco);
        Add("LQRNO", query.Lqrno);
        Add("LQBNO", query.Lqbno);
        Add("AULOC", query.Auloc);
        Add("IMPLV", query.Implv);
        Add("DICNO", query.Dicno);
        Add("LASEQ", query.Laseq);
        Add("CHKER", query.Chker);
        Add("ZSTAT", query.Zstat);
        return parameters;
    }

    /// <summary>View D parameters (LQRNO exact; IMPLV and LCITM narrow the items).</summary>
    public static Dictionary<string, object> LpaDetailParameters(BmesLpaScrapeService.LpaDetailQuery query)
    {
        var parameters = new Dictionary<string, object> { ["ZTYPE"] = "D", ["LQRNO"] = query.Lqrno.Trim() };
        if (!string.IsNullOrWhiteSpace(query.Implv)) parameters["IMPLV"] = query.Implv.Trim();
        if (!string.IsNullOrWhiteSpace(query.Lcitm)) parameters["LCITM"] = query.Lcitm.Trim();
        return parameters;
    }

    /// <summary>Tool answer → generic rows: the result set that has the given key columns.</summary>
    public static (List<string> Columns, List<IReadOnlyDictionary<string, string>> Rows) LpaRowsFromResult(
        BmesApiResult result, params string[] keyColumns)
    {
        BmesApiResultSet? set = BmesToolRows.FindSet(result, keyColumns);
        return set is null ? ([], []) : BmesToolRows.ToGenericRows(set);
    }

    /// <summary>The screen's USEYN condition (Y = in use). The Tool has no such parameter, so it is applied here;
    /// a row without a USEYN value is kept.</summary>
    public static List<IReadOnlyDictionary<string, string>> FilterLpaUseYn(
        List<IReadOnlyDictionary<string, string>> rows, string useYn)
    {
        if (string.IsNullOrWhiteSpace(useYn)) return rows;
        return rows
            .Where(r => !r.TryGetValue("USEYN", out string? value) || value.Length == 0
                        || value.Equals(useYn.Trim(), StringComparison.OrdinalIgnoreCase))
            .ToList();
    }

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
}
