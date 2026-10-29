namespace PortalGNLAB.Web.Services;

public static partial class BmesGateway
{
    // Tool column name → internal column name (merged from DataMaker CONSTANT._columnMap + ListSTRManager).
    // AI_TOOL_00055 view S returns the MES020210 screen's own column names, so this is the same map as before.
    private static readonly Dictionary<string, string> NgRecordColumnMap =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["AUFNR"]    = "WORKORDER",
            ["WERKS"]    = "PLANT",
            ["VERID_TX"] = "PRODUCTION_LINE",
            ["ZSHIF"]    = "Shift",
            ["WDATE"]    = "PRODUCT_DATE",
            ["MATNR"]    = "MATERIALCODE",
            ["MAKTX"]    = "MATERIALNAME",
            ["KTSCH"]    = "PROCESSCODE",
            ["KTSCH_TX"] = "PROCESSNAME",
            ["ZCODE"]    = "NGCODE",
            ["ZCODE_TX"] = "NGNAME",
            ["INQTY_O"]  = "QTYINPUT",
            ["USEYN"]    = "USE",
            ["NGQTY_O"]  = "QTYNG",
            ["INQTY"]    = "INPUTQUANTITY",
            ["NGQTY"]    = "NGQUANTITY",
            ["PLNFL"]    = "WORKODER",
            ["VORNR"]    = "ACTIVITYNO",
            ["ERNAM"]    = "LASTREGPERSON",
            ["ERDAT"]    = "LASTREGDATE",
        };

    /// <summary>The API call for MES020210: AI_TOOL_00055 view S (saved input / NG per W/O × process × NG code).
    /// FRDAT..STDAT is the W/O date range; INPYN is left out so unsaved rows are listed like the screen's INPYN=N.</summary>
    public static Dictionary<string, object> NgRecordParameters(string werks, DateTime start, DateTime end) => new()
    {
        ["ZTYPE"] = "S",
        ["WERKS"] = werks,
        ["FRDAT"] = start.ToString("yyyy-MM-dd"),
        ["STDAT"] = end.ToString("yyyy-MM-dd"),
    };

    /// <summary>Tool answer → rows keyed by the app's internal column names (the removed VERID column is dropped).
    /// Every row is kept, including USEYN other than Y (e.g. H): the screen call returned them too, and its totals
    /// match the Tool's (checked 2026-10-27 plant 3200: 3,727 rows, input 26,474,836).</summary>
    public static List<Dictionary<string, string>> NgRecordRowsFromResult(BmesApiResult result)
    {
        BmesApiResultSet? set = BmesToolRows.FindSet(result, "AUFNR", "WDATE", "NGQTY_O");
        return set is null ? [] : BmesToolRows.ToRowDictionaries(set, NgRecordColumnMap, "VERID");
    }

    /// <summary>
    /// MES020210 — NG records (input and NG quantity per work order, process and NG code)
    /// for one plant over a date range, from PLATFORM_API (AI_TOOL_00055, view S).
    /// Ranges longer than 31 days are fetched in 31-day windows, one call each.
    /// </summary>
    /// <param name="werks">Plant code, e.g. 3200 or 3220.</param>
    /// <param name="start">First production date, yyyy-MM-dd.</param>
    /// <param name="end">Last production date, yyyy-MM-dd.</param>
    /// <returns>One dictionary per row keyed by the app's internal column names
    /// (PRODUCT_DATE, PROCESSNAME, NGNAME, QTYINPUT, QTYNG, ...); null when the call or
    /// the parse failed (a warning is reported).</returns>
    public static async Task<List<Dictionary<string, string>>?> FetchNgRecordsAsync(
        BmesWebSession session, string werks, string start, string end,
        IProgress<string>? progress)
    {
        try
        {
            BmesApiCredentials credentials = RequireApiCredentials(session);
            var rows = new List<Dictionary<string, string>>();
            foreach (var (windowStart, windowEnd) in BmesToolRows.SplitPeriod(DateTime.Parse(start), DateTime.Parse(end)))
            {
                BmesApiResult result = await BmesApi.ExecuteAsync(
                    credentials, "AI_TOOL_00055", NgRecordParameters(werks, windowStart, windowEnd));
                rows.AddRange(NgRecordRowsFromResult(result));
            }

            progress?.Report($"  WERKS {werks}: {rows.Count:N0} rows");
            return rows;
        }
        catch (Exception ex)
        {
            progress?.Report($"[WARN] WERKS {werks} API error: {ex.Message}");
            return null;
        }
    }
}
