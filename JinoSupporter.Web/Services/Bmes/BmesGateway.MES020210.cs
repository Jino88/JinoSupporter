using System.Text.Json;

namespace JinoSupporter.Web.Services;

public static partial class BmesGateway
{
    // BMES API column name → internal column name (merged from DataMaker CONSTANT._columnMap + ListSTRManager)
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

    /// <summary>
    /// MES020210 — NG records (input and NG quantity per work order, process and NG code)
    /// for one plant over a date range.
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
        string url = $"{BmesWebSession.BaseUrl}/MES020210/SearchList?perPage=" +
                     $"&Condition.WERKS={werks}&Condition.SDATE={start}" +
                     $"&Condition.EDATE={end}&Condition.INPYN=N&Condition.USEYN=Y&page=1";
        try
        {
            var response = await session.Client.GetAsync(url);
            if (!response.IsSuccessStatusCode)
            {
                progress?.Report($"[WARN] WERKS {werks} response error: {response.StatusCode}");
                return null;
            }

            string json = await response.Content.ReadAsStringAsync();
            using var doc      = JsonDocument.Parse(json);
            var       contents = doc.RootElement.GetProperty("data").GetProperty("contents");

            var rows = new List<Dictionary<string, string>>();
            foreach (var item in contents.EnumerateArray())
            {
                var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var prop in item.EnumerateObject())
                {
                    if (prop.Name.Equals("VERID", StringComparison.OrdinalIgnoreCase))
                        continue; // skip removed column

                    string colName = NgRecordColumnMap.TryGetValue(prop.Name, out string? mapped)
                        ? mapped : prop.Name;

                    row[colName] = prop.Value.ValueKind == JsonValueKind.Null
                        ? string.Empty
                        : prop.Value.ToString();
                }
                rows.Add(row);
            }

            progress?.Report($"  WERKS {werks}: {rows.Count:N0} rows");
            return rows;
        }
        catch (Exception ex)
        {
            progress?.Report($"[WARN] WERKS {werks} parse error: {ex.Message}");
            return null;
        }
    }
}
