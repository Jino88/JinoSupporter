namespace PortalGNLAB.Web.Services;

public static partial class BmesGateway
{
    /// <summary>Plants whose routing master is pulled, in call order.</summary>
    public static readonly string[] RoutingWerks = { "3200", "3210", "3220" };

    // All columns observed in the MES020240/SearchList response.
    public static readonly string[] RoutingColumns =
    {
        "WERKS", "MATNR", "MAKTX",
        "PTYPE", "PTYPE_TX",
        "PLNAL", "DATUV", "TAQTY", "TTIME", "SERNO",
        "VLSCH", "VLSCH_TX",
        "MAPNO", "MAPNO_TX",
        "OQMNO", "OQMNO_TX",
        "PLNF1", "VORN1", "PLNF2", "VORN2",
        "LGUBN", "LGUBN_TX",
        "RGUBN", "RGUBN_TX",
        "VGW01", "STPER",
        "MTYPE", "MTYPE_TX",
        "OQCRO", "CTQRO", "CCFRO",
        "STEUS",
        "PCTYP", "PCTYP_TX",
        "PCLVL", "PCLVL_TX",
        "BTYPE", "BTYPE_TX",
        "TTYPE", "TTYPE_TX",
    };

    /// <summary>
    /// MES020240/SearchList — routing master, one call per WERKS in <see cref="RoutingWerks"/>.
    /// Input: logged-in session. Output: every row of every plant as column→text dictionaries
    /// (keys = <see cref="RoutingColumns"/>), or null when a request or parse failed (the reason is
    /// reported through <paramref name="progress"/> as "[ERROR] ..."). A plant without a contents
    /// array is skipped with a "[WARN] ..." line.
    /// </summary>
    public static async Task<List<Dictionary<string, string>>?> FetchRoutingRowsAsync(
        BmesWebSession session, IProgress<string>? progress = null, CancellationToken cancellationToken = default)
    {
        var allRows = new List<Dictionary<string, string>>();
        foreach (var werks in RoutingWerks)
        {
            progress?.Report($"Calling MES020240/SearchList (WERKS={werks})…");
            string url =
                $"{BmesWebSession.BaseUrl}/MES020240/SearchList?perPage=" +
                $"&Condition.WERKS={werks}" +
                "&Condition.WPHNO=&Condition.MATNR=" +
                "&Condition.DATUV_S=1900-01-01&Condition.DATUV_E=2050-12-31" +
                "&Condition.FDATE=1900-01-01&Condition.TDATE=2050-12-31" +
                "&Condition.PTYPE=&Condition.BTYPE=&Condition.TTYPE=&page=1";

            HttpResponseMessage resp;
            try   { resp = await session.Client.GetAsync(url, cancellationToken); }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
            catch (Exception ex)
            {
                progress?.Report($"[ERROR] WERKS={werks} request failed: {ex.Message}");
                return null;
            }
            if (!resp.IsSuccessStatusCode)
            {
                progress?.Report($"[ERROR] WERKS={werks} HTTP {(int)resp.StatusCode} {resp.StatusCode}");
                return null;
            }

            string json = await resp.Content.ReadAsStringAsync(cancellationToken);
            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(json);
                if (!doc.RootElement.TryGetProperty("data", out var dataEl) ||
                    !dataEl.TryGetProperty("contents", out var contents) ||
                    contents.ValueKind != System.Text.Json.JsonValueKind.Array)
                {
                    progress?.Report($"[WARN] WERKS={werks}: no contents array.");
                    continue;
                }

                int countBefore = allRows.Count;
                foreach (var item in contents.EnumerateArray())
                {
                    var row = new Dictionary<string, string>(RoutingColumns.Length);
                    foreach (var col in RoutingColumns)
                        row[col] = ReadRowText(item, col);
                    allRows.Add(row);
                }
                progress?.Report($"  WERKS={werks}: +{allRows.Count - countBefore:N0} rows");
            }
            catch (Exception ex)
            {
                progress?.Report($"[ERROR] WERKS={werks} parse failed: {ex.Message}");
                return null;
            }
        }
        return allRows;
    }
}
