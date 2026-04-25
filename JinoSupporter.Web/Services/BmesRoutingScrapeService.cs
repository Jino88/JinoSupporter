using System.Net;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Fetches routing master records from BMES MES020240/SearchList for each WERKS
/// (3200, 3210, 3220) and dumps all rows verbatim to a standalone SQLite DB
/// (bmes_routing_raw.db) next to ngrate_settings.db. Column set mirrors the API
/// response; all values are stored as TEXT. Full refresh — table is dropped and
/// recreated on every call.
/// </summary>
public sealed class BmesRoutingScrapeService(NgRateSettingsService settings)
{
    private const string BaseUrl = "http://bmes.bujeon.com";

    private static readonly string[] Werks = { "3200", "3210", "3220" };

    // All columns observed in the MES020240/SearchList response.
    private static readonly string[] Columns =
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

    public string RawDbPath =>
        Path.Combine(settings.SettingsDbDirectory, "bmes_routing_raw.db");

    /// <summary>
    /// Returns number of rows saved across all WERKS requests, or -1 on failure.
    /// </summary>
    public async Task<int> FetchAllAsync(IProgress<string>? progress = null)
    {
        string loginId  = settings.LoginId;
        string password = settings.Password;
        if (string.IsNullOrWhiteSpace(loginId) || string.IsNullOrWhiteSpace(password))
        {
            progress?.Report("[ERROR] BMES credentials not configured.");
            return -1;
        }

        using var handler = new HttpClientHandler
        {
            UseCookies      = true,
            CookieContainer = new CookieContainer(),
        };
        using var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(300) };
        client.DefaultRequestHeaders.Add("X-Requested-With", "XMLHttpRequest");

        progress?.Report("Fetching verification token…");
        string token = await GetTokenAsync(client);
        if (string.IsNullOrEmpty(token))
        {
            progress?.Report("[ERROR] Failed to obtain token.");
            return -1;
        }

        progress?.Report("Logging in…");
        if (!await LoginAsync(client, token, loginId, password))
        {
            progress?.Report("[ERROR] Login failed — check credentials in BMES Setting.");
            return -1;
        }

        var allRows = new List<Dictionary<string, string>>();
        foreach (var werks in Werks)
        {
            progress?.Report($"Calling MES020240/SearchList (WERKS={werks})…");
            string url =
                $"{BaseUrl}/MES020240/SearchList?perPage=" +
                $"&Condition.WERKS={werks}" +
                "&Condition.WPHNO=&Condition.MATNR=" +
                "&Condition.DATUV_S=1900-01-01&Condition.DATUV_E=2050-12-31" +
                "&Condition.FDATE=1900-01-01&Condition.TDATE=2050-12-31" +
                "&Condition.PTYPE=&Condition.BTYPE=&Condition.TTYPE=&page=1";

            HttpResponseMessage resp;
            try   { resp = await client.GetAsync(url); }
            catch (Exception ex)
            {
                progress?.Report($"[ERROR] WERKS={werks} request failed: {ex.Message}");
                return -1;
            }
            if (!resp.IsSuccessStatusCode)
            {
                progress?.Report($"[ERROR] WERKS={werks} HTTP {(int)resp.StatusCode} {resp.StatusCode}");
                return -1;
            }

            string json = await resp.Content.ReadAsStringAsync();
            try
            {
                using var doc = JsonDocument.Parse(json);
                if (!doc.RootElement.TryGetProperty("data", out var dataEl) ||
                    !dataEl.TryGetProperty("contents", out var contents) ||
                    contents.ValueKind != JsonValueKind.Array)
                {
                    progress?.Report($"[WARN] WERKS={werks}: no contents array.");
                    continue;
                }

                int countBefore = allRows.Count;
                foreach (var item in contents.EnumerateArray())
                {
                    var row = new Dictionary<string, string>(Columns.Length);
                    foreach (var col in Columns)
                        row[col] = ReadStr(item, col);
                    allRows.Add(row);
                }
                progress?.Report($"  WERKS={werks}: +{allRows.Count - countBefore:N0} rows");
            }
            catch (Exception ex)
            {
                progress?.Report($"[ERROR] WERKS={werks} parse failed: {ex.Message}");
                return -1;
            }
        }

        progress?.Report($"Parsed {allRows.Count:N0} total rows. Saving to {Path.GetFileName(RawDbPath)}…");
        int saved = await Task.Run(() => SaveToSqlite(allRows));
        progress?.Report($"✓ Saved {saved:N0} row(s) to bmes_routing_raw.db");

        progress?.Report("Merging into RoutingTable…");
        int added = await Task.Run(() => settings.MergeRoutingFromRawDb(RawDbPath, progress));
        progress?.Report(added > 0
            ? $"✓ Done. Saved {saved:N0} raw / added {added:N0} new RoutingTable row(s)."
            : $"✓ Done. Saved {saved:N0} raw / no new RoutingTable rows.");
        return saved;
    }

    private int SaveToSqlite(List<Dictionary<string, string>> rows)
    {
        using var conn = new SqliteConnection($"Data Source={RawDbPath}");
        conn.Open();

        string colDefs = string.Join(", ", Columns.Select(c => $"[{c}] TEXT"));
        using (var drop = conn.CreateCommand())
        {
            drop.CommandText = "DROP TABLE IF EXISTS [BmesRouting];";
            drop.ExecuteNonQuery();
        }
        using (var create = conn.CreateCommand())
        {
            create.CommandText = $"CREATE TABLE [BmesRouting] ({colDefs}, [FetchedAt] TEXT);";
            create.ExecuteNonQuery();
        }

        string colList   = string.Join(", ", Columns.Select(c => $"[{c}]"));
        string paramList = string.Join(", ", Columns.Select((_, i) => $"@p{i}"));
        string fetchedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        using var tx  = conn.BeginTransaction();
        using var ins = conn.CreateCommand();
        ins.CommandText = $"INSERT INTO [BmesRouting] ({colList}, [FetchedAt]) VALUES ({paramList}, @fetchedAt);";
        for (int i = 0; i < Columns.Length; i++) ins.Parameters.Add(new SqliteParameter($"@p{i}", string.Empty));
        ins.Parameters.Add(new SqliteParameter("@fetchedAt", fetchedAt));

        int saved = 0;
        foreach (var row in rows)
        {
            for (int i = 0; i < Columns.Length; i++)
                ins.Parameters[i].Value = row.TryGetValue(Columns[i], out var v) ? v : string.Empty;
            ins.Parameters[Columns.Length].Value = fetchedAt;
            ins.ExecuteNonQuery();
            saved++;
        }
        tx.Commit();
        return saved;
    }

    private static string ReadStr(JsonElement obj, string key)
    {
        if (!obj.TryGetProperty(key, out var prop)) return string.Empty;
        return prop.ValueKind switch
        {
            JsonValueKind.Null   => string.Empty,
            JsonValueKind.String => prop.GetString() ?? string.Empty,
            _                    => prop.ToString(),
        };
    }

    private static async Task<string> GetTokenAsync(HttpClient client)
    {
        try
        {
            string html = await client.GetStringAsync(BaseUrl);
            var m = Regex.Match(html,
                @"<input[^>]+name=""__RequestVerificationToken""[^>]+value=""([^""]+)""",
                RegexOptions.IgnoreCase);
            if (!m.Success)
                m = Regex.Match(html,
                    @"<input[^>]+value=""([^""]+)""[^>]+name=""__RequestVerificationToken""",
                    RegexOptions.IgnoreCase);
            return m.Success ? m.Groups[1].Value : string.Empty;
        }
        catch { return string.Empty; }
    }

    private static async Task<bool> LoginAsync(
        HttpClient client, string token, string id, string password)
    {
        var content = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["UserInfo.USRID"] = id,
            ["UserInfo.PWNO"]  = password,
            ["UserInfo.LANG"]  = "EN",
            ["UserInfo.FACCO"] = "GN",
            ["UserInfo.STYPE"] = "P",
            ["UserInfo.VTYPE"] = "P",
            ["__RequestVerificationToken"] = token,
        });
        try
        {
            var response = await client.PostAsync(BaseUrl + "/MES000000/LoginCheck", content);
            string body  = await response.Content.ReadAsStringAsync();
            return body.Contains("\"Result\":\"M\"");
        }
        catch { return false; }
    }
}
