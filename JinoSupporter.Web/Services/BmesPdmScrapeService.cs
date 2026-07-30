using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Fetches the drawing/spec master from BMES MES073300/SearchList and dumps every row
/// verbatim to a standalone SQLite DB (bmes_pdm_raw.db) next to ngrate_settings.db.
/// All values are stored as TEXT. Full refresh — the table is dropped and recreated on
/// every call.
/// </summary>
public sealed class BmesPdmScrapeService(NgRateSettingsService settings)
{
    // HTTPS, matching BmesRoutingScrapeService: over plain http the anti-forgery cookie does
    // not survive the redirect to https and the login POST fails validation.
    private const string BaseUrl = "https://bmes.bujeon.com";

    /// <summary>
    /// Search condition captured from the BMES UI. PDTNO is sent as observed — the response
    /// comes back with many different PDTNO values, so it does not appear to narrow the set.
    /// </summary>
    private const string SearchPayload =
        """
        {"Condition":{"LASTY":"L","SDATE":"1900-01-01","EDATE":"2050-12-31","PDMNO":"","PITTX":"","PDRTX":"","BMOTX":"","USEYN":"Y","PDTNO":"0001"}}
        """;

    // All columns observed in the MES073300/SearchList response.
    private static readonly string[] Columns =
    {
        "APSTA", "APSTA_TX",
        "PDMNO", "PDMVR", "PDMV",
        "PDTNO", "PDTNO_TX",
        "PDRTX", "PITTX", "PRRTX",
        "BMONO", "BMONO_TX", "BMOTX",
        "USEYN", "DELYN", "APPYN", "DETYN", "REVYN", "MODYN", "CWFYN",
        "APKEY", "DESCR",
        "ERNAM", "ERNAM_TX", "ERDAT",
        "ARNAM_TX", "ARDAT",
    };

    public string RawDbPath =>
        Path.Combine(settings.SettingsDbDirectory, "bmes_pdm_raw.db");

    /// <summary>
    /// Returns number of rows saved, or -1 on failure.
    /// </summary>
    public async Task<int> FetchAllAsync(IProgress<string>? progress = null)
    {
        string loginId = settings.LoginId;
        string password = settings.Password;
        if (string.IsNullOrWhiteSpace(loginId) || string.IsNullOrWhiteSpace(password))
        {
            progress?.Report("[ERROR] BMES credentials not configured.");
            return -1;
        }

        using var handler = new HttpClientHandler
        {
            UseCookies = true,
            CookieContainer = new CookieContainer(),
            AutomaticDecompression = DecompressionMethods.All,
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
            progress?.Report("[ERROR] Login failed — check credentials in Setting.");
            return -1;
        }

        progress?.Report("Calling MES073300/SearchList…");
        string json;
        try
        {
            using var content = new StringContent(SearchPayload, Encoding.UTF8, "application/json");
            var resp = await client.PostAsync($"{BaseUrl}/MES073300/SearchList", content);
            if (!resp.IsSuccessStatusCode)
            {
                progress?.Report($"[ERROR] HTTP {(int)resp.StatusCode} {resp.StatusCode}");
                return -1;
            }
            json = await resp.Content.ReadAsStringAsync();
        }
        catch (Exception ex)
        {
            progress?.Report($"[ERROR] Request failed: {ex.Message}");
            return -1;
        }

        var rows = new List<Dictionary<string, string>>();
        try
        {
            using var doc = JsonDocument.Parse(json);
            if (!doc.RootElement.TryGetProperty("data", out var dataEl) ||
                !dataEl.TryGetProperty("contents", out var contents) ||
                contents.ValueKind != JsonValueKind.Array)
            {
                progress?.Report("[ERROR] Response has no data.contents array.");
                return -1;
            }

            foreach (var item in contents.EnumerateArray())
            {
                var row = new Dictionary<string, string>(Columns.Length);
                foreach (var col in Columns)
                    row[col] = ReadStr(item, col);
                rows.Add(row);
            }
        }
        catch (Exception ex)
        {
            progress?.Report($"[ERROR] Parse failed: {ex.Message}");
            return -1;
        }

        progress?.Report($"Parsed {rows.Count:N0} rows. Saving to {Path.GetFileName(RawDbPath)}…");
        int saved = await Task.Run(() => SaveToSqlite(rows));
        progress?.Report($"✓ Saved {saved:N0} row(s) to bmes_pdm_raw.db");
        return saved;
    }

    /// <summary>
    /// Rows whose PDRTX (part name) or BMONO_TX (model name) contains <paramref name="search"/>.
    /// Returns an empty list when the raw DB has not been fetched yet.
    /// </summary>
    public List<BmesPdmRow> Search(string search, int maxRows = 300)
    {
        if (!File.Exists(RawDbPath))
            return [];

        using var conn = new SqliteConnection($"Data Source={RawDbPath};Mode=ReadOnly");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            """
            SELECT PDMNO, PDMVR, PDTNO_TX, PDRTX, BMONO_TX, USEYN, ERDAT
            FROM BmesPdm
            WHERE (@search = '' OR PDRTX LIKE @like OR BMONO_TX LIKE @like)
            -- BMONO_TX is populated on only a small fraction of rows, so ordering by it first
            -- would bury every linked row behind thousands of blanks.
            ORDER BY PDRTX, BMONO_TX
            LIMIT @maxRows;
            """;
        string term = (search ?? string.Empty).Trim();
        cmd.Parameters.AddWithValue("@search", term);
        cmd.Parameters.AddWithValue("@like", "%" + term + "%");
        cmd.Parameters.AddWithValue("@maxRows", Math.Clamp(maxRows, 1, 5000));

        var result = new List<BmesPdmRow>();
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new BmesPdmRow
            {
                Pdmno = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                Pdmvr = reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
                PdtnoTx = reader.IsDBNull(2) ? string.Empty : reader.GetString(2),
                Pdrtx = reader.IsDBNull(3) ? string.Empty : reader.GetString(3),
                BmonoTx = reader.IsDBNull(4) ? string.Empty : reader.GetString(4),
                Useyn = reader.IsDBNull(5) ? string.Empty : reader.GetString(5),
                Erdat = reader.IsDBNull(6) ? string.Empty : reader.GetString(6),
            });
        }
        return result;
    }

    /// <summary>
    /// Resolves the drawing attached to a PDMNO and returns its bytes.
    /// Chain: SearchDetail (find the file) → DownloadCheck (server-side permission check) →
    /// the file endpoint itself.
    /// </summary>
    public async Task<BmesPdmDownload> DownloadDrawingAsync(
        string pdmno,
        string pdmvr = "0",
        IProgress<string>? progress = null)
    {
        string docNo = (pdmno ?? string.Empty).Trim();
        if (docNo.Length == 0)
            return BmesPdmDownload.Failed("PDMNO is empty.");

        string loginId = settings.LoginId;
        string password = settings.Password;
        if (string.IsNullOrWhiteSpace(loginId) || string.IsNullOrWhiteSpace(password))
            return BmesPdmDownload.Failed("BMES credentials not configured.");

        using var handler = new HttpClientHandler
        {
            UseCookies = true,
            CookieContainer = new CookieContainer(),
            AutomaticDecompression = DecompressionMethods.All,
        };
        // No default X-Requested-With here: the file action is reached by navigation in the
        // BMES UI, and the header is only added to the two JSON calls below.
        using var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(300) };
        client.DefaultRequestHeaders.Referrer = new Uri($"{BaseUrl}/MES073300");

        string token = await GetTokenAsync(client);
        if (string.IsNullOrEmpty(token))
            return BmesPdmDownload.Failed("Failed to obtain verification token.");
        if (!await LoginAsync(client, token, loginId, password))
            return BmesPdmDownload.Failed("Login failed — check credentials in Setting.");

        string version = string.IsNullOrWhiteSpace(pdmvr) ? "0" : pdmvr.Trim();

        // 1. SearchDetail — locate the attached file.
        progress?.Report("Reading drawing detail…");
        string detailJson;
        try
        {
            string payload = JsonSerializer.Serialize(new
            {
                Condition = new { PDMNO = docNo, PDMVR = version },
            });
            var resp = await PostJsonAsync(client, $"{BaseUrl}/MES073300/SearchDetail", payload);
            if (!resp.IsSuccessStatusCode)
                return BmesPdmDownload.Failed($"SearchDetail HTTP {(int)resp.StatusCode}.");
            detailJson = await resp.Content.ReadAsStringAsync();
        }
        catch (Exception ex)
        {
            return BmesPdmDownload.Failed("SearchDetail failed: " + ex.Message);
        }

        string pattx, pattp, folder, physical, fileVersion;
        try
        {
            using var doc = JsonDocument.Parse(detailJson);
            if (!doc.RootElement.TryGetProperty("Files", out var files) ||
                files.ValueKind != JsonValueKind.Array ||
                files.GetArrayLength() == 0)
            {
                return BmesPdmDownload.Failed($"{docNo} has no attached file.");
            }

            // Prefer a usable DWG; otherwise fall back to the first attachment.
            JsonElement pick = files[0];
            foreach (var file in files.EnumerateArray())
            {
                bool usable = !string.Equals(ReadStr(file, "USEYN"), "N", StringComparison.OrdinalIgnoreCase);
                if (usable && ReadStr(file, "PATTP").Equals("DWG", StringComparison.OrdinalIgnoreCase))
                {
                    pick = file;
                    break;
                }
            }

            pattx = ReadStr(pick, "PATTX");
            pattp = ReadStr(pick, "PATTP");
            folder = ReadStr(pick, "P_FOLDE");
            physical = ReadStr(pick, "P_PATTX");

            // The file carries its own revision, which is not always the version asked for.
            fileVersion = ReadStr(pick, "PDMVR");
            if (string.IsNullOrWhiteSpace(fileVersion))
                fileVersion = version;
        }
        catch (Exception ex)
        {
            return BmesPdmDownload.Failed("SearchDetail parse failed: " + ex.Message);
        }

        // 2. DownloadCheck — the server refuses the file unless this passes first.
        progress?.Report($"Checking download permission for {pattx}…");
        try
        {
            string payload = JsonSerializer.Serialize(new
            {
                Download = new
                {
                    PDMNO = docNo,
                    PDMVR = fileVersion,
                    PATTX = pattx,
                    P_FOLDE = folder,
                    P_PATTX = physical,
                },
            });
            var resp = await PostJsonAsync(client, $"{BaseUrl}/MES073300/DownloadCheck", payload);
            string body = await resp.Content.ReadAsStringAsync();
            if (!resp.IsSuccessStatusCode)
                return BmesPdmDownload.Failed($"DownloadCheck HTTP {(int)resp.StatusCode}.");

            using var doc = JsonDocument.Parse(body);
            string result = doc.RootElement.TryGetProperty("Result", out var r) ? r.GetString() ?? "" : "";
            if (!string.Equals(result, "S", StringComparison.OrdinalIgnoreCase))
            {
                string msg = doc.RootElement.TryGetProperty("Msg", out var m) ? m.GetString() ?? "" : "";
                return BmesPdmDownload.Failed($"DownloadCheck refused: {result} {msg}".Trim());
            }
        }
        catch (Exception ex)
        {
            return BmesPdmDownload.Failed("DownloadCheck failed: " + ex.Message);
        }

        // 3. The file itself — a form-encoded POST, not JSON, using Download[...] field names.
        progress?.Report($"Downloading {pattx}…");
        try
        {
            using var form = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["Download[PDMNO]"] = docNo,
                ["Download[PDMVR]"] = fileVersion,
                ["Download[PATTX]"] = pattx,
                ["Download[P_FOLDE]"] = folder,
                ["Download[P_PATTX]"] = physical,
            });
            var resp = await client.PostAsync($"{BaseUrl}/MES073300/Download", form);
            if (!resp.IsSuccessStatusCode)
                return BmesPdmDownload.Failed($"Download HTTP {(int)resp.StatusCode}.");

            // A failure comes back as an HTML error page rather than a non-200 status.
            string mediaType = resp.Content.Headers.ContentType?.MediaType ?? string.Empty;
            if (resp.Content.Headers.ContentDisposition is null &&
                mediaType.Contains("html", StringComparison.OrdinalIgnoreCase))
            {
                return BmesPdmDownload.Failed("Download returned an HTML page instead of a file.");
            }

            byte[] bytes = await resp.Content.ReadAsByteArrayAsync();
            if (bytes.Length == 0)
                return BmesPdmDownload.Failed("Download returned an empty body.");

            string fileName = resp.Content.Headers.ContentDisposition?.FileNameStar
                ?? resp.Content.Headers.ContentDisposition?.FileName?.Trim('"')
                ?? (string.IsNullOrWhiteSpace(pattx) ? $"{docNo}.{pattp}" : pattx);

            return BmesPdmDownload.Ok(fileName, bytes);
        }
        catch (Exception ex)
        {
            return BmesPdmDownload.Failed("Download failed: " + ex.Message);
        }
    }

    private static Task<HttpResponseMessage> PostJsonAsync(HttpClient client, string url, string payload)
    {
        var request = new HttpRequestMessage(HttpMethod.Post, url)
        {
            Content = new StringContent(payload, Encoding.UTF8, "application/json"),
        };
        request.Headers.Add("X-Requested-With", "XMLHttpRequest");
        return client.SendAsync(request);
    }

    /// <summary>
    /// Rows whose PDRTX matches one of <paramref name="partNames"/> or whose BMONO_TX matches
    /// one of <paramref name="modelNames"/>. Used to narrow the PDM master down to what a
    /// loaded BOM actually contains. Matching is exact (trimmed, case-insensitive).
    /// </summary>
    public List<BmesPdmRow> SearchByNames(
        IEnumerable<string> partNames,
        IEnumerable<string> modelNames,
        int maxRows = 2000)
    {
        if (!File.Exists(RawDbPath))
            return [];

        var parts = Normalize(partNames);
        var models = Normalize(modelNames);
        if (parts.Count == 0 && models.Count == 0)
            return [];

        using var conn = new SqliteConnection($"Data Source={RawDbPath};Mode=ReadOnly");
        conn.Open();

        var found = new Dictionary<string, BmesPdmRow>(StringComparer.OrdinalIgnoreCase);

        // Batched so the parameter count stays well inside SQLite's limit.
        foreach (var batch in Batch(parts, 400))
            Collect(conn, "PDRTX", batch, found);
        foreach (var batch in Batch(models, 400))
            Collect(conn, "BMONO_TX", batch, found);

        return found.Values
            .OrderBy(r => r.Pdrtx, StringComparer.OrdinalIgnoreCase)
            .ThenBy(r => r.Pdmno, StringComparer.OrdinalIgnoreCase)
            .Take(Math.Clamp(maxRows, 1, 20000))
            .ToList();
    }

    private static void Collect(
        SqliteConnection conn,
        string column,
        List<string> values,
        Dictionary<string, BmesPdmRow> into)
    {
        string placeholders = string.Join(", ", values.Select((_, i) => $"@v{i}"));
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            $"""
            SELECT PDMNO, PDMVR, PDTNO_TX, PDRTX, BMONO_TX, USEYN, ERDAT
            FROM BmesPdm
            WHERE TRIM(UPPER({column})) IN ({placeholders});
            """;
        for (int i = 0; i < values.Count; i++)
            cmd.Parameters.AddWithValue($"@v{i}", values[i]);

        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            var row = new BmesPdmRow
            {
                Pdmno = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                Pdmvr = reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
                PdtnoTx = reader.IsDBNull(2) ? string.Empty : reader.GetString(2),
                Pdrtx = reader.IsDBNull(3) ? string.Empty : reader.GetString(3),
                BmonoTx = reader.IsDBNull(4) ? string.Empty : reader.GetString(4),
                Useyn = reader.IsDBNull(5) ? string.Empty : reader.GetString(5),
                Erdat = reader.IsDBNull(6) ? string.Empty : reader.GetString(6),
            };
            into[row.Pdmno + "\t" + row.Pdrtx] = row;
        }
    }

    private static List<string> Normalize(IEnumerable<string> values)
        => values
            .Select(v => (v ?? string.Empty).Trim().ToUpperInvariant())
            .Where(v => v.Length > 0)
            .Distinct(StringComparer.Ordinal)
            .ToList();

    private static IEnumerable<List<string>> Batch(List<string> values, int size)
    {
        for (int i = 0; i < values.Count; i += size)
            yield return values.GetRange(i, Math.Min(size, values.Count - i));
    }

    /// <summary>Row count and last fetch time, for showing whether a refresh is needed.</summary>
    public (int Rows, string FetchedAt) GetStatus()
    {
        if (!File.Exists(RawDbPath))
            return (0, string.Empty);

        try
        {
            using var conn = new SqliteConnection($"Data Source={RawDbPath};Mode=ReadOnly");
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*), MAX(FetchedAt) FROM BmesPdm;";
            using var reader = cmd.ExecuteReader();
            if (!reader.Read())
                return (0, string.Empty);
            return (
                reader.IsDBNull(0) ? 0 : reader.GetInt32(0),
                reader.IsDBNull(1) ? string.Empty : reader.GetString(1));
        }
        catch
        {
            return (0, string.Empty);
        }
    }

    private int SaveToSqlite(List<Dictionary<string, string>> rows)
    {
        using var conn = new SqliteConnection($"Data Source={RawDbPath}");
        conn.Open();

        string colDefs = string.Join(", ", Columns.Select(c => $"[{c}] TEXT"));
        using (var drop = conn.CreateCommand())
        {
            drop.CommandText = "DROP TABLE IF EXISTS [BmesPdm];";
            drop.ExecuteNonQuery();
        }
        using (var create = conn.CreateCommand())
        {
            create.CommandText = $"CREATE TABLE [BmesPdm] ({colDefs}, [FetchedAt] TEXT);";
            create.ExecuteNonQuery();
        }

        string colList = string.Join(", ", Columns.Select(c => $"[{c}]"));
        string paramList = string.Join(", ", Columns.Select((_, i) => $"@p{i}"));
        string fetchedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        using var tx = conn.BeginTransaction();
        using var ins = conn.CreateCommand();
        ins.CommandText = $"INSERT INTO [BmesPdm] ({colList}, [FetchedAt]) VALUES ({paramList}, @fetchedAt);";
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

        using (var index = conn.CreateCommand())
        {
            index.CommandText =
                "CREATE INDEX IF NOT EXISTS IX_BmesPdm_Search ON BmesPdm (BMONO_TX, PDRTX);";
            index.ExecuteNonQuery();
        }

        tx.Commit();
        return saved;
    }

    private static string ReadStr(JsonElement obj, string key)
    {
        if (!obj.TryGetProperty(key, out var prop)) return string.Empty;
        return prop.ValueKind switch
        {
            JsonValueKind.Null => string.Empty,
            JsonValueKind.String => prop.GetString() ?? string.Empty,
            _ => prop.ToString(),
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
            ["UserInfo.PWNO"] = password,
            ["UserInfo.LANG"] = "EN",
            ["UserInfo.FACCO"] = "GN",
            ["UserInfo.STYPE"] = "P",
            ["UserInfo.VTYPE"] = "P",
            ["__RequestVerificationToken"] = token,
        });
        try
        {
            var response = await client.PostAsync(BaseUrl + "/MES000000/LoginCheck", content);
            string body = await response.Content.ReadAsStringAsync();
            return body.Contains("\"Result\":\"M\"");
        }
        catch { return false; }
    }
}

/// <summary>Result of a drawing download attempt.</summary>
public sealed class BmesPdmDownload
{
    public bool Success { get; init; }
    public string FileName { get; init; } = string.Empty;
    public byte[] Bytes { get; init; } = [];
    public string Error { get; init; } = string.Empty;

    public static BmesPdmDownload Ok(string fileName, byte[] bytes)
        => new() { Success = true, FileName = fileName, Bytes = bytes };

    public static BmesPdmDownload Failed(string error)
        => new() { Success = false, Error = error };
}

/// <summary>One drawing/spec master row, trimmed to the fields used for picking a model.</summary>
public sealed class BmesPdmRow
{
    public string Pdmno { get; init; } = string.Empty;

    /// <summary>Document revision. Not always "0" — it is required to download the right file.</summary>
    public string Pdmvr { get; init; } = string.Empty;

    public string PdtnoTx { get; init; } = string.Empty;
    public string Pdrtx { get; init; } = string.Empty;
    public string BmonoTx { get; init; } = string.Empty;
    public string Useyn { get; init; } = string.Empty;
    public string Erdat { get; init; } = string.Empty;

    /// <summary>
    /// Name used as the BOM root. BMONO_TX (the model link) is filled on only a small
    /// fraction of the master, so PDRTX is the practical fallback.
    /// </summary>
    public string BomTarget =>
        string.IsNullOrWhiteSpace(BmonoTx) ? Pdrtx.Trim() : BmonoTx.Trim();
}
