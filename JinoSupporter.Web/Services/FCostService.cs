using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

/// <summary>
/// BMES F-Cost (MES072400) data collection → SQLite save.
/// Mirrors NgRateService's auth flow (token → login → GET) but hits a different endpoint
/// that returns Input Amount / F-Cost / F-Cost Rate per (Model, Material) plus a Total row,
/// each carrying 14 pre-aggregated period columns (COL0001..COL0014).
/// Re-uses NgRateSettingsService for credentials + save directory so the user only has
/// to configure BMES login once.
/// </summary>
public sealed class FCostService(NgRateSettingsService settings, AppActivityLogger activity)
{
    private readonly NgRateSettingsService _settings = settings;
    private readonly AppActivityLogger _activity = activity;
    private const string BaseUrl = "http://bmes.bujeon.com";

    private void Log(IProgress<string>? progress, string msg)
    {
        _activity.Log("FCost", msg);
        progress?.Report(msg);
    }

    /// <summary>Login → fetch F-Cost rows for a single SDATE → write to a fresh
    /// SQLite file under DbSaveDirectory. Returns the saved DB path or null on failure.</summary>
    public async Task<string?> FetchAndSaveAsync(
        DateTime queryDate,
        IProgress<string>? progress = null)
    {
        Log(progress, "[start] Initializing…");

        string saveDir = _settings.FCostDbSaveDirectory;
        Directory.CreateDirectory(saveDir);

        string id  = _settings.LoginId;
        string pwd = _settings.Password;
        if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(pwd))
        {
            Log(progress, "[ERROR] BMES credentials not configured. Open Admin → Settings.");
            return null;
        }

        using var handler = new HttpClientHandler
        {
            UseCookies      = true,
            CookieContainer = new System.Net.CookieContainer(),
        };
        using var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(300) };
        client.DefaultRequestHeaders.Add("X-Requested-With", "XMLHttpRequest");

        Log(progress, "Fetching BMES login token…");
        string token = await GetTokenAsync(client);
        if (string.IsNullOrEmpty(token))
        {
            Log(progress, "[ERROR] Failed to read CSRF token from BMES home page.");
            return null;
        }

        Log(progress, $"Logging in as {id}…");
        if (!await LoginAsync(client, token, id, pwd))
        {
            Log(progress, "[ERROR] BMES login failed (check id/password).");
            return null;
        }

        Log(progress, $"Fetching F-Cost rows for SDATE={queryDate:yyyy-MM-dd}…");
        var fetched = await FetchFCostRowsAsync(client, queryDate, progress);
        if (fetched is null)
        {
            Log(progress, "[ERROR] F-Cost fetch failed.");
            return null;
        }

        string dbPath = Path.Combine(
            saveDir,
            $"fcost_{queryDate:yyyyMMdd}_{DateTime.Now:HHmmss}.db");

        Log(progress, $"Writing {fetched.Rows.Count:N0} rows + {fetched.Columns.Count} column metas → {dbPath}");
        await Task.Run(() => SaveRowsToSqlite(dbPath, queryDate, fetched.Rows, fetched.Columns));

        // Clean up older fcost_*.db files except the newest 5 to bound disk usage.
        try
        {
            var olds = Directory.GetFiles(saveDir, "fcost_????????_??????.db")
                .OrderByDescending(File.GetLastWriteTime)
                .Skip(5)
                .ToList();
            foreach (var p in olds) { try { File.Delete(p); } catch { } }
        }
        catch { }

        Log(progress, "[done]");
        return dbPath;
    }

    public string? FindMostRecentDb()
    {
        string dir = _settings.FCostDbSaveDirectory;
        if (!Directory.Exists(dir)) return null;
        return Directory.GetFiles(dir, "fcost_????????_??????.db")
            .OrderByDescending(File.GetLastWriteTime)
            .FirstOrDefault();
    }

    // ── HTTP ─────────────────────────────────────────────────────────────────────

    public string GetRawDbPath()
        => Path.Combine(_settings.FCostDbSaveDirectory, "fcost_raw.db");

    public FCostRawStatus GetRawStatus()
    {
        string dbPath = GetRawDbPath();
        var status = new FCostRawStatus
        {
            DbPath = dbPath,
            Exists = File.Exists(dbPath),
        };
        if (!status.Exists) return status;

        try
        {
            using var conn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
            conn.Open();

            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    """
                    SELECT
                        COUNT(*),
                        COALESCE(SUM(CASE WHEN Status = 'OK' THEN 1 ELSE 0 END), 0),
                        COALESCE(SUM(CASE WHEN Status <> 'OK' THEN 1 ELSE 0 END), 0),
                        COALESCE(MIN(QueryDate), ''),
                        COALESCE(MAX(QueryDate), '')
                    FROM FCostRawPulls;
                    """;
                using var r = cmd.ExecuteReader();
                if (r.Read())
                {
                    status.PullCount = Convert.ToInt32(r.GetValue(0));
                    status.SuccessCount = Convert.ToInt32(r.GetValue(1));
                    status.FailedCount = Convert.ToInt32(r.GetValue(2));
                    status.FirstDate = r.IsDBNull(3) ? string.Empty : r.GetString(3);
                    status.LastDate = r.IsDBNull(4) ? string.Empty : r.GetString(4);
                }
            }

            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT COUNT(*) FROM FCostRawRows;";
                status.TotalRows = Convert.ToInt32(cmd.ExecuteScalar() ?? 0);
            }
        }
        catch
        {
            status.PullCount = 0;
            status.SuccessCount = 0;
            status.FailedCount = 0;
            status.TotalRows = 0;
            status.FirstDate = string.Empty;
            status.LastDate = string.Empty;
        }

        return status;
    }

    public async Task<FCostRawBackfillResult> BackfillRawAsync(
        DateTime startDate,
        DateTime endDate,
        bool force = false,
        int delayMs = 1200,
        DateTime? forceFromDate = null,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        startDate = startDate.Date;
        endDate = endDate.Date;
        DateTime today = DateTime.Today;
        if (endDate > today) endDate = today;
        if (startDate > endDate)
            throw new ArgumentException("Start date must be before or equal to end date.");

        string saveDir = _settings.FCostDbSaveDirectory;
        Directory.CreateDirectory(saveDir);
        string dbPath = GetRawDbPath();
        await Task.Run(() => EnsureRawDb(dbPath), cancellationToken);

        var result = new FCostRawBackfillResult
        {
            DbPath = dbPath,
            StartDate = startDate,
            EndDate = endDate,
        };

        string id = _settings.LoginId;
        string pwd = _settings.Password;
        if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(pwd))
        {
            string msg = "BMES credentials are not configured.";
            Log(progress, "[ERROR] " + msg);
            result.Failures.Add(msg);
            return result;
        }

        using var handler = new HttpClientHandler
        {
            UseCookies = true,
            CookieContainer = new System.Net.CookieContainer(),
        };
        using var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(300) };
        client.DefaultRequestHeaders.Add("X-Requested-With", "XMLHttpRequest");

        Log(progress, "RAW backfill: login token");
        string token = await GetTokenAsync(client);
        if (string.IsNullOrEmpty(token))
        {
            string msg = "Failed to read BMES login token.";
            Log(progress, "[ERROR] " + msg);
            result.Failures.Add(msg);
            return result;
        }

        Log(progress, $"RAW backfill: login as {id}");
        if (!await LoginAsync(client, token, id, pwd))
        {
            string msg = "BMES login failed.";
            Log(progress, "[ERROR] " + msg);
            result.Failures.Add(msg);
            return result;
        }

        int totalDays = (endDate - startDate).Days + 1;
        int idx = 0;
        for (DateTime day = startDate; day <= endDate; day = day.AddDays(1))
        {
            cancellationToken.ThrowIfCancellationRequested();
            idx++;
            result.AttemptedDays++;

            bool mustRefetch = force || (forceFromDate is DateTime refreshFrom && day >= refreshFrom.Date);
            if (!mustRefetch && RawPullSucceeded(dbPath, day))
            {
                result.SkippedDays++;
                Log(progress, $"[{idx:N0}/{totalDays:N0}] {day:yyyy-MM-dd}: skip (already OK)");
                continue;
            }

            Log(progress, $"[{idx:N0}/{totalDays:N0}] {day:yyyy-MM-dd}: fetch");
            try
            {
                var fetched = await FetchFCostRowsAsync(client, day, progress);
                if (fetched is null)
                {
                    result.FailedDays++;
                    string msg = $"{day:yyyy-MM-dd}: fetch failed";
                    result.Failures.Add(msg);
                    await Task.Run(() => SaveRawFailure(dbPath, day, msg), cancellationToken);
                    continue;
                }

                await Task.Run(
                    () => SaveRawSuccess(dbPath, day, fetched.Rows, fetched.Columns),
                    cancellationToken);
                result.FetchedDays++;
                result.TotalRows += fetched.Rows.Count;
                Log(progress, $"{day:yyyy-MM-dd}: saved {fetched.Rows.Count:N0} rows");
            }
            catch (Exception ex)
            {
                result.FailedDays++;
                string msg = $"{day:yyyy-MM-dd}: {ex.Message}";
                result.Failures.Add(msg);
                await Task.Run(() => SaveRawFailure(dbPath, day, msg), cancellationToken);
                Log(progress, "[WARN] " + msg);
            }

            if (delayMs > 0 && day < endDate)
                await Task.Delay(delayMs, cancellationToken);
        }

        Log(progress, $"RAW backfill done: fetched={result.FetchedDays:N0}, skipped={result.SkippedDays:N0}, failed={result.FailedDays:N0}");
        return result;
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

    /// <summary>Parsed BMES F-Cost response: data rows + the per-response column metadata
    /// from BottomGridColumnList that tells us what date / week / month each COLnnnn means.</summary>
    private sealed record FetchedFCost(List<FCostRow> Rows, List<FCostColumnMeta> Columns);

    private static async Task<FetchedFCost?> FetchFCostRowsAsync(
        HttpClient client, DateTime sdate, IProgress<string>? progress)
    {
        // Query string mirrors what the BMES UI sends (captured via DevTools).
        // ZGUBN=D (day), ZTYPE=A (all), L_FACCO=GN (factory) — all empty filters left blank.
        string url =
            $"{BaseUrl}/MES072400/SearchList?perPage=" +
            $"&Condition.SDATE={sdate:yyyy-MM-dd}" +
            "&Condition.ZGUBN=D" +
            "&Condition.ZTYPE=A" +
            "&L_FACCO%5B0%5D=GN" +
            "&L_DIVIS%5B0%5D=" +
            "&L_CATEG%5B0%5D=" +
            "&L_PRODU%5B0%5D=" +
            "&L_ITEMS%5B0%5D=" +
            "&L_MODEL%5B0%5D=" +
            "&L_SERIE%5B0%5D=" +
            "&L_MODEL2%5B0%5D=" +
            "&L_VERID%5B0%5D=" +
            "&L_MATNR%5B0%5D=" +
            "&REMEM=N&page=1";

        try
        {
            var response = await client.GetAsync(url);
            if (!response.IsSuccessStatusCode)
            {
                progress?.Report($"[WARN] F-Cost response error: {response.StatusCode}");
                return null;
            }

            string json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);

            // Server returns { result, message, data: { contents: [...] } }
            var root = doc.RootElement;
            if (root.TryGetProperty("result", out var ok) && ok.ValueKind == JsonValueKind.False)
            {
                progress?.Report("[WARN] BMES returned result=false");
                return null;
            }

            var dataObj  = root.GetProperty("data");
            var contents = dataObj.GetProperty("contents");
            var rows     = new List<FCostRow>(capacity: contents.GetArrayLength());

            foreach (var item in contents.EnumerateArray())
            {
                rows.Add(ParseRow(item));
            }

            // BMES emits each Material as a strict INAMT → FCOST → FRATE triplet but blanks
            // out the metadata text fields on the FCOST/FRATE continuation rows (visual de-dup
            // for their own UI). Forward-fill from the most recent INAMT so every saved row
            // is self-contained and any later DB query / export sees the full Material context.
            ForwardFillTriplets(rows);

            // BottomGridColumnList tells us what date / week / month each COLnnnn represents
            // for THIS response. Older / newer queries shift the mapping (today moves), so we
            // persist it alongside the rows rather than hard-coding it.
            var columns = ParseBottomGridColumnList(dataObj);

            progress?.Report($"  F-Cost: {rows.Count:N0} rows · {columns.Count} column meta");
            return new FetchedFCost(rows, columns);
        }
        catch (Exception ex)
        {
            progress?.Report($"[WARN] F-Cost parse error: {ex.Message}");
            return null;
        }
    }

    /// <summary>Reads the response's BottomGridColumnList — one entry per COLnnnn — and
    /// returns the 14-row metadata. Falls back to an empty list if the response shape is
    /// older / different (the report layer then uses generic "C1..C14" headers).</summary>
    private static List<FCostColumnMeta> ParseBottomGridColumnList(JsonElement dataObj)
    {
        var list = new List<FCostColumnMeta>();
        if (!dataObj.TryGetProperty("BottomGridColumnList", out var arr) ||
            arr.ValueKind != JsonValueKind.Array)
        {
            return list;
        }

        foreach (var item in arr.EnumerateArray())
        {
            string code   = item.TryGetProperty("ZPVTT",    out var z)  ? (z.GetString()  ?? "") : "";
            string header = item.TryGetProperty("ZPVTT_TX", out var zt) ? (zt.GetString() ?? "") : "";
            string pdate  = item.TryGetProperty("PDATE",    out var pd) ? (pd.GetString() ?? "") : "";
            string colName = item.TryGetProperty("ZCOLN",   out var cn) ? (cn.GetString() ?? "") : "";

            // ZCOLN is "COL0001".."COL0014" — strip prefix to recover the 1..14 index.
            int index = 0;
            if (colName.Length > 3 && int.TryParse(colName[3..], out int parsed))
                index = parsed;

            if (index < 1 || index > 14) continue;

            list.Add(new FCostColumnMeta
            {
                Index  = index,
                Code   = code,
                Header = header,
                PDate  = pdate,
            });
        }

        // Sort by index so the page can iterate in 1..14 order.
        list.Sort((a, b) => a.Index.CompareTo(b.Index));
        return list;
    }

    // ── Forward-fill blank metadata across INAMT/FCOST/FRATE triplets ───────────

    /// <summary>BMES blanks out the _TX (and many code) fields on FCOST/FRATE continuation
    /// rows because their own UI relies on Excel-style "see the row above" rendering. We
    /// copy them down so each saved row stands on its own — useful for direct SQL queries,
    /// CSV exports, or anyone reading rows out of order.</summary>
    private static void ForwardFillTriplets(List<FCostRow> rows)
    {
        FCostRow? anchor = null;
        foreach (var r in rows)
        {
            // INAMT marks the start of a new Material — it always carries full metadata.
            if (string.Equals(r.ZType, "INAMT", StringComparison.OrdinalIgnoreCase))
            {
                anchor = r;
                continue;
            }
            if (anchor is null) continue;

            // Only fill blanks; never overwrite a value the row legitimately carries.
            if (string.IsNullOrEmpty(r.FaccoTx)) r.FaccoTx = anchor.FaccoTx;
            if (string.IsNullOrEmpty(r.PrdGrTx)) r.PrdGrTx = anchor.PrdGrTx;
            if (string.IsNullOrEmpty(r.VeridTx)) r.VeridTx = anchor.VeridTx;
            if (string.IsNullOrEmpty(r.ModNoTx)) r.ModNoTx = anchor.ModNoTx;
            if (string.IsNullOrEmpty(r.AssemTx)) r.AssemTx = anchor.AssemTx;
            if (string.IsNullOrEmpty(r.AbChgTx)) r.AbChgTx = anchor.AbChgTx;
            if (string.IsNullOrEmpty(r.MCodeTx)) r.MCodeTx = anchor.MCodeTx;
            if (string.IsNullOrEmpty(r.MatnrTx)) r.MatnrTx = anchor.MatnrTx;

            if (string.IsNullOrEmpty(r.Facco)) r.Facco = anchor.Facco;
            if (string.IsNullOrEmpty(r.Werks)) r.Werks = anchor.Werks;
            if (string.IsNullOrEmpty(r.PrdGr)) r.PrdGr = anchor.PrdGr;
            if (string.IsNullOrEmpty(r.ModNo)) r.ModNo = anchor.ModNo;
            if (string.IsNullOrEmpty(r.Verid)) r.Verid = anchor.Verid;
            if (string.IsNullOrEmpty(r.AbChg)) r.AbChg = anchor.AbChg;
            if (string.IsNullOrEmpty(r.Cat01)) r.Cat01 = anchor.Cat01;
            if (string.IsNullOrEmpty(r.Matnr)) r.Matnr = anchor.Matnr;
            if (string.IsNullOrEmpty(r.Assem)) r.Assem = anchor.Assem;
            if (string.IsNullOrEmpty(r.TwSyn)) r.TwSyn = anchor.TwSyn;
        }
    }

    // ── JSON → row ───────────────────────────────────────────────────────────────

    private static FCostRow ParseRow(JsonElement item)
    {
        string S(string n) =>
            item.TryGetProperty(n, out var p) && p.ValueKind != JsonValueKind.Null
                ? p.ToString()
                : string.Empty;

        // BMES emits numbers as strings ("364672.000"). Parse with invariant culture.
        double D(string n)
        {
            if (!item.TryGetProperty(n, out var p) || p.ValueKind == JsonValueKind.Null)
                return 0;
            string raw = p.ToString();
            return double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double v)
                ? v : 0;
        }

        return new FCostRow
        {
            RawJson  = item.GetRawText(),
            FaccoTx  = S("FACCO_TX"),  PrdGrTx = S("PRDGR_TX"), VeridTx = S("VERID_TX"),
            ModNoTx  = S("MODNO_TX"),  AssemTx = S("ASSEM_TX"), AbChgTx = S("ABCHG_TX"),
            MCodeTx  = S("MCODE_TX"),  MatnrTx = S("MATNR_TX"), ZTypeTx = S("ZTYPE_TX"),
            TSort    = S("TSORT"),     ZSort   = S("ZSORT"),    ZType   = S("ZTYPE"),
            Facco    = S("FACCO"),     Werks   = S("WERKS"),    PrdGr   = S("PRDGR"),
            ModNo    = S("MODNO"),     Verid   = S("VERID"),    AbChg   = S("ABCHG"),
            Cat01    = S("CAT01"),     Matnr   = S("MATNR"),    Assem   = S("ASSEM"),
            TwSyn    = S("TWSYN"),     ZValu   = S("ZVALU"),
            Col0001  = D("COL0001"), Col0002 = D("COL0002"), Col0003 = D("COL0003"),
            Col0004  = D("COL0004"), Col0005 = D("COL0005"), Col0006 = D("COL0006"),
            Col0007  = D("COL0007"), Col0008 = D("COL0008"), Col0009 = D("COL0009"),
            Col0010  = D("COL0010"), Col0011 = D("COL0011"), Col0012 = D("COL0012"),
            Col0013  = D("COL0013"), Col0014 = D("COL0014"),
        };
    }

    // ── SQLite ───────────────────────────────────────────────────────────────────

    private static bool RawPullSucceeded(string dbPath, DateTime queryDate)
    {
        try
        {
            using var conn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT Status FROM FCostRawPulls WHERE QueryDate = @qdate;";
            cmd.Parameters.AddWithValue("@qdate", queryDate.ToString("yyyy-MM-dd"));
            return string.Equals(cmd.ExecuteScalar() as string, "OK", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static void EnsureRawDb(string dbPath)
    {
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            """
            PRAGMA journal_mode=WAL;

            CREATE TABLE IF NOT EXISTS FCostRawPulls (
                QueryDate   TEXT PRIMARY KEY,
                Status      TEXT NOT NULL,
                FetchedAt   TEXT NOT NULL,
                RowCount    INTEGER NOT NULL DEFAULT 0,
                ColumnCount INTEGER NOT NULL DEFAULT 0,
                ErrorMessage TEXT NOT NULL DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS FCostRawRows (
                QueryDate TEXT NOT NULL,
                RowNo INTEGER NOT NULL,
                FetchedAt TEXT NOT NULL,
                RawJson TEXT NOT NULL DEFAULT '',
                FaccoTx TEXT, PrdGrTx TEXT, VeridTx TEXT, ModNoTx TEXT,
                AssemTx TEXT, AbChgTx TEXT, MCodeTx TEXT, MatnrTx TEXT,
                ZTypeTx TEXT, ZType TEXT, TSort TEXT, ZSort TEXT,
                Facco TEXT, Werks TEXT, PrdGr TEXT, ModNo TEXT, Verid TEXT,
                AbChg TEXT, Cat01 TEXT, Matnr TEXT, Assem TEXT,
                TwSyn TEXT, ZValu TEXT,
                Col0001 REAL, Col0002 REAL, Col0003 REAL, Col0004 REAL,
                Col0005 REAL, Col0006 REAL, Col0007 REAL, Col0008 REAL,
                Col0009 REAL, Col0010 REAL, Col0011 REAL, Col0012 REAL,
                Col0013 REAL, Col0014 REAL,
                PRIMARY KEY (QueryDate, RowNo)
            );
            CREATE INDEX IF NOT EXISTS IX_FCostRawRows_QueryDate ON FCostRawRows(QueryDate);
            CREATE INDEX IF NOT EXISTS IX_FCostRawRows_Matnr ON FCostRawRows(Matnr);
            CREATE INDEX IF NOT EXISTS IX_FCostRawRows_ZType ON FCostRawRows(ZType);
            CREATE INDEX IF NOT EXISTS IX_FCostRawRows_Werks ON FCostRawRows(Werks);

            CREATE TABLE IF NOT EXISTS FCostRawColumns (
                QueryDate TEXT NOT NULL,
                ColIndex INTEGER NOT NULL,
                Code TEXT NOT NULL,
                Header TEXT NOT NULL,
                PDate TEXT NOT NULL,
                Kind TEXT NOT NULL,
                PRIMARY KEY (QueryDate, ColIndex)
            );
            CREATE INDEX IF NOT EXISTS IX_FCostRawColumns_PDate ON FCostRawColumns(PDate);
            """;
        cmd.ExecuteNonQuery();
    }

    private static void SaveRawFailure(string dbPath, DateTime queryDate, string errorMessage)
    {
        EnsureRawDb(dbPath);
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            """
            INSERT INTO FCostRawPulls (QueryDate, Status, FetchedAt, RowCount, ColumnCount, ErrorMessage)
            VALUES (@qdate, 'FAILED', @fetched, 0, 0, @error)
            ON CONFLICT(QueryDate) DO UPDATE SET
                Status = excluded.Status,
                FetchedAt = excluded.FetchedAt,
                ErrorMessage = excluded.ErrorMessage;
            """;
        cmd.Parameters.AddWithValue("@qdate", queryDate.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@fetched", DateTime.Now.ToString("O"));
        cmd.Parameters.AddWithValue("@error", errorMessage);
        cmd.ExecuteNonQuery();
    }

    private static void SaveRawSuccess(
        string dbPath,
        DateTime queryDate,
        List<FCostRow> rows,
        List<FCostColumnMeta> columns)
    {
        EnsureRawDb(dbPath);
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        using var tx = conn.BeginTransaction();
        string qIso = queryDate.ToString("yyyy-MM-dd");
        string nowIso = DateTime.Now.ToString("O");

        using (var del = conn.CreateCommand())
        {
            del.Transaction = tx;
            del.CommandText =
                """
                DELETE FROM FCostRawRows WHERE QueryDate = @qdate;
                DELETE FROM FCostRawColumns WHERE QueryDate = @qdate;
                """;
            del.Parameters.AddWithValue("@qdate", qIso);
            del.ExecuteNonQuery();
        }

        using (var pull = conn.CreateCommand())
        {
            pull.Transaction = tx;
            pull.CommandText =
                """
                INSERT INTO FCostRawPulls (QueryDate, Status, FetchedAt, RowCount, ColumnCount, ErrorMessage)
                VALUES (@qdate, 'OK', @fetched, @rows, @cols, '')
                ON CONFLICT(QueryDate) DO UPDATE SET
                    Status = excluded.Status,
                    FetchedAt = excluded.FetchedAt,
                    RowCount = excluded.RowCount,
                    ColumnCount = excluded.ColumnCount,
                    ErrorMessage = excluded.ErrorMessage;
                """;
            pull.Parameters.AddWithValue("@qdate", qIso);
            pull.Parameters.AddWithValue("@fetched", nowIso);
            pull.Parameters.AddWithValue("@rows", rows.Count);
            pull.Parameters.AddWithValue("@cols", columns.Count);
            pull.ExecuteNonQuery();
        }

        if (columns.Count > 0)
        {
            using var colIns = conn.CreateCommand();
            colIns.Transaction = tx;
            colIns.CommandText =
                """
                INSERT INTO FCostRawColumns (QueryDate, ColIndex, Code, Header, PDate, Kind)
                VALUES (@qdate, @idx, @code, @header, @pdate, @kind);
                """;
            var cpQ = colIns.Parameters.Add("@qdate", SqliteType.Text);
            var cpI = colIns.Parameters.Add("@idx", SqliteType.Integer);
            var cpC = colIns.Parameters.Add("@code", SqliteType.Text);
            var cpH = colIns.Parameters.Add("@header", SqliteType.Text);
            var cpP = colIns.Parameters.Add("@pdate", SqliteType.Text);
            var cpK = colIns.Parameters.Add("@kind", SqliteType.Text);
            colIns.Prepare();

            foreach (var c in columns)
            {
                cpQ.Value = qIso;
                cpI.Value = c.Index;
                cpC.Value = c.Code;
                cpH.Value = c.Header;
                cpP.Value = c.PDate;
                cpK.Value = c.Kind.ToString();
                colIns.ExecuteNonQuery();
            }
        }

        using var ins = conn.CreateCommand();
        ins.Transaction = tx;
        ins.CommandText =
            """
            INSERT INTO FCostRawRows
            (QueryDate, RowNo, FetchedAt, RawJson, FaccoTx, PrdGrTx, VeridTx, ModNoTx,
             AssemTx, AbChgTx, MCodeTx, MatnrTx, ZTypeTx, ZType, TSort, ZSort,
             Facco, Werks, PrdGr, ModNo, Verid, AbChg, Cat01, Matnr, Assem,
             TwSyn, ZValu,
             Col0001, Col0002, Col0003, Col0004, Col0005, Col0006, Col0007,
             Col0008, Col0009, Col0010, Col0011, Col0012, Col0013, Col0014)
            VALUES
            (@qdate, @rowno, @fetched, @rawjson, @facco_tx, @prdgr_tx, @verid_tx, @modno_tx,
             @assem_tx, @abchg_tx, @mcode_tx, @matnr_tx, @ztype_tx, @ztype, @tsort, @zsort,
             @facco, @werks, @prdgr, @modno, @verid, @abchg, @cat01, @matnr, @assem,
             @twsyn, @zvalu,
             @c01, @c02, @c03, @c04, @c05, @c06, @c07, @c08, @c09, @c10,
             @c11, @c12, @c13, @c14);
            """;

        SqliteParameter P(string n, SqliteType t) => ins.Parameters.Add(n, t);
        var pQDate    = P("@qdate",    SqliteType.Text);
        var pRowNo    = P("@rowno",    SqliteType.Integer);
        var pFetched  = P("@fetched",  SqliteType.Text);
        var pRawJson  = P("@rawjson",  SqliteType.Text);
        var pFaccoTx  = P("@facco_tx", SqliteType.Text);
        var pPrdGrTx  = P("@prdgr_tx", SqliteType.Text);
        var pVeridTx  = P("@verid_tx", SqliteType.Text);
        var pModNoTx  = P("@modno_tx", SqliteType.Text);
        var pAssemTx  = P("@assem_tx", SqliteType.Text);
        var pAbChgTx  = P("@abchg_tx", SqliteType.Text);
        var pMCodeTx  = P("@mcode_tx", SqliteType.Text);
        var pMatnrTx  = P("@matnr_tx", SqliteType.Text);
        var pZTypeTx  = P("@ztype_tx", SqliteType.Text);
        var pZType    = P("@ztype",    SqliteType.Text);
        var pTSort    = P("@tsort",    SqliteType.Text);
        var pZSort    = P("@zsort",    SqliteType.Text);
        var pFacco    = P("@facco",    SqliteType.Text);
        var pWerks    = P("@werks",    SqliteType.Text);
        var pPrdGr    = P("@prdgr",    SqliteType.Text);
        var pModNo    = P("@modno",    SqliteType.Text);
        var pVerid    = P("@verid",    SqliteType.Text);
        var pAbChg    = P("@abchg",    SqliteType.Text);
        var pCat01    = P("@cat01",    SqliteType.Text);
        var pMatnr    = P("@matnr",    SqliteType.Text);
        var pAssem    = P("@assem",    SqliteType.Text);
        var pTwSyn    = P("@twsyn",    SqliteType.Text);
        var pZValu    = P("@zvalu",    SqliteType.Text);
        var pC = new SqliteParameter[14];
        for (int i = 0; i < 14; i++) pC[i] = P($"@c{i + 1:D2}", SqliteType.Real);
        ins.Prepare();

        for (int i = 0; i < rows.Count; i++)
        {
            var r = rows[i];
            pQDate.Value = qIso;
            pRowNo.Value = i + 1;
            pFetched.Value = nowIso;
            pRawJson.Value = r.RawJson;
            pFaccoTx.Value = r.FaccoTx; pPrdGrTx.Value = r.PrdGrTx; pVeridTx.Value = r.VeridTx;
            pModNoTx.Value = r.ModNoTx; pAssemTx.Value = r.AssemTx; pAbChgTx.Value = r.AbChgTx;
            pMCodeTx.Value = r.MCodeTx; pMatnrTx.Value = r.MatnrTx; pZTypeTx.Value = r.ZTypeTx;
            pZType.Value   = r.ZType;   pTSort.Value   = r.TSort;   pZSort.Value   = r.ZSort;
            pFacco.Value   = r.Facco;   pWerks.Value   = r.Werks;   pPrdGr.Value   = r.PrdGr;
            pModNo.Value   = r.ModNo;   pVerid.Value   = r.Verid;   pAbChg.Value   = r.AbChg;
            pCat01.Value   = r.Cat01;   pMatnr.Value   = r.Matnr;   pAssem.Value   = r.Assem;
            pTwSyn.Value   = r.TwSyn;   pZValu.Value   = r.ZValu;
            pC[0].Value  = r.Col0001;  pC[1].Value  = r.Col0002;  pC[2].Value  = r.Col0003;
            pC[3].Value  = r.Col0004;  pC[4].Value  = r.Col0005;  pC[5].Value  = r.Col0006;
            pC[6].Value  = r.Col0007;  pC[7].Value  = r.Col0008;  pC[8].Value  = r.Col0009;
            pC[9].Value  = r.Col0010;  pC[10].Value = r.Col0011;  pC[11].Value = r.Col0012;
            pC[12].Value = r.Col0013;  pC[13].Value = r.Col0014;
            ins.ExecuteNonQuery();
        }

        tx.Commit();
    }

    private static void SaveRowsToSqlite(
        string dbPath,
        DateTime queryDate,
        List<FCostRow> rows,
        List<FCostColumnMeta> columns)
    {
        if (File.Exists(dbPath)) File.Delete(dbPath);
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText =
                """
                CREATE TABLE FCostRows (
                    Id        INTEGER PRIMARY KEY AUTOINCREMENT,
                    FetchedAt TEXT NOT NULL,
                    QueryDate TEXT NOT NULL,
                    FaccoTx TEXT, PrdGrTx TEXT, VeridTx TEXT, ModNoTx TEXT,
                    AssemTx TEXT, AbChgTx TEXT, MCodeTx TEXT, MatnrTx TEXT,
                    ZTypeTx TEXT, ZType TEXT, TSort TEXT, ZSort TEXT,
                    Facco TEXT, Werks TEXT, PrdGr TEXT, ModNo TEXT, Verid TEXT,
                    AbChg TEXT, Cat01 TEXT, Matnr TEXT, Assem TEXT,
                    TwSyn TEXT, ZValu TEXT,
                    Col0001 REAL, Col0002 REAL, Col0003 REAL, Col0004 REAL,
                    Col0005 REAL, Col0006 REAL, Col0007 REAL, Col0008 REAL,
                    Col0009 REAL, Col0010 REAL, Col0011 REAL, Col0012 REAL,
                    Col0013 REAL, Col0014 REAL
                );
                CREATE INDEX IX_FCostRows_Matnr ON FCostRows(Matnr);
                CREATE INDEX IX_FCostRows_ZType ON FCostRows(ZType);

                CREATE TABLE FCostColumns (
                    ColIndex INTEGER PRIMARY KEY,   -- 1..14
                    Code     TEXT NOT NULL,         -- ZPVTT  e.g. "20260504" / "2026-19" / "202605"
                    Header   TEXT NOT NULL,         -- ZPVTT_TX e.g. "05-04" / "26-W19" / "26-05"
                    PDate    TEXT NOT NULL          -- normalized PDATE
                );
                """;
            cmd.ExecuteNonQuery();
        }

        // Insert column metadata first (small fixed-size table — 14 rows).
        if (columns.Count > 0)
        {
            using var colTx  = conn.BeginTransaction();
            using var colIns = conn.CreateCommand();
            colIns.Transaction = colTx;
            colIns.CommandText =
                "INSERT INTO FCostColumns (ColIndex, Code, Header, PDate) " +
                "VALUES (@idx, @code, @hdr, @pdate);";
            var pIdx   = colIns.Parameters.Add("@idx",   SqliteType.Integer);
            var pCode  = colIns.Parameters.Add("@code",  SqliteType.Text);
            var pHdr   = colIns.Parameters.Add("@hdr",   SqliteType.Text);
            var pPDate = colIns.Parameters.Add("@pdate", SqliteType.Text);
            colIns.Prepare();

            foreach (var c in columns)
            {
                pIdx.Value   = c.Index;
                pCode.Value  = c.Code;
                pHdr.Value   = c.Header;
                pPDate.Value = c.PDate;
                colIns.ExecuteNonQuery();
            }
            colTx.Commit();
        }

        using var tx = conn.BeginTransaction();
        using var ins = conn.CreateCommand();
        ins.Transaction = tx;
        ins.CommandText =
            """
            INSERT INTO FCostRows
            (FetchedAt, QueryDate, FaccoTx, PrdGrTx, VeridTx, ModNoTx,
             AssemTx, AbChgTx, MCodeTx, MatnrTx, ZTypeTx, ZType, TSort, ZSort,
             Facco, Werks, PrdGr, ModNo, Verid, AbChg, Cat01, Matnr, Assem,
             TwSyn, ZValu,
             Col0001, Col0002, Col0003, Col0004, Col0005, Col0006, Col0007,
             Col0008, Col0009, Col0010, Col0011, Col0012, Col0013, Col0014)
            VALUES
            (@fetched, @qdate, @facco_tx, @prdgr_tx, @verid_tx, @modno_tx,
             @assem_tx, @abchg_tx, @mcode_tx, @matnr_tx, @ztype_tx, @ztype, @tsort, @zsort,
             @facco, @werks, @prdgr, @modno, @verid, @abchg, @cat01, @matnr, @assem,
             @twsyn, @zvalu,
             @c01, @c02, @c03, @c04, @c05, @c06, @c07, @c08, @c09, @c10,
             @c11, @c12, @c13, @c14);
            """;

        // Pre-create parameters once and reset Value per row for speed.
        SqliteParameter P(string n, SqliteType t) => ins.Parameters.Add(n, t);
        var pFetched  = P("@fetched",  SqliteType.Text);
        var pQDate    = P("@qdate",    SqliteType.Text);
        var pFaccoTx  = P("@facco_tx", SqliteType.Text);
        var pPrdGrTx  = P("@prdgr_tx", SqliteType.Text);
        var pVeridTx  = P("@verid_tx", SqliteType.Text);
        var pModNoTx  = P("@modno_tx", SqliteType.Text);
        var pAssemTx  = P("@assem_tx", SqliteType.Text);
        var pAbChgTx  = P("@abchg_tx", SqliteType.Text);
        var pMCodeTx  = P("@mcode_tx", SqliteType.Text);
        var pMatnrTx  = P("@matnr_tx", SqliteType.Text);
        var pZTypeTx  = P("@ztype_tx", SqliteType.Text);
        var pZType    = P("@ztype",    SqliteType.Text);
        var pTSort    = P("@tsort",    SqliteType.Text);
        var pZSort    = P("@zsort",    SqliteType.Text);
        var pFacco    = P("@facco",    SqliteType.Text);
        var pWerks    = P("@werks",    SqliteType.Text);
        var pPrdGr    = P("@prdgr",    SqliteType.Text);
        var pModNo    = P("@modno",    SqliteType.Text);
        var pVerid    = P("@verid",    SqliteType.Text);
        var pAbChg    = P("@abchg",    SqliteType.Text);
        var pCat01    = P("@cat01",    SqliteType.Text);
        var pMatnr    = P("@matnr",    SqliteType.Text);
        var pAssem    = P("@assem",    SqliteType.Text);
        var pTwSyn    = P("@twsyn",    SqliteType.Text);
        var pZValu    = P("@zvalu",    SqliteType.Text);
        var pC = new SqliteParameter[14];
        for (int i = 0; i < 14; i++) pC[i] = P($"@c{i + 1:D2}", SqliteType.Real);
        ins.Prepare();

        string nowIso = DateTime.Now.ToString("O");
        string qIso   = queryDate.ToString("yyyy-MM-dd");
        foreach (var r in rows)
        {
            pFetched.Value = nowIso;
            pQDate.Value   = qIso;
            pFaccoTx.Value = r.FaccoTx; pPrdGrTx.Value = r.PrdGrTx; pVeridTx.Value = r.VeridTx;
            pModNoTx.Value = r.ModNoTx; pAssemTx.Value = r.AssemTx; pAbChgTx.Value = r.AbChgTx;
            pMCodeTx.Value = r.MCodeTx; pMatnrTx.Value = r.MatnrTx; pZTypeTx.Value = r.ZTypeTx;
            pZType.Value   = r.ZType;   pTSort.Value   = r.TSort;   pZSort.Value   = r.ZSort;
            pFacco.Value   = r.Facco;   pWerks.Value   = r.Werks;   pPrdGr.Value   = r.PrdGr;
            pModNo.Value   = r.ModNo;   pVerid.Value   = r.Verid;   pAbChg.Value   = r.AbChg;
            pCat01.Value   = r.Cat01;   pMatnr.Value   = r.Matnr;   pAssem.Value   = r.Assem;
            pTwSyn.Value   = r.TwSyn;   pZValu.Value   = r.ZValu;
            pC[0].Value  = r.Col0001;  pC[1].Value  = r.Col0002;  pC[2].Value  = r.Col0003;
            pC[3].Value  = r.Col0004;  pC[4].Value  = r.Col0005;  pC[5].Value  = r.Col0006;
            pC[6].Value  = r.Col0007;  pC[7].Value  = r.Col0008;  pC[8].Value  = r.Col0009;
            pC[9].Value  = r.Col0010;  pC[10].Value = r.Col0011;  pC[11].Value = r.Col0012;
            pC[12].Value = r.Col0013;  pC[13].Value = r.Col0014;
            ins.ExecuteNonQuery();
        }
        tx.Commit();
    }
}
