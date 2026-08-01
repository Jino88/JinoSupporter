using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

/// <summary>
/// BMES NG Rate data collection → SQLite save → Processing service.
/// Web port of the DataMaker WPF app's clFetchBMES + clFetchBMESNGDATA + clDataProcessor logic.
/// Paths / credentials are read from NgRateSettingsService.
/// </summary>
public sealed class NgRateService(
    NgRateSettingsService settings,
    AppActivityLogger activity,
    BmesSettingsSyncService settingsSync)
{
    private readonly NgRateSettingsService _settings = settings;
    private readonly AppActivityLogger _activity = activity;
    private readonly BmesSettingsSyncService _settingsSync = settingsSync;
    private const string BaseUrl = "https://bmes.bujeon.com";
    private const string DailyCacheSubdirectory = "daily";
    private const string MonthlyCacheSubdirectory = "monthly";
    private static readonly object ReusableDbLock = new();
    private static ReusableDb? _lastReusableDb;
    private static readonly TimeSpan RecentRangeReuseTtl = TimeSpan.FromMinutes(15);
    private static readonly TimeSpan RecentSourceCacheTtl = TimeSpan.FromMinutes(15);

    private sealed record ReusableDb(
        DateTime StartDate,
        DateTime EndDate,
        string LineShiftFilterKey,
        DateTimeOffset CachedAt,
        string DbPath);
    private sealed record BmesLoginResult(bool Succeeded, string Diagnostic);

    /// <summary>Mirrors a progress message to both the UI and the Debug/launching-console output.
    /// Use this for every status line so you can diagnose stalls live in the dotnet console
    /// and the VS "Debug" output simultaneously.</summary>
    private void Log(IProgress<string>? progress, string msg)
    {
        _activity.Log("NgRate", msg);
        progress?.Report(msg);
    }

    // BMES API column name → internal column name (merged from DataMaker CONSTANT._columnMap + ListSTRManager)
    private static readonly Dictionary<string, string> ApiColumnMap =
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

    // OrginalTable column order (per DataMaker CONSTANT.ListSTRManager)
    private static readonly string[] OrgTableColumns =
    {
        "PRODUCTION_LINE", "PROCESSCODE", "PROCESSNAME", "NGCODE", "NGNAME",
        "USE", "QTYINPUT", "QTYNG", "INPUTQUANTITY", "NGQUANTITY",
        "MATERIALCODE", "MATERIALNAME", "PLANT", "WORKORDER", "Shift",
        "PRODUCT_DATE", "WORKODER", "ACTIVITYNO", "LASTREGPERSON", "LASTREGDATE",
    };

    // ── Public API ──────────────────────────────────────────────────────────────

    /// <summary>
    /// Full flow: BMES data fetch → SQLite save → processing.
    /// - Dates 3+ days ago : use monthly DB cache (fetch + cache if missing)
    /// - Dates 0–2 days ago: always fetch from server (data may still change)
    /// - Result is merged into a temp DB (temp_*.db); previous temp DBs are auto-deleted.
    /// </summary>
    public async Task<string?> GetOrFetchAsync(
        DateTime startDate,
        DateTime endDate,
        IProgress<string>? progress = null,
        IReadOnlyCollection<string>? lineShiftFilter = null)
    {
        HashSet<string>? normalizedLineShifts = NormalizeLineShiftFilter(lineShiftFilter);
        string filterKey = BuildLineShiftFilterKey(normalizedLineShifts);
        if (TryGetReusableDb(startDate, endDate, filterKey, out string? dbPath))
        {
            Log(progress, $"Reusing fetched DB: {Path.GetFileName(dbPath)}");
            return dbPath;
        }

        return await FetchAndSaveAsync(startDate, endDate, progress, normalizedLineShifts);
    }

    public async Task<string?> FetchAndSaveAsync(
        DateTime startDate,
        DateTime endDate,
        IProgress<string>? progress = null,
        IReadOnlyCollection<string>? lineShiftFilter = null)
    {
        HashSet<string>? normalizedLineShifts = NormalizeLineShiftFilter(lineShiftFilter);
        string filterKey = BuildLineShiftFilterKey(normalizedLineShifts);

        // Immediate ack so the UI shows the operation has actually started
        // (helps diagnose slow path-config / DB-init / cache-scan delays).
        var swStart = System.Diagnostics.Stopwatch.StartNew();
        progress?.Report("[start] Initializing…");

        // Resolve & cache the cache root once (avoids one ngrate_settings.db
        // round-trip per date being classified).
        progress?.Report("[start] Resolving NgRate save directory…");
        string saveDir = _settings.DbSaveDirectory;
        if (!_settings.IsNgRateStorageConfigured || string.IsNullOrWhiteSpace(saveDir))
        {
            progress?.Report("[ERROR] Configure the Working Folder in Setting first.");
            return null;
        }

        progress?.Report("[start] Using local Routing/Reason tables.");
        progress?.Report($"[start] Save dir = {saveDir}  ({swStart.ElapsedMilliseconds} ms elapsed)");

        // ── 1. Classify dates ────────────────────────────────────────────────
        // · Recent three days   → reuse monthly cache for a short TTL
        // · Otherwise           → use monthly DB cache if present, else fetch
        // Recent three days use monthly cache markers for a short TTL.
        var recentCutoff = DateTime.Today.AddDays(-2);
        var allDates = Enumerable
            .Range(0, (int)(endDate.Date - startDate.Date).TotalDays + 1)
            .Select(i => startDate.Date.AddDays(i))
            .ToList();

        var toFetch = new List<DateTime>(); // needs BMES server fetch
        var toCache = new List<DateTime>(); // load from monthly DB cache

        var oldDates = allDates.Where(d => d < recentCutoff).ToList();
        await Task.Run(() => EnsureMonthlyCacheFromDaily(oldDates, progress));
        var monthlyCachedDates = GetMonthlyReusableCachedDateSet(allDates, recentCutoff);

        foreach (var date in allDates)
        {
            if (!monthlyCachedDates.Contains(date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)))
                toFetch.Add(date);
            else
                toCache.Add(date);
        }
        progress?.Report($"[start] Date scan done in {swStart.ElapsedMilliseconds} ms.");

        progress?.Report(
            $"Date range: {startDate:MM/dd} – {endDate:MM/dd}  " +
            $"(server: {toFetch.Count} day(s) / cache: {toCache.Count} day(s))");
        if (normalizedLineShifts is { Count: > 0 })
            progress?.Report($"LineShift pre-filter: {normalizedLineShifts.Count:N0} selected.");

        // ── 2. BMES server fetch ─────────────────────────────────────────────
        var freshRows = new List<Dictionary<string, string>>();

        if (toFetch.Count > 0)
        {
            progress?.Report($"─── Server fetch  {toFetch.Min():MM/dd} – {toFetch.Max():MM/dd} ({toFetch.Count} day(s))");

            string loginId  = _settings.LoginId;
            string password = _settings.Password;
            if (string.IsNullOrWhiteSpace(loginId) || string.IsNullOrWhiteSpace(password))
            {
                progress?.Report("[ERROR] BMES credentials not configured. Ask admin to set them in NG Rate Settings.");
                return null;
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
                return null;
            }

            progress?.Report("Logging in to BMES…");
            BmesLoginResult login = await LoginAsync(client, token, loginId, password);
            if (!login.Succeeded)
            {
                progress?.Report($"[ERROR] Login failed — {login.Diagnostic}");
                return null;
            }
            progress?.Report("Login successful.");

            // Merge consecutive dates into ranges; split the request when the sequence breaks
            var serverRows = new List<Dictionary<string, string>>();

            var sorted = toFetch.OrderBy(d => d).ToList();
            var ranges = new List<(DateTime Start, DateTime End)>();
            var rs = sorted[0]; var re = sorted[0];
            for (int i = 1; i < sorted.Count; i++)
            {
                if (sorted[i] == re.AddDays(1)) { re = sorted[i]; }
                else { ranges.Add((rs, re)); rs = sorted[i]; re = sorted[i]; }
            }
            ranges.Add((rs, re));

            // The server fetch dominates the run and its size is known up front, so
            // report it as counted work — consumers can draw a real progress bar.
            int fetched = 0;
            progress.ReportSteps(0, ranges.Count, "Server fetch");

            foreach (var (start, end) in ranges)
            {
                string s     = start.ToString("yyyy-MM-dd");
                string e     = end.ToString("yyyy-MM-dd");
                string label = start == end ? $"{start:MM/dd}" : $"{start:MM/dd} – {end:MM/dd}";
                progress?.Report($"Fetching {label}… ({fetched + 1}/{ranges.Count})");
                var fetch3200 = FetchRawRowsAsync(client, "3200", s, e, progress);
                var fetch3220 = FetchRawRowsAsync(client, "3220", s, e, progress);
                await Task.WhenAll(fetch3200, fetch3220);
                var rows3200 = await fetch3200;
                var rows3220 = await fetch3220;
                if (rows3200 != null) serverRows.AddRange(rows3200);
                if (rows3220 != null) serverRows.AddRange(rows3220);

                progress.ReportSteps(++fetched, ranges.Count, "Server fetch");
            }

            Log(progress, $"Collected {serverRows.Count:N0} rows. Removing duplicates…");
            // Heavy CPU work — push to threadpool so the Blazor renderer keeps
            // processing queued progress callbacks (no UI freeze).
            serverRows = await Task.Run(() => RemoveDuplicates(serverRows));
            Log(progress, $"{serverRows.Count:N0} rows after deduplication.");

            // Cache every fetched date. Recent markers are reused only for a short TTL.
            foreach (var date in toFetch)
            {
                string dateStr = date.ToString("yyyy-MM-dd");
                var    dayRows = serverRows
                    .Where(r => GetCol(r, "PRODUCT_DATE") == dateStr)
                    .ToList();
                await Task.Run(() => SaveMonthlyDateDb(date, dayRows, progress));
            }

            freshRows = FilterRowsByLineShift(serverRows, normalizedLineShifts);
            if (normalizedLineShifts is { Count: > 0 })
                Log(progress, $"  Recent rows after LineShift filter: {freshRows.Count:N0}");
        }

        // ── 3. Load monthly cache ───────────────────────────────────────────
        if (toCache.Count > 0)
            Log(progress, $"─── Cache load  {toCache.Min():MM/dd} – {toCache.Max():MM/dd} ({toCache.Count} day(s))");

        var cachedRows = new List<Dictionary<string, string>>();
        foreach (var monthGroup in toCache.GroupBy(d => new DateTime(d.Year, d.Month, 1)).OrderBy(g => g.Key))
        {
            var dates = monthGroup.OrderBy(d => d).ToList();
            var rows = await Task.Run(
                () => LoadFromMonthlyDb(monthGroup.Key, dates, normalizedLineShifts));
            cachedRows.AddRange(rows);
            Log(
                progress,
                normalizedLineShifts is { Count: > 0 }
                    ? $"  Monthly cache filtered {monthGroup.Key:yyyy-MM}: {rows.Count:N0} rows ({dates.Count:N0} day(s))"
                    : $"  Monthly cache hit {monthGroup.Key:yyyy-MM}: {rows.Count:N0} rows ({dates.Count:N0} day(s))");
        }

        // ── 4. Merge ─────────────────────────────────────────────────────────
        var allRows = new List<Dictionary<string, string>>(freshRows.Count + cachedRows.Count);
        allRows.AddRange(freshRows);
        allRows.AddRange(cachedRows);

        if (allRows.Count == 0)
        {
            progress?.Report("[ERROR] No data available for the selected date range.");
            return null;
        }

        // ── 5. Clean up old temp DBs → create new temp DB ───────────────────
        await Task.Run(() => CleanupTempDbs(progress));

        string tempPath = GetTempDbPath();
        try { Directory.CreateDirectory(Path.GetDirectoryName(tempPath)!); }
        catch (Exception ex)
        {
            Log(progress, $"[ERROR] Cannot create output folder: {ex.Message}");
            return null;
        }

        var swSave = System.Diagnostics.Stopwatch.StartNew();
        Log(progress, $"Saving to temp DB: {Path.GetFileName(tempPath)} ({allRows.Count:N0} rows)");
        await Task.Run(() => SaveToSqlite(tempPath, allRows));
        Log(progress, $"[FetchAndSave] SaveToSqlite done ({swSave.ElapsedMilliseconds} ms)");

        // ── 6. Post-processing ───────────────────────────────────────────────
        var swProc = System.Diagnostics.Stopwatch.StartNew();
        await EnsureServerTablesAsync(progress);
        Log(progress, "Running post-processing (Routing / Reason / LineShift)…");
        await Task.Run(() => ProcessData(tempPath, progress));
        Log(progress, $"[FetchAndSave] ProcessData done ({swProc.ElapsedMilliseconds} ms)");

        RememberReusableDb(startDate, endDate, filterKey, tempPath);
        Log(progress, $"Done. DB: {Path.GetFileName(tempPath)}  (total {swStart.ElapsedMilliseconds} ms since start)");
        return tempPath;
    }

    private bool TryGetReusableDb(
        DateTime startDate,
        DateTime endDate,
        string lineShiftFilterKey,
        out string? dbPath)
    {
        lock (ReusableDbLock)
        {
            if (_lastReusableDb is { } cached
                && cached.StartDate == startDate.Date
                && cached.EndDate == endDate.Date
                && string.Equals(cached.LineShiftFilterKey, lineShiftFilterKey, StringComparison.Ordinal)
                && (endDate.Date < DateTime.Today.AddDays(-2) ||
                    DateTimeOffset.Now - cached.CachedAt <= RecentRangeReuseTtl)
                && File.Exists(cached.DbPath))
            {
                dbPath = cached.DbPath;
                return true;
            }
        }

        dbPath = null;
        return false;
    }

    private void RememberReusableDb(
        DateTime startDate,
        DateTime endDate,
        string lineShiftFilterKey,
        string dbPath)
    {
        lock (ReusableDbLock)
        {
            _lastReusableDb = new ReusableDb(
                startDate.Date,
                endDate.Date,
                lineShiftFilterKey,
                DateTimeOffset.Now,
                dbPath);
        }
    }

    private async Task EnsureServerTablesAsync(IProgress<string>? progress)
    {
        try
        {
            var routingRows = _settings.GetRoutingRows();
            var reasonRows = _settings.GetReasonRows();
            bool needRouting = routingRows.Count == 0
                || routingRows.All(r => string.IsNullOrWhiteSpace(r.ProcessType));
            bool needReason = reasonRows.Count == 0;

            if (!needRouting && !needReason)
                return;

            if (needRouting)
            {
                Log(progress, "Routing table is missing. Loading from server…");
                var routing = await _settingsSync.PullRoutingRowsFromServerAsync();
                Log(progress, routing.Succeeded
                    ? $"Routing table loaded: {routing.Rows:N0} rows."
                    : $"[WARN] Routing table load failed: {routing.Message}");
            }

            if (needReason)
            {
                Log(progress, "Reason table is missing. Loading from server…");
                var reason = await _settingsSync.PullReasonRowsFromServerAsync();
                Log(progress, reason.Succeeded
                    ? $"Reason table loaded: {reason.Rows:N0} rows."
                    : $"[WARN] Reason table load failed: {reason.Message}");
            }
        }
        catch (Exception ex)
        {
            Log(progress, $"[WARN] Server table auto-load failed: {ex.Message}");
        }
    }

    // ── Monthly DB / Temp DB helpers ─────────────────────────────────────────

    /// <summary>Per-day cache DB path: {DbSaveDirectory}/daily/yyyyMMdd.db</summary>
    private string GetPerDayDbPath(DateTime date)
        => Path.Combine(_settings.DbSaveDirectory, DailyCacheSubdirectory, $"{date:yyyyMMdd}.db");

    /// <summary>Monthly cache DB path: {DbSaveDirectory}/monthly/yyyyMM.db</summary>
    private string GetMonthlyDbPath(DateTime date)
        => Path.Combine(_settings.DbSaveDirectory, MonthlyCacheSubdirectory, $"{date:yyyyMM}.db");

    /// <summary>Temp merged DB path: {DbSaveDirectory}/temp_yyyyMMdd_HHmmss.db</summary>
    private string GetTempDbPath()
    {
        string ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        return Path.Combine(_settings.DbSaveDirectory, $"temp_{ts}.db");
    }

    private void SaveMonthlyDateDb(
        DateTime date, List<Dictionary<string, string>> rows, IProgress<string>? progress)
    {
        try
        {
            string path = GetMonthlyDbPath(date);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);

            using var conn = new SqliteConnection($"Data Source={path}");
            conn.Open();
            EnsureMonthlyCacheSchema(conn);
            using var tx = conn.BeginTransaction();
            ReplaceMonthlyDateRows(conn, tx, date, rows);
            tx.Commit();
            SaveMeta(conn);

            progress?.Report(rows.Count == 0
                ? $"  Cached {date:MM/dd}: no data → {Path.GetFileName(path)}"
                : $"  Cached {date:MM/dd}: {rows.Count:N0} rows → {Path.GetFileName(path)}");
        }
        catch (Exception ex)
        {
            progress?.Report($"[WARN] Failed to cache {date:MM/dd}: {ex.Message}");
        }
    }

    private void EnsureMonthlyCacheFromDaily(IReadOnlyList<DateTime> requestedDates, IProgress<string>? progress)
    {
        if (requestedDates.Count == 0)
            return;

        string dailyDir = Path.Combine(_settings.DbSaveDirectory, DailyCacheSubdirectory);
        if (!Directory.Exists(dailyDir))
            return;

        var requestedDateSet = requestedDates
            .Select(d => d.Date)
            .ToHashSet();

        var dailyFiles = Directory
            .EnumerateFiles(dailyDir, "????????.db", SearchOption.TopDirectoryOnly)
            .Select(path => (Path: path, Date: TryParseDailyDbDate(path)))
            .Where(item => item.Date.HasValue)
            .Select(item => (item.Path, Date: item.Date!.Value))
            .Where(item => requestedDateSet.Contains(item.Date.Date))
            .GroupBy(item => new DateTime(item.Date.Year, item.Date.Month, 1))
            .OrderBy(group => group.Key)
            .ToList();

        if (dailyFiles.Count == 0)
            return;

        string monthlyDir = Path.Combine(_settings.DbSaveDirectory, MonthlyCacheSubdirectory);
        Directory.CreateDirectory(monthlyDir);

        foreach (var monthGroup in dailyFiles)
        {
            string monthPath = GetMonthlyDbPath(monthGroup.Key);
            var cachedDates = GetMonthlyCachedDates(monthPath);
            var missing = monthGroup
                .Where(item => !cachedDates.Contains(item.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)))
                .OrderBy(item => item.Date)
                .ToList();

            if (missing.Count == 0)
                continue;

            var sw = System.Diagnostics.Stopwatch.StartNew();
            int rowCount = 0;
            using var conn = new SqliteConnection($"Data Source={monthPath}");
            conn.Open();
            EnsureMonthlyCacheSchema(conn);
            using var tx = conn.BeginTransaction();

            foreach (var item in missing)
            {
                var rows = LoadFromPerDayDb(item.Path);
                ReplaceMonthlyDateRows(conn, tx, item.Date, rows);
                rowCount += rows.Count;
            }

            tx.Commit();
            SaveMeta(conn);
            Log(progress, $"  Monthly cache built {monthGroup.Key:yyyy-MM}: {missing.Count:N0} day(s), {rowCount:N0} rows ({sw.ElapsedMilliseconds} ms)");
        }
    }

    private HashSet<string> GetMonthlyReusableCachedDateSet(
        IReadOnlyList<DateTime> dates,
        DateTime recentCutoff)
    {
        var result = new HashSet<string>(StringComparer.Ordinal);
        foreach (var month in dates.Select(d => new DateTime(d.Year, d.Month, 1)).Distinct())
        {
            string monthPath = GetMonthlyDbPath(month);
            if (!File.Exists(monthPath))
                continue;

            try
            {
                using var conn = new SqliteConnection($"Data Source={monthPath};Mode=ReadOnly");
                conn.Open();
                if (!TableExists(conn, "__NgRateCachedDates"))
                    continue;

                using var cmd = conn.CreateCommand();
                cmd.CommandText =
                    "SELECT [PRODUCT_DATE], [CachedAt] FROM [__NgRateCachedDates]";
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string dateText = reader.IsDBNull(0) ? string.Empty : reader.GetString(0);
                    if (!DateTime.TryParseExact(
                            dateText,
                            "yyyy-MM-dd",
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None,
                            out DateTime date))
                    {
                        continue;
                    }

                    if (date.Date < recentCutoff.Date)
                    {
                        result.Add(dateText);
                        continue;
                    }

                    string cachedAtText = reader.IsDBNull(1) ? string.Empty : reader.GetString(1);
                    if (DateTime.TryParseExact(
                            cachedAtText,
                            "yyyy-MM-dd HH:mm:ss",
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None,
                            out DateTime cachedAt) &&
                        DateTime.Now - cachedAt <= RecentSourceCacheTtl)
                    {
                        result.Add(dateText);
                    }
                }
            }
            catch
            {
                // A missing/corrupt cache marker simply falls back to BMES.
            }
        }
        return result;
    }

    private static DateTime? TryParseDailyDbDate(string path)
    {
        string name = Path.GetFileNameWithoutExtension(path);
        return DateTime.TryParseExact(name, "yyyyMMdd", CultureInfo.InvariantCulture,
            DateTimeStyles.None, out var date)
            ? date.Date
            : null;
    }

    private static HashSet<string> GetMonthlyCachedDates(string dbPath)
    {
        var dates = new HashSet<string>(StringComparer.Ordinal);
        if (!File.Exists(dbPath))
            return dates;

        try
        {
            using var conn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
            conn.Open();

            if (TableExists(conn, "__NgRateCachedDates"))
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT [PRODUCT_DATE] FROM [__NgRateCachedDates]";
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                    dates.Add(reader.IsDBNull(0) ? string.Empty : reader.GetString(0));
                return dates;
            }

            if (TableExists(conn, "OrginalTable"))
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT DISTINCT [PRODUCT_DATE] FROM [OrginalTable] WHERE COALESCE([PRODUCT_DATE], '') <> ''";
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                    dates.Add(reader.IsDBNull(0) ? string.Empty : reader.GetString(0));
            }
        }
        catch { }

        return dates;
    }

    private List<Dictionary<string, string>> LoadFromMonthlyDb(
        DateTime month,
        IReadOnlyList<DateTime> dates,
        IReadOnlyCollection<string>? lineShiftFilter = null)
    {
        var rows = new List<Dictionary<string, string>>();
        if (dates.Count == 0)
            return rows;

        string dbPath = GetMonthlyDbPath(month);
        if (!File.Exists(dbPath))
            return rows;

        try
        {
            using var conn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
            conn.Open();

            var columns = GetOriginalTableColumns(conn);
            if (columns.Count == 0)
                return rows;

            using var cmd = conn.CreateCommand();
            var parameters = dates
                .Select((date, index) => (Date: date, Name: $"@d{index}"))
                .ToList();
            cmd.CommandText =
                $"SELECT * FROM [OrginalTable] WHERE [PRODUCT_DATE] IN ({string.Join(", ", parameters.Select(p => p.Name))})";
            foreach (var p in parameters)
                cmd.Parameters.AddWithValue(p.Name, p.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));

            HashSet<string>? normalizedLineShifts = NormalizeLineShiftFilter(lineShiftFilter);
            if (normalizedLineShifts is { Count: > 0 })
            {
                var lineShiftParameters = normalizedLineShifts
                    .OrderBy(value => value, StringComparer.Ordinal)
                    .Select((value, index) => (Value: value, Name: $"@ls{index}"))
                    .ToList();
                cmd.CommandText +=
                    $" AND ([MATERIALNAME] || '_' || [PRODUCTION_LINE]) IN " +
                    $"({string.Join(", ", lineShiftParameters.Select(p => p.Name))})";
                foreach (var p in lineShiftParameters)
                    cmd.Parameters.AddWithValue(p.Name, p.Value);
            }

            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var row = new Dictionary<string, string>(columns.Count, StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < columns.Count; i++)
                    row[columns[i]] = reader.IsDBNull(i) ? string.Empty : reader.GetValue(i).ToString()!;
                rows.Add(row);
            }
        }
        catch { }

        return rows;
    }

    private static HashSet<string>? NormalizeLineShiftFilter(
        IEnumerable<string>? lineShiftFilter)
    {
        if (lineShiftFilter is null)
            return null;

        var normalized = lineShiftFilter
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value.Trim())
            .ToHashSet(StringComparer.Ordinal);
        return normalized.Count == 0 ? null : normalized;
    }

    private static string BuildLineShiftFilterKey(IReadOnlyCollection<string>? lineShiftFilter)
        => lineShiftFilter is not { Count: > 0 }
            ? string.Empty
            : string.Join('\n', lineShiftFilter.OrderBy(value => value, StringComparer.Ordinal));

    private static List<Dictionary<string, string>> FilterRowsByLineShift(
        List<Dictionary<string, string>> rows,
        IReadOnlySet<string>? lineShiftFilter)
    {
        if (lineShiftFilter is not { Count: > 0 })
            return rows;

        return rows
            .Where(row => lineShiftFilter.Contains(
                GetCol(row, "MATERIALNAME") + "_" + GetCol(row, "PRODUCTION_LINE")))
            .ToList();
    }

    private static void EnsureMonthlyCacheSchema(SqliteConnection conn)
    {
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = $"CREATE TABLE IF NOT EXISTS [OrginalTable] " +
                $"({string.Join(", ", OrgTableColumns.Select(c => $"[{c}] TEXT"))})";
            cmd.ExecuteNonQuery();
        }

        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText =
                "CREATE INDEX IF NOT EXISTS [IX_OrginalTable_PRODUCT_DATE] " +
                "ON [OrginalTable]([PRODUCT_DATE])";
            cmd.ExecuteNonQuery();
        }

        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText =
                "CREATE TABLE IF NOT EXISTS [__NgRateCachedDates] " +
                "([PRODUCT_DATE] TEXT PRIMARY KEY, [RowCount] INTEGER NOT NULL, [CachedAt] TEXT NOT NULL)";
            cmd.ExecuteNonQuery();
        }
    }

    private static void ReplaceMonthlyDateRows(
        SqliteConnection conn,
        SqliteTransaction tx,
        DateTime date,
        List<Dictionary<string, string>> rows)
    {
        string dateText = date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

        using (var delete = conn.CreateCommand())
        {
            delete.Transaction = tx;
            delete.CommandText = "DELETE FROM [OrginalTable] WHERE [PRODUCT_DATE] = @date";
            delete.Parameters.AddWithValue("@date", dateText);
            delete.ExecuteNonQuery();
        }

        if (rows.Count > 0)
        {
            string colList = string.Join(", ", OrgTableColumns.Select(c => $"[{c}]"));
            string paramList = string.Join(", ", OrgTableColumns.Select((_, i) => $"@p{i}"));
            using var insert = conn.CreateCommand();
            insert.Transaction = tx;
            insert.CommandText = $"INSERT INTO [OrginalTable] ({colList}) VALUES ({paramList})";

            foreach (var row in rows)
            {
                insert.Parameters.Clear();
                for (int i = 0; i < OrgTableColumns.Length; i++)
                    insert.Parameters.AddWithValue(
                        $"@p{i}",
                        row.TryGetValue(OrgTableColumns[i], out var value) ? value : string.Empty);
                insert.ExecuteNonQuery();
            }
        }

        using (var marker = conn.CreateCommand())
        {
            marker.Transaction = tx;
            marker.CommandText =
                "INSERT INTO [__NgRateCachedDates] ([PRODUCT_DATE], [RowCount], [CachedAt]) VALUES (@date, @count, @cachedAt) " +
                "ON CONFLICT([PRODUCT_DATE]) DO UPDATE SET [RowCount] = excluded.[RowCount], [CachedAt] = excluded.[CachedAt]";
            marker.Parameters.AddWithValue("@date", dateText);
            marker.Parameters.AddWithValue("@count", rows.Count);
            marker.Parameters.AddWithValue("@cachedAt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
            marker.ExecuteNonQuery();
        }
    }

    private static bool TableExists(SqliteConnection conn, string tableName)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = @name LIMIT 1";
        cmd.Parameters.AddWithValue("@name", tableName);
        return cmd.ExecuteScalar() is not null;
    }

    private static List<string> GetOriginalTableColumns(SqliteConnection conn)
    {
        var columns = new List<string>();
        using var pragma = conn.CreateCommand();
        pragma.CommandText = "PRAGMA table_info([OrginalTable])";
        using var r = pragma.ExecuteReader();
        while (r.Read())
            columns.Add(r.GetString(1));
        return columns;
    }

    private static List<Dictionary<string, string>> LoadFromPerDayDb(string dbPath)
    {
        var rows = new List<Dictionary<string, string>>();
        try
        {
            using var conn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
            conn.Open();

            var columns = new List<string>();
            using (var pragma = conn.CreateCommand())
            {
                pragma.CommandText = "PRAGMA table_info([OrginalTable])";
                using var r = pragma.ExecuteReader();
                while (r.Read()) columns.Add(r.GetString(1));
            }
            if (columns.Count == 0) return rows;

            using var cmd    = conn.CreateCommand();
            cmd.CommandText  = "SELECT * FROM [OrginalTable]";
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var row = new Dictionary<string, string>(
                    columns.Count, StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < columns.Count; i++)
                    row[columns[i]] = reader.IsDBNull(i)
                        ? string.Empty : reader.GetValue(i).ToString()!;
                rows.Add(row);
            }
        }
        catch { }
        return rows;
    }

    private void CleanupTempDbs(IProgress<string>? progress)
    {
        try
        {
            string dir = _settings.DbSaveDirectory;
            if (!Directory.Exists(dir)) return;
            foreach (string f in Directory.GetFiles(dir, "temp_*.db"))
            {
                try
                {
                    File.Delete(f);
                    progress?.Report($"  Deleted old temp DB: {Path.GetFileName(f)}");
                }
                catch { }
            }
        }
        catch { }
    }

    // ── Private: HTTP ───────────────────────────────────────────────────────────

    private static async Task<string> GetTokenAsync(HttpClient client)
    {
        try
        {
            string html = await client.GetStringAsync(BaseUrl);
            // Also handle the "value before name" pattern
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

    private static async Task<BmesLoginResult> LoginAsync(
        HttpClient client, string token, string id, string password)
    {
        var content = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["UserInfo[USRID]"] = id,
            ["UserInfo[PWNO]"]  = password,
            ["UserInfo[LANG]"]  = "EN",
            ["UserInfo[FACCO]"] = "GN",
            ["UserInfo[STYPE]"] = "P",
            ["UserInfo[VTYPE]"] = "P",
            ["__RequestVerificationToken"] = token,
        });
        try
        {
            var response = await client.PostAsync(BaseUrl + "/MES000000/LoginCheck", content);
            string body  = await response.Content.ReadAsStringAsync();
            return new BmesLoginResult(
                body.Contains("\"Result\":\"M\""),
                BuildLoginDiagnostic(response, body));
        }
        catch (Exception ex) { return new BmesLoginResult(false, $"request error: {ex.GetType().Name}"); }
    }

    private static string BuildLoginDiagnostic(HttpResponseMessage response, string body)
    {
        string status = $"HTTP {(int)response.StatusCode} {response.ReasonPhrase}".Trim();
        Match result = Regex.Match(body, "\"Result\"\\s*:\\s*\"([^\"]*)\"", RegexOptions.IgnoreCase);
        Match message = Regex.Match(body, "\"(?:Message|ErrorMessage)\"\\s*:\\s*\"([^\"]*)\"", RegexOptions.IgnoreCase);
        if (result.Success) status += $", Result={result.Groups[1].Value}";
        if (message.Success) status += $", Message={SanitizeLoginMessage(message.Groups[1].Value)}";
        return status;
    }

    private static string SanitizeLoginMessage(string value)
    {
        string normalized = Regex.Replace(value, @"\s+", " ").Trim();
        return normalized.Length <= 160 ? normalized : normalized[..160] + "...";
    }

    private static async Task<List<Dictionary<string, string>>?> FetchRawRowsAsync(
        HttpClient client, string werks, string start, string end,
        IProgress<string>? progress)
    {
        string url = $"{BaseUrl}/MES020210/SearchList?perPage=" +
                     $"&Condition.WERKS={werks}&Condition.SDATE={start}" +
                     $"&Condition.EDATE={end}&Condition.INPYN=N&Condition.USEYN=Y&page=1";
        try
        {
            var response = await client.GetAsync(url);
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

                    string colName = ApiColumnMap.TryGetValue(prop.Name, out string? mapped)
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

    // ── Private: data processing ─────────────────────────────────────────────────

    /// <summary>
    /// Same behavior as the WPF clMakeProcTable.SelectRowsForProcTable path.
    /// Within each (LINE, CODE, PROCESSNAME, NGNAME, MATERIAL, DATE, SHIFT) group,
    /// compare QTYINPUT / QTYNG and auto-merge / auto-pick rows.
    /// </summary>
    private static List<Dictionary<string, string>> RemoveDuplicates(
        List<Dictionary<string, string>> rows)
    {
        var result = new List<Dictionary<string, string>>(rows.Count);

        var groups = rows.GroupBy(row =>
            $"{GetCol(row, "PRODUCTION_LINE")}|" +
            $"{GetCol(row, "PROCESSCODE")}|" +
            $"{NormalizeText(GetCol(row, "PROCESSNAME"))}|" +
            $"{NormalizeText(GetCol(row, "NGNAME"))}|" +
            $"{GetCol(row, "MATERIALNAME")}|" +
            $"{GetCol(row, "PRODUCT_DATE")}|" +
            $"{GetCol(row, "Shift")}",
            StringComparer.Ordinal);

        foreach (var group in groups)
        {
            var groupRows = group.ToList();
            if (groupRows.Count == 1)
            {
                result.Add(groupRows[0]);
                continue;
            }

            // Iterate from the last row and compare/merge with preceding rows in order (same as WPF)
            var selected = groupRows[groupRows.Count - 1];
            for (int i = groupRows.Count - 2; i >= 0; i--)
                selected = ResolveRow(groupRows[i], selected);

            result.Add(selected);
        }

        return result;
    }

    /// <summary>
    /// Auto-resolution rules mirroring WPF TryResolveRowsWithoutPrompt (no prompting):
    /// 1. Values identical           → keep B
    /// 2. QTYINPUT differs           → keep the larger QTYINPUT (WPF asks user; web auto)
    /// 3. QTYINPUT equal, one QTYNG=0 → keep the non-zero one
    /// 4. QTYINPUT equal, both 0     → keep B (WPF asks user; web auto)
    /// 5. QTYINPUT equal, both non-zero → merge by summing QTYNG
    /// </summary>
    private static Dictionary<string, string> ResolveRow(
        Dictionary<string, string> optionA, Dictionary<string, string> optionB)
    {
        double inputA = ParseDouble(GetCol(optionA, "QTYINPUT"));
        double inputB = ParseDouble(GetCol(optionB, "QTYINPUT"));
        double ngA    = ParseDouble(GetCol(optionA, "QTYNG"));
        double ngB    = ParseDouble(GetCol(optionB, "QTYNG"));

        // 1. Identical values
        if (inputA == inputB && ngA == ngB)
            return optionB;

        // 2. QTYINPUT differs → pick the larger
        if (inputA != inputB)
            return inputA > inputB ? optionA : optionB;

        // QTYINPUT equal, QTYNG differs
        bool aZero = ngA == 0;
        bool bZero = ngB == 0;

        // 3. One side is 0 → pick the non-zero one
        if (aZero != bZero)
            return aZero ? optionB : optionA;

        // 4. Both zero
        if (aZero)
            return optionB;

        // 5. Both non-zero → merge by summing QTYNG
        var merged = new Dictionary<string, string>(optionB, StringComparer.OrdinalIgnoreCase);
        merged["QTYNG"] = (ngA + ngB).ToString(CultureInfo.InvariantCulture);
        return merged;
    }

    private static double ParseDouble(string s)
        => double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double v) ? v : 0;

    private static string GetCol(Dictionary<string, string> row, string key)
        => row.TryGetValue(key, out var v) ? v.Trim() : string.Empty;

    private static string NormalizeText(string input)
    {
        if (string.IsNullOrWhiteSpace(input)) return string.Empty;
        input = input.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
        input = input.Replace("\u2018", "'").Replace("\u2019", "'")
                     .Replace("\u201c", "\"").Replace("\u201d", "\"");
        input = input.Replace("'", " ").Replace("\"", " ").Replace("~", " ");
        input = input.Replace("[", "").Replace("]", "_").Replace("+", " ");
        input = Regex.Replace(input, @"\s{2,}", " ");
        return input.Trim();
    }

    // ── Private: SQLite ──────────────────────────────────────────────────────────

    private static void SaveToSqlite(string dbPath, List<Dictionary<string, string>> rows)
    {
        // Detect which keys are actually present and order them by OrgTableColumns
        var dataKeys = rows.SelectMany(r => r.Keys)
                           .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var columns = OrgTableColumns
            .Where(c => dataKeys.Contains(c))
            .ToList();

        // Append any remaining columns not listed in OrgTableColumns
        foreach (string k in dataKeys)
        {
            if (!columns.Any(c => c.Equals(k, StringComparison.OrdinalIgnoreCase)))
                columns.Add(k);
        }

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        // Create table
        string colDefs = string.Join(", ", columns.Select(c => $"[{c}] TEXT"));
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = $"CREATE TABLE IF NOT EXISTS [OrginalTable] ({colDefs})";
            cmd.ExecuteNonQuery();
        }

        // Insert rows (transactional)
        string colList  = string.Join(", ", columns.Select(c => $"[{c}]"));
        string paramList = string.Join(", ", columns.Select((_, i) => $"@p{i}"));
        string insertSql = $"INSERT INTO [OrginalTable] ({colList}) VALUES ({paramList})";

        using var tx        = conn.BeginTransaction();
        using var insertCmd = conn.CreateCommand();
        insertCmd.CommandText = insertSql;

        foreach (var row in rows)
        {
            insertCmd.Parameters.Clear();
            for (int i = 0; i < columns.Count; i++)
            {
                insertCmd.Parameters.AddWithValue(
                    $"@p{i}",
                    row.TryGetValue(columns[i], out var v) ? v : string.Empty);
            }
            insertCmd.ExecuteNonQuery();
        }
        tx.Commit();

        // Persist meta table
        SaveMeta(conn);
    }

    private static void SaveMeta(SqliteConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "CREATE TABLE IF NOT EXISTS [__DataMakerMeta] " +
            "([MetaKey] TEXT PRIMARY KEY, [MetaValue] TEXT NOT NULL)";
        cmd.ExecuteNonQuery();

        cmd.CommandText =
            "INSERT INTO [__DataMakerMeta] ([MetaKey], [MetaValue]) VALUES (@key, @val) " +
            "ON CONFLICT([MetaKey]) DO UPDATE SET [MetaValue] = excluded.[MetaValue]";
        cmd.Parameters.AddWithValue("@key", "OriginalTableUpdatedAt");
        cmd.Parameters.AddWithValue("@val", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        cmd.ExecuteNonQuery();
    }

    // ── Private: Processing ──────────────────────────────────────────────────────

    /// <summary>
    /// Same flow as DataMaker clDataProcessor.ProcessData():
    ///   1. Load RoutingTable (Routing.txt)
    ///   2. Load ReasonTable  (reason.txt)
    ///   3. Populate OrginalTable's LineShift column
    /// </summary>
    private void ProcessData(string dbPath, IProgress<string>? progress)
    {
        var swTotal = System.Diagnostics.Stopwatch.StartNew();
        Log(progress, $"[ProcessData] Open SQLite: {Path.GetFileName(dbPath)}");
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        Log(progress, $"[ProcessData] DB opened ({swTotal.ElapsedMilliseconds} ms)");

        // ── Routing ──────────────────────────────────────────────────────────
        var swSection = System.Diagnostics.Stopwatch.StartNew();
        var routingDbRows = _settings.GetRoutingRows();
        Log(progress, $"[ProcessData] Routing source = settings DB rows={routingDbRows.Count} ({swSection.ElapsedMilliseconds} ms)");
        if (routingDbRows.Count > 0)
        {
            swSection.Restart();
            Log(progress, $"Loading Routing table from settings DB ({routingDbRows.Count} rows)…");
            LoadRoutingFromSettings(conn, routingDbRows);
            Log(progress, $"[ProcessData]   LoadRoutingFromSettings done ({swSection.ElapsedMilliseconds} ms)");
            swSection.Restart();
            NormalizeTableColumns(conn, "RoutingTable", new[] { "ProcessName", "ProcessType" });
            Log(progress, $"[ProcessData]   NormalizeTableColumns(RoutingTable) done ({swSection.ElapsedMilliseconds} ms)");
            Log(progress, "Routing table done.");
        }
        else if (File.Exists(_settings.RoutingFilePath))
        {
            swSection.Restart();
            Log(progress, "Loading Routing table from file…");
            LoadTsvToTable(conn, _settings.RoutingFilePath, "RoutingTable",
                new[] { "모델명", "ProcessCode", "ProcessName", "ProcessType" });
            Log(progress, $"[ProcessData]   LoadTsvToTable(Routing) done ({swSection.ElapsedMilliseconds} ms)");
            swSection.Restart();
            NormalizeTableColumns(conn, "RoutingTable", new[] { "ProcessName", "ProcessType" });
            Log(progress, $"[ProcessData]   NormalizeTableColumns(RoutingTable) done ({swSection.ElapsedMilliseconds} ms)");
            Log(progress, "Routing table done.");
        }
        else
        {
            Log(progress, "[WARN] No Routing data (settings DB empty, file not found).");
        }

        // ── Reason ───────────────────────────────────────────────────────────
        swSection.Restart();
        var reasonDbRows = _settings.GetReasonRows();
        Log(progress, $"[ProcessData] Reason source = settings DB rows={reasonDbRows.Count} ({swSection.ElapsedMilliseconds} ms)");
        if (reasonDbRows.Count > 0)
        {
            swSection.Restart();
            Log(progress, $"Loading Reason table from settings DB ({reasonDbRows.Count} rows)…");
            LoadReasonFromSettings(conn, reasonDbRows);
            Log(progress, $"[ProcessData]   LoadReasonFromSettings done ({swSection.ElapsedMilliseconds} ms)");
            swSection.Restart();
            NormalizeTableColumns(conn, "Reason", new[] { "processName", "NgName" });
            Log(progress, $"[ProcessData]   NormalizeTableColumns(Reason) done ({swSection.ElapsedMilliseconds} ms)");
            Log(progress, "Reason table done.");
        }
        else if (File.Exists(_settings.ReasonFilePath))
        {
            swSection.Restart();
            Log(progress, "Loading Reason table from file…");
            LoadTsvToTable(conn, _settings.ReasonFilePath, "Reason",
                new[] { "processName", "NgName", "Reason" });
            Log(progress, $"[ProcessData]   LoadTsvToTable(Reason) done ({swSection.ElapsedMilliseconds} ms)");
            swSection.Restart();
            NormalizeTableColumns(conn, "Reason", new[] { "processName", "NgName" });
            Log(progress, $"[ProcessData]   NormalizeTableColumns(Reason) done ({swSection.ElapsedMilliseconds} ms)");
            Log(progress, "Reason table done.");
        }
        else
        {
            Log(progress, "[WARN] No Reason data (settings DB empty, file not found).");
        }

        // ── LineShift column population ──────────────────────────────────────
        swSection.Restart();
        Log(progress, "Setting LineShift column…");
        SetLineShift(conn);
        Log(progress, $"[ProcessData] SetLineShift done ({swSection.ElapsedMilliseconds} ms)");
        Log(progress, $"[ProcessData] Processing complete (total {swTotal.ElapsedMilliseconds} ms)");
    }

    /// <summary>
    /// Tab-separated text file → SQLite table (skip header row; use Columns parameter).
    /// Same logic as clMakeTxtTable.MakeDataTable.
    /// </summary>
    private static void LoadRoutingFromSettings(SqliteConnection conn, List<RoutingRow> rows)
    {
        using (var drop = conn.CreateCommand()) { drop.CommandText = "DROP TABLE IF EXISTS [RoutingTable]"; drop.ExecuteNonQuery(); }
        using (var create = conn.CreateCommand())
        {
            create.CommandText = "CREATE TABLE [RoutingTable] ([모델명] TEXT, [ProcessCode] TEXT, [ProcessName] TEXT, [ProcessType] TEXT)";
            create.ExecuteNonQuery();
        }
        using var tx = conn.BeginTransaction();
        foreach (var r in rows)
        {
            using var ins = conn.CreateCommand();
            ins.CommandText = "INSERT INTO [RoutingTable] ([모델명],[ProcessCode],[ProcessName],[ProcessType]) VALUES (@m,@pc,@pn,@pt)";
            ins.Parameters.AddWithValue("@m",  r.ModelName);
            ins.Parameters.AddWithValue("@pc", r.ProcessCode);
            ins.Parameters.AddWithValue("@pn", r.ProcessName);
            ins.Parameters.AddWithValue("@pt", r.ProcessType);
            ins.ExecuteNonQuery();
        }
        tx.Commit();
    }

    private static void LoadReasonFromSettings(SqliteConnection conn, List<ReasonRow> rows)
    {
        using (var drop = conn.CreateCommand()) { drop.CommandText = "DROP TABLE IF EXISTS [Reason]"; drop.ExecuteNonQuery(); }
        using (var create = conn.CreateCommand())
        {
            create.CommandText = "CREATE TABLE [Reason] ([processName] TEXT, [NgName] TEXT, [Reason] TEXT)";
            create.ExecuteNonQuery();
        }
        using var tx = conn.BeginTransaction();
        foreach (var r in rows)
        {
            using var ins = conn.CreateCommand();
            ins.CommandText = "INSERT INTO [Reason] ([processName],[NgName],[Reason]) VALUES (@pn,@ng,@rs)";
            ins.Parameters.AddWithValue("@pn", r.ProcessName);
            ins.Parameters.AddWithValue("@ng", r.NgName);
            ins.Parameters.AddWithValue("@rs", r.Reason);
            ins.ExecuteNonQuery();
        }
        tx.Commit();
    }

    private static void LoadTsvToTable(
        SqliteConnection conn, string filePath,
        string tableName, string[] columns)
    {
        // Drop existing
        using (var drop = conn.CreateCommand())
        {
            drop.CommandText = $"DROP TABLE IF EXISTS [{tableName}]";
            drop.ExecuteNonQuery();
        }

        // Read file
        string[] lines = File.ReadAllLines(filePath, Encoding.UTF8);
        if (lines.Length < 2) return; // header-only or empty file

        // Create table
        string colDefs = string.Join(", ", columns.Select(c => $"[{c}] TEXT"));
        using (var create = conn.CreateCommand())
        {
            create.CommandText = $"CREATE TABLE [{tableName}] ({colDefs})";
            create.ExecuteNonQuery();
        }

        // Insert (skip first line = header)
        string colList  = string.Join(", ", columns.Select(c => $"[{c}]"));
        string paramList = string.Join(", ", columns.Select((_, i) => $"@p{i}"));

        using var tx  = conn.BeginTransaction();
        using var ins = conn.CreateCommand();
        ins.CommandText = $"INSERT INTO [{tableName}] ({colList}) VALUES ({paramList})";

        for (int i = 1; i < lines.Length; i++)
        {
            if (string.IsNullOrWhiteSpace(lines[i])) continue;
            string[] parts = lines[i].Split('\t');
            ins.Parameters.Clear();
            for (int j = 0; j < columns.Length; j++)
                ins.Parameters.AddWithValue($"@p{j}",
                    j < parts.Length ? parts[j].Trim() : string.Empty);
            ins.ExecuteNonQuery();
        }
        tx.Commit();
    }

    /// <summary>
    /// Apply the same normalization as CONSTANT.Normalize to the given columns and persist.
    /// </summary>
    private static void NormalizeTableColumns(
        SqliteConnection conn, string tableName, string[] columnNames)
    {
        // Normalize column-by-column via UPDATE (SQLite has no regex func; load + save approach).
        // Use only for small master tables.
        using var selCmd = conn.CreateCommand();
        selCmd.CommandText = $"SELECT rowid, {string.Join(", ", columnNames.Select(c => $"[{c}]"))} FROM [{tableName}]";

        var updates = new List<(long Rowid, string[] Values)>();
        using (var reader = selCmd.ExecuteReader())
        {
            while (reader.Read())
            {
                long rowid  = reader.GetInt64(0);
                var  values = new string[columnNames.Length];
                for (int i = 0; i < columnNames.Length; i++)
                    values[i] = NormalizeText(reader.IsDBNull(i + 1) ? string.Empty : reader.GetString(i + 1));
                updates.Add((rowid, values));
            }
        }

        string setClause = string.Join(", ",
            columnNames.Select((c, i) => $"[{c}] = @v{i}"));

        using var tx  = conn.BeginTransaction();
        using var upd = conn.CreateCommand();
        upd.CommandText = $"UPDATE [{tableName}] SET {setClause} WHERE rowid = @rowid";

        foreach (var (rowid, values) in updates)
        {
            upd.Parameters.Clear();
            upd.Parameters.AddWithValue("@rowid", rowid);
            for (int i = 0; i < values.Length; i++)
                upd.Parameters.AddWithValue($"@v{i}", values[i]);
            upd.ExecuteNonQuery();
        }
        tx.Commit();
    }

    private static void SetLineShift(SqliteConnection conn)
    {
        // ADD COLUMN (ignore if it already exists)
        try
        {
            using var alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE [OrginalTable] ADD COLUMN [LineShift] TEXT";
            alter.ExecuteNonQuery();
        }
        catch { /* already exists */ }

        using var upd = conn.CreateCommand();
        upd.CommandText =
            "UPDATE [OrginalTable] SET [LineShift] = MATERIALNAME || '_' || PRODUCTION_LINE";
        upd.ExecuteNonQuery();
    }
}
