using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace PortalGNLAB.Web.Services;

/// <summary>
/// BMES NG Rate data collection → in-memory data set (NgRateDataStore) service.
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
    private static readonly object ReusableDbLock = new();
    private static ReusableDb? _lastReusableDb;
    private static readonly TimeSpan RecentRangeReuseTtl = TimeSpan.FromMinutes(15);
    private static readonly TimeSpan RecentSourceCacheTtl = TimeSpan.FromMinutes(15);
    // Dates within this many days of today may still change on BMES, so their cache
    // is only trusted for RecentSourceCacheTtl. A snapshot taken while a date was
    // still inside this window is never treated as final, even after the date ages out.
    private const int RecentWindowDays = 2;

    private sealed record ReusableDb(
        DateTime StartDate,
        DateTime EndDate,
        string LineShiftFilterKey,
        DateTimeOffset CachedAt,
        string Handle);

    /// <summary>Mirrors a progress message to both the UI and the Debug/launching-console output.
    /// Use this for every status line so you can diagnose stalls live in the dotnet console
    /// and the VS "Debug" output simultaneously.</summary>
    private void Log(IProgress<string>? progress, string msg)
    {
        _activity.Log("NgRate", msg);
        progress?.Report(msg);
    }

    private static string Rows(List<Dictionary<string, string>>? rows) =>
        rows is null ? "failed" : $"{rows.Count:N0} rows";

    // OrginalTable column order (per DataMaker CONSTANT.ListSTRManager); also the NgRateCache layout
    private static readonly string[] OrgTableColumns = NgRateCacheSchema.Columns;

    // ── Public API ──────────────────────────────────────────────────────────────

    /// <summary>
    /// Full flow: BMES data fetch → merged in-memory data set.
    /// - Dates 3+ days ago : use monthly DB cache (fetch + cache if missing)
    /// - Dates 0–2 days ago: always fetch from server (data may still change)
    /// - Returns the <see cref="NgRateDataSet.Handle"/> of the merged rows in
    ///   <see cref="NgRateDataStore"/>; nothing is written to disk besides the NgRateCache.
    /// </summary>
    public async Task<string?> GetOrFetchAsync(
        DateTime startDate,
        DateTime endDate,
        IProgress<string>? progress = null,
        IReadOnlyCollection<string>? lineShiftFilter = null)
    {
        HashSet<string>? normalizedLineShifts = NormalizeLineShiftFilter(lineShiftFilter);
        string filterKey = BuildLineShiftFilterKey(normalizedLineShifts);
        if (TryGetReusableDb(startDate, endDate, filterKey, out string? handle))
        {
            Log(progress, $"Reusing fetched data: {handle}");
            return handle;
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
        var recentCutoff = DateTime.Today.AddDays(-RecentWindowDays);
        var allDates = Enumerable
            .Range(0, (int)(endDate.Date - startDate.Date).TotalDays + 1)
            .Select(i => startDate.Date.AddDays(i))
            .ToList();

        var toFetch = new List<DateTime>(); // needs BMES server fetch
        var toCache = new List<DateTime>(); // load from monthly DB cache

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

            // NG records come only from PLATFORM_API; the BMES web login is just the fallback API login.
            BmesLoginOutcome login = await BmesWebSession.OpenAsync(_settings.LoginId, _settings.Password, progress: progress);
            if (!login.Succeeded)
            {
                progress?.Report("[ERROR] BMES API login not configured. Ask admin to set it in Setting.");
                return null;
            }
            using BmesWebSession session = login.Session!;
            Log(progress, "Source: PLATFORM_API AI_TOOL_00055 view S (MES020210), plants 3200 / 3220.");

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
                var fetch3200 = BmesGateway.FetchNgRecordsAsync(session, "3200", s, e, progress);
                var fetch3220 = BmesGateway.FetchNgRecordsAsync(session, "3220", s, e, progress);
                await Task.WhenAll(fetch3200, fetch3220);
                var rows3200 = await fetch3200;
                var rows3220 = await fetch3220;
                if (rows3200 != null) serverRows.AddRange(rows3200);
                if (rows3220 != null) serverRows.AddRange(rows3220);
                Log(progress, $"  API {s} ~ {e}: 3200 {Rows(rows3200)}, 3220 {Rows(rows3220)}");

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

        // ── 5. Routing / Reason tables, frozen with the rows ────────────────
        await Task.Run(() => DeleteLeftoverTempDbs(progress));
        await EnsureServerTablesAsync(progress);
        var swBuild = System.Diagnostics.Stopwatch.StartNew();
        var routing = LoadRoutingEntries(progress);
        var reasons = LoadReasonEntries(progress);

        // ── 6. Keep the merged rows in memory ───────────────────────────────
        NgRateDataSet data = await Task.Run(() => NgRateDataSet.Create(allRows, routing, reasons));
        NgRateDataStore.Add(data);
        Log(progress, $"NG data ready in memory: {data.Handle} ({data.Rows.Count:N0} rows, " +
                      $"routing {routing.Count:N0}, reason {reasons.Count:N0}, {swBuild.ElapsedMilliseconds} ms)");

        RememberReusableDb(startDate, endDate, filterKey, data.Handle);
        Log(progress, $"Done. Data: {data.Handle}  (total {swStart.ElapsedMilliseconds} ms since start)");
        return data.Handle;
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
                && (endDate.Date < DateTime.Today.AddDays(-RecentWindowDays) ||
                    DateTimeOffset.Now - cached.CachedAt <= RecentRangeReuseTtl)
                && NgRateDataStore.Contains(cached.Handle))
            {
                dbPath = cached.Handle;
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
        string handle)
    {
        lock (ReusableDbLock)
        {
            _lastReusableDb = new ReusableDb(
                startDate.Date,
                endDate.Date,
                lineShiftFilterKey,
                DateTimeOffset.Now,
                handle);
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

    // ── Monthly DB helpers ───────────────────────────────────────────────────

    private void SaveMonthlyDateDb(
        DateTime date, List<Dictionary<string, string>> rows, IProgress<string>? progress)
    {
        try
        {
            using var conn = PortalDb.Open();
            NgRateCacheSchema.Ensure(conn);
            using var tx = conn.BeginTransaction();
            ReplaceMonthlyDateRows(conn, tx, date, rows);
            tx.Commit();

            progress?.Report(rows.Count == 0
                ? $"  Cached {date:MM/dd}: no data → {NgRateCacheSchema.RowsTable}"
                : $"  Cached {date:MM/dd}: {rows.Count:N0} rows → {NgRateCacheSchema.RowsTable}");
        }
        catch (Exception ex)
        {
            progress?.Report($"[WARN] Failed to cache {date:MM/dd}: {ex.Message}");
        }
    }

    private HashSet<string> GetMonthlyReusableCachedDateSet(
        IReadOnlyList<DateTime> dates,
        DateTime recentCutoff)
    {
        var result = new HashSet<string>(StringComparer.Ordinal);
        if (dates.Count == 0)
            return result;

        var wanted = dates
            .Select(d => d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
            .ToHashSet(StringComparer.Ordinal);
        string min = wanted.Min(StringComparer.Ordinal)!;
        string max = wanted.Max(StringComparer.Ordinal)!;

        try
        {
            using var conn = PortalDb.Open();
            NgRateCacheSchema.Ensure(conn);

            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                $"SELECT [PRODUCT_DATE], [CachedAt] FROM [{NgRateCacheSchema.DatesTable}] " +
                "WHERE [PRODUCT_DATE] BETWEEN @min AND @max";
            cmd.Parameters.AddWithValue("@min", min);
            cmd.Parameters.AddWithValue("@max", max);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                string dateText = reader.IsDBNull(0) ? string.Empty : reader.GetString(0);
                if (!wanted.Contains(dateText))
                    continue;
                if (!DateTime.TryParseExact(
                        dateText,
                        "yyyy-MM-dd",
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out DateTime date))
                {
                    continue;
                }

                string cachedAtText = reader.IsDBNull(1) ? string.Empty : reader.GetString(1);
                if (!DateTime.TryParseExact(
                        cachedAtText,
                        "yyyy-MM-dd HH:mm:ss",
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out DateTime cachedAt))
                {
                    continue;
                }

                // Final only when the snapshot was taken after the date left the
                // recent window. A marker written on the date itself (e.g. a morning
                // run before quantities were entered) or written ahead of the date
                // (a range that reached into the future) must be fetched once more
                // the first time the date is requested as settled history.
                bool settledSnapshot = cachedAt.Date > date.Date.AddDays(RecentWindowDays);
                if (date.Date < recentCutoff.Date && settledSnapshot)
                {
                    result.Add(dateText);
                    continue;
                }

                if (DateTime.Now - cachedAt <= RecentSourceCacheTtl)
                {
                    result.Add(dateText);
                }
            }
        }
        catch
        {
            // A missing/corrupt cache marker simply falls back to BMES.
        }
        return result;
    }

    private List<Dictionary<string, string>> LoadFromMonthlyDb(
        DateTime month,
        IReadOnlyList<DateTime> dates,
        IReadOnlyCollection<string>? lineShiftFilter = null)
    {
        var rows = new List<Dictionary<string, string>>();
        if (dates.Count == 0)
            return rows;

        try
        {
            using var conn = PortalDb.Open();
            NgRateCacheSchema.Ensure(conn);

            using var cmd = conn.CreateCommand();
            var parameters = dates
                .Select((date, index) => (Date: date, Name: $"@d{index}"))
                .ToList();
            cmd.CommandText =
                $"SELECT {string.Join(", ", OrgTableColumns.Select(c => $"[{c}]"))} FROM [{NgRateCacheSchema.RowsTable}] " +
                $"WHERE [PRODUCT_DATE] IN ({string.Join(", ", parameters.Select(p => p.Name))})";
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
                var row = new Dictionary<string, string>(OrgTableColumns.Length, StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < OrgTableColumns.Length; i++)
                    row[OrgTableColumns[i]] = reader.IsDBNull(i) ? string.Empty : reader.GetValue(i).ToString()!;
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
            delete.CommandText = $"DELETE FROM [{NgRateCacheSchema.RowsTable}] WHERE [PRODUCT_DATE] = @date";
            delete.Parameters.AddWithValue("@date", dateText);
            delete.ExecuteNonQuery();
        }

        if (rows.Count > 0)
        {
            string colList = string.Join(", ", OrgTableColumns.Select(c => $"[{c}]"));
            string paramList = string.Join(", ", OrgTableColumns.Select((_, i) => $"@p{i}"));
            using var insert = conn.CreateCommand();
            insert.Transaction = tx;
            insert.CommandText = $"INSERT INTO [{NgRateCacheSchema.RowsTable}] ({colList}) VALUES ({paramList})";

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
                $"INSERT INTO [{NgRateCacheSchema.DatesTable}] ([PRODUCT_DATE], [RowCount], [CachedAt]) VALUES (@date, @count, @cachedAt) " +
                "ON CONFLICT([PRODUCT_DATE]) DO UPDATE SET [RowCount] = excluded.[RowCount], [CachedAt] = excluded.[CachedAt]";
            marker.Parameters.AddWithValue("@date", dateText);
            marker.Parameters.AddWithValue("@count", rows.Count);
            marker.Parameters.AddWithValue("@cachedAt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
            marker.ExecuteNonQuery();
        }
    }

    /// <summary>Removes temp_*.db files the old flow left in DbSaveDirectory (one per fetch).</summary>
    private void DeleteLeftoverTempDbs(IProgress<string>? progress)
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
                    progress?.Report($"  Deleted leftover temp DB: {Path.GetFileName(f)}");
                }
                catch { }
            }
        }
        catch { }
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

    // ── Private: Routing / Reason ───────────────────────────────────────────────

    /// <summary>Routing rows for the report: the settings DB rows, else the Routing file
    /// (tab-separated, header row skipped). ProcessName / ProcessType are normalized,
    /// as the old temp DB RoutingTable was.</summary>
    private List<NgRateRoutingEntry> LoadRoutingEntries(IProgress<string>? progress)
    {
        var settingsRows = _settings.GetRoutingRows();
        if (settingsRows.Count > 0)
        {
            Log(progress, $"Routing table from settings DB ({settingsRows.Count} rows).");
            return settingsRows
                .Select(r => new NgRateRoutingEntry(
                    r.ModelName ?? string.Empty, r.ProcessCode ?? string.Empty,
                    NormalizeText(r.ProcessName ?? string.Empty), NormalizeText(r.ProcessType ?? string.Empty)))
                .ToList();
        }
        if (File.Exists(_settings.RoutingFilePath))
        {
            Log(progress, "Routing table from file.");
            return ReadTsvRows(_settings.RoutingFilePath, 4)
                .Select(c => new NgRateRoutingEntry(c[0], c[1], NormalizeText(c[2]), NormalizeText(c[3])))
                .ToList();
        }
        Log(progress, "[WARN] No Routing data (settings DB empty, file not found).");
        return [];
    }

    /// <summary>Reason rows for the report: the settings DB rows, else the Reason file.
    /// ProcessName / NgName are normalized, as the old temp DB Reason table was.</summary>
    private List<NgRateReasonEntry> LoadReasonEntries(IProgress<string>? progress)
    {
        var settingsRows = _settings.GetReasonRows();
        if (settingsRows.Count > 0)
        {
            Log(progress, $"Reason table from settings DB ({settingsRows.Count} rows).");
            return settingsRows
                .Select(r => new NgRateReasonEntry(
                    NormalizeText(r.ProcessName ?? string.Empty), NormalizeText(r.NgName ?? string.Empty),
                    r.Reason ?? string.Empty))
                .ToList();
        }
        if (File.Exists(_settings.ReasonFilePath))
        {
            Log(progress, "Reason table from file.");
            return ReadTsvRows(_settings.ReasonFilePath, 3)
                .Select(c => new NgRateReasonEntry(NormalizeText(c[0]), NormalizeText(c[1]), c[2]))
                .ToList();
        }
        Log(progress, "[WARN] No Reason data (settings DB empty, file not found).");
        return [];
    }

    /// <summary>Tab-separated text file → rows of <paramref name="columnCount"/> trimmed cells
    /// (header row and blank lines skipped, missing cells empty). Same reading as the WPF
    /// clMakeTxtTable.MakeDataTable.</summary>
    private static List<string[]> ReadTsvRows(string filePath, int columnCount)
    {
        var rows = new List<string[]>();
        string[] lines = File.ReadAllLines(filePath, Encoding.UTF8);
        for (int i = 1; i < lines.Length; i++)
        {
            if (string.IsNullOrWhiteSpace(lines[i])) continue;
            string[] parts = lines[i].Split('\t');
            var cells = new string[columnCount];
            for (int j = 0; j < columnCount; j++)
                cells[j] = j < parts.Length ? parts[j].Trim() : string.Empty;
            rows.Add(cells);
        }
        return rows;
    }
}
