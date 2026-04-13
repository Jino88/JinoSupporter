using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

/// <summary>
/// BMES NG Rate 데이터 수집 → SQLite 저장 → Processing 서비스
/// DataMaker WPF 앱의 clFetchBMES + clFetchBMESNGDATA + clDataProcessor 로직을 웹 환경에 포팅
/// 경로/자격증명은 NgRateSettingsService에서 읽음
/// </summary>
public sealed class NgRateService(NgRateSettingsService settings)
{
    private readonly NgRateSettingsService _settings = settings;
    private const string BaseUrl = "http://bmes.bujeon.com";

    // BMES API 컬럼명 → 내부 컬럼명 (DataMaker CONSTANT._columnMap + ListSTRManager 통합)
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

    // OrginalTable 컬럼 순서 (DataMaker CONSTANT.ListSTRManager 기준)
    private static readonly string[] OrgTableColumns =
    {
        "PRODUCTION_LINE", "PROCESSCODE", "PROCESSNAME", "NGCODE", "NGNAME",
        "USE", "QTYINPUT", "QTYNG", "INPUTQUANTITY", "NGQUANTITY",
        "MATERIALCODE", "MATERIALNAME", "PLANT", "WORKORDER", "Shift",
        "PRODUCT_DATE", "WORKODER", "ACTIVITYNO", "LASTREGPERSON", "LASTREGDATE",
    };

    // ── Public API ──────────────────────────────────────────────────────────────

    /// <summary>
    /// BMES 데이터 수집 → SQLite 저장 → Processing 전체 흐름 실행.
    /// ─ 3일 이상 지난 날짜: per-day DB 캐시 사용 (없으면 서버 조회 후 캐시 저장)
    /// ─ 0~2일 전 날짜:      항상 서버에서 조회 (데이터 변경 가능)
    /// ─ 결과는 temp DB (temp_*.db) 에 병합, 이전 temp DB는 자동 삭제
    /// </summary>
    public async Task<string?> FetchAndSaveAsync(
        DateTime startDate,
        DateTime endDate,
        IProgress<string>? progress = null)
    {
        // ── 1. 날짜 분류 ──────────────────────────────────────────────────────
        // · 오늘 / 어제  → 항상 서버 조회 (데이터 변경 가능)
        // · 그 외        → per-day DB 파일 있으면 캐시, 없으면 서버 조회
        var recentCutoff = DateTime.Today.AddDays(-2); // 이 날짜 이상이면 항상 서버 조회 (오늘·어제·2일전)
        var allDates = Enumerable
            .Range(0, (int)(endDate.Date - startDate.Date).TotalDays + 1)
            .Select(i => startDate.Date.AddDays(i))
            .ToList();

        var toFetch = new List<DateTime>(); // BMES 서버 조회 필요
        var toCache = new List<DateTime>(); // per-day DB 캐시 사용

        foreach (var date in allDates)
        {
            if (date >= recentCutoff || !File.Exists(GetPerDayDbPath(date)))
                toFetch.Add(date);
            else
                toCache.Add(date);
        }

        progress?.Report(
            $"날짜 범위: {startDate:MM/dd} – {endDate:MM/dd}  " +
            $"(서버: {toFetch.Count}일 / 캐시: {toCache.Count}일)");

        // ── 2. BMES 서버 조회 ────────────────────────────────────────────────
        var freshRows = new List<Dictionary<string, string>>();

        if (toFetch.Count > 0)
        {
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
            if (!await LoginAsync(client, token, loginId, password))
            {
                progress?.Report("[ERROR] Login failed — check credentials in NG Rate Settings.");
                return null;
            }
            progress?.Report("Login successful.");

            // toFetch 날짜 중 최소~최대 범위로 한 번에 요청
            string fetchStart = toFetch.Min().ToString("yyyy-MM-dd");
            string fetchEnd   = toFetch.Max().ToString("yyyy-MM-dd");

            progress?.Report($"Fetching WERKS 3200 ({fetchStart} ~ {fetchEnd})…");
            var rows3200 = await FetchRawRowsAsync(client, "3200", fetchStart, fetchEnd, progress);

            progress?.Report($"Fetching WERKS 3220 ({fetchStart} ~ {fetchEnd})…");
            var rows3220 = await FetchRawRowsAsync(client, "3220", fetchStart, fetchEnd, progress);

            var serverRows = new List<Dictionary<string, string>>();
            if (rows3200 != null) serverRows.AddRange(rows3200);
            if (rows3220 != null) serverRows.AddRange(rows3220);

            // toFetch에 포함된 날짜만 필터 (캐시 날짜 데이터 제외)
            var fetchDateSet = new HashSet<string>(
                toFetch.Select(d => d.ToString("yyyy-MM-dd")), StringComparer.Ordinal);
            serverRows = serverRows
                .Where(r => fetchDateSet.Contains(GetCol(r, "PRODUCT_DATE")))
                .ToList();

            progress?.Report($"Collected {serverRows.Count:N0} rows. Removing duplicates…");
            serverRows = RemoveDuplicates(serverRows);
            progress?.Report($"{serverRows.Count:N0} rows after deduplication.");

            // 오늘/어제 제외한 날짜 → per-day DB에 캐시 저장 (0건이어도 빈 파일 → 다음 조회 시 서버 생략)
            foreach (var date in toFetch.Where(d => d < recentCutoff))
            {
                string dateStr = date.ToString("yyyy-MM-dd");
                var    dayRows = serverRows
                    .Where(r => GetCol(r, "PRODUCT_DATE") == dateStr)
                    .ToList();
                SavePerDayDb(date, dayRows, progress);
            }

            freshRows = serverRows;
        }

        // ── 3. per-day 캐시 로드 ────────────────────────────────────────────
        var cachedRows = new List<Dictionary<string, string>>();
        foreach (var date in toCache)
        {
            var rows = LoadFromPerDayDb(GetPerDayDbPath(date));
            cachedRows.AddRange(rows);
            progress?.Report($"  Cache hit {date:MM/dd}: {rows.Count:N0} rows");
        }

        // ── 4. 병합 ──────────────────────────────────────────────────────────
        var allRows = new List<Dictionary<string, string>>(freshRows.Count + cachedRows.Count);
        allRows.AddRange(freshRows);
        allRows.AddRange(cachedRows);

        if (allRows.Count == 0)
        {
            progress?.Report("[ERROR] No data available for the selected date range.");
            return null;
        }

        // ── 5. 이전 temp DB 정리 → 새 temp DB 생성 ──────────────────────────
        CleanupTempDbs(progress);

        string tempPath = GetTempDbPath();
        try { Directory.CreateDirectory(Path.GetDirectoryName(tempPath)!); }
        catch (Exception ex)
        {
            progress?.Report($"[ERROR] Cannot create output folder: {ex.Message}");
            return null;
        }

        progress?.Report($"Saving to temp DB: {Path.GetFileName(tempPath)}");
        SaveToSqlite(tempPath, allRows);

        // ── 6. Post-processing ───────────────────────────────────────────────
        progress?.Report("Running post-processing (Routing / Reason / LineShift)…");
        await Task.Run(() => ProcessData(tempPath, progress));

        progress?.Report($"Done. DB: {Path.GetFileName(tempPath)}");
        return tempPath;
    }

    // ── Per-day DB / Temp DB 헬퍼 ────────────────────────────────────────────

    /// <summary>per-day 캐시 DB 경로: {DbSaveDirectory}/daily/yyyyMMdd.db</summary>
    private string GetPerDayDbPath(DateTime date)
        => Path.Combine(_settings.DbSaveDirectory, "daily", $"{date:yyyyMMdd}.db");

    /// <summary>임시 병합 DB 경로: {DbSaveDirectory}/temp_yyyyMMdd_HHmmss.db</summary>
    private string GetTempDbPath()
    {
        string ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        return Path.Combine(_settings.DbSaveDirectory, $"temp_{ts}.db");
    }

    private void SavePerDayDb(
        DateTime date, List<Dictionary<string, string>> rows, IProgress<string>? progress)
    {
        try
        {
            string path = GetPerDayDbPath(date);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);

            if (rows.Count == 0)
            {
                // 데이터 없는 날짜도 빈 DB 생성 → 다음 조회 시 File.Exists == true → 서버 생략
                using var conn = new SqliteConnection($"Data Source={path}");
                conn.Open();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $"CREATE TABLE IF NOT EXISTS [OrginalTable] " +
                    $"({string.Join(", ", OrgTableColumns.Select(c => $"[{c}] TEXT"))})";
                cmd.ExecuteNonQuery();
                progress?.Report($"  Cached {date:MM/dd}: no data (empty marker)");
                return;
            }

            SaveToSqlite(path, rows);
            progress?.Report($"  Cached {date:MM/dd}: {rows.Count:N0} rows → {Path.GetFileName(path)}");
        }
        catch (Exception ex)
        {
            progress?.Report($"[WARN] Failed to cache {date:MM/dd}: {ex.Message}");
        }
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
            // value before name 패턴도 처리
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

    // ── Private: 데이터 처리 ─────────────────────────────────────────────────────

    /// <summary>
    /// WPF clMakeProcTable.SelectRowsForProcTable 방식 동일.
    /// 같은 (LINE, CODE, PROCESSNAME, NGNAME, MATERIAL, DATE, SHIFT) 그룹 내에서
    /// QTYINPUT / QTYNG 비교 후 자동 병합·선택.
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

            // 마지막 행부터 시작해서 앞쪽 행과 순서대로 비교·병합 (WPF 방식 동일)
            var selected = groupRows[groupRows.Count - 1];
            for (int i = groupRows.Count - 2; i >= 0; i--)
                selected = ResolveRow(groupRows[i], selected);

            result.Add(selected);
        }

        return result;
    }

    /// <summary>
    /// WPF TryResolveRowsWithoutPrompt 자동 해결 로직 (프롬프트 없이 규칙만 적용).
    /// 1. 동일 값이면 → B 유지
    /// 2. QTYINPUT 다르면 → QTYINPUT 큰 쪽 선택 (WPF는 사용자 선택, 웹은 자동)
    /// 3. QTYINPUT 같고 QTYNG 한쪽만 0 → non-zero 선택
    /// 4. QTYINPUT 같고 둘 다 0 → B 유지 (WPF는 사용자 선택, 웹은 자동)
    /// 5. QTYINPUT 같고 둘 다 non-zero → QTYNG 합산 후 병합
    /// </summary>
    private static Dictionary<string, string> ResolveRow(
        Dictionary<string, string> optionA, Dictionary<string, string> optionB)
    {
        double inputA = ParseDouble(GetCol(optionA, "QTYINPUT"));
        double inputB = ParseDouble(GetCol(optionB, "QTYINPUT"));
        double ngA    = ParseDouble(GetCol(optionA, "QTYNG"));
        double ngB    = ParseDouble(GetCol(optionB, "QTYNG"));

        // 1. 동일 값
        if (inputA == inputB && ngA == ngB)
            return optionB;

        // 2. QTYINPUT 다름 → 큰 쪽 선택
        if (inputA != inputB)
            return inputA > inputB ? optionA : optionB;

        // QTYINPUT 같고 QTYNG 다름
        bool aZero = ngA == 0;
        bool bZero = ngB == 0;

        // 3. 한쪽만 0 → non-zero 선택
        if (aZero != bZero)
            return aZero ? optionB : optionA;

        // 4. 둘 다 0
        if (aZero)
            return optionB;

        // 5. 둘 다 non-zero → QTYNG 합산 병합
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
        // 실제 데이터에 있는 키를 파악하여 OrgTableColumns 순서로 정렬
        var dataKeys = rows.SelectMany(r => r.Keys)
                           .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var columns = OrgTableColumns
            .Where(c => dataKeys.Contains(c))
            .ToList();

        // OrgTableColumns에 없는 나머지 컬럼 추가
        foreach (string k in dataKeys)
        {
            if (!columns.Any(c => c.Equals(k, StringComparison.OrdinalIgnoreCase)))
                columns.Add(k);
        }

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        // 테이블 생성
        string colDefs = string.Join(", ", columns.Select(c => $"[{c}] TEXT"));
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = $"CREATE TABLE IF NOT EXISTS [OrginalTable] ({colDefs})";
            cmd.ExecuteNonQuery();
        }

        // 행 삽입 (트랜잭션)
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

        // 메타 테이블 저장
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
    /// DataMaker clDataProcessor.ProcessData() 와 동일한 흐름:
    ///   1. RoutingTable 로드 (Routing.txt)
    ///   2. ReasonTable 로드 (reason.txt)
    ///   3. OrginalTable LineShift 컬럼 설정
    /// </summary>
    private void ProcessData(string dbPath, IProgress<string>? progress)
    {
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        // Routing — settings DB 우선, 없으면 파일 폴백
        var routingDbRows = _settings.GetRoutingRows();
        if (routingDbRows.Count > 0)
        {
            progress?.Report($"Loading Routing table from settings DB ({routingDbRows.Count} rows)…");
            LoadRoutingFromSettings(conn, routingDbRows);
            NormalizeTableColumns(conn, "RoutingTable", new[] { "ProcessName", "ProcessType" });
            progress?.Report("Routing table done.");
        }
        else if (File.Exists(_settings.RoutingFilePath))
        {
            progress?.Report("Loading Routing table from file…");
            LoadTsvToTable(conn, _settings.RoutingFilePath, "RoutingTable",
                new[] { "모델명", "ProcessCode", "ProcessName", "ProcessType" });
            NormalizeTableColumns(conn, "RoutingTable", new[] { "ProcessName", "ProcessType" });
            progress?.Report("Routing table done.");
        }
        else
        {
            progress?.Report("[WARN] No Routing data (settings DB empty, file not found).");
        }

        // Reason — settings DB 우선, 없으면 파일 폴백
        var reasonDbRows = _settings.GetReasonRows();
        if (reasonDbRows.Count > 0)
        {
            progress?.Report($"Loading Reason table from settings DB ({reasonDbRows.Count} rows)…");
            LoadReasonFromSettings(conn, reasonDbRows);
            NormalizeTableColumns(conn, "Reason", new[] { "processName", "NgName" });
            progress?.Report("Reason table done.");
        }
        else if (File.Exists(_settings.ReasonFilePath))
        {
            progress?.Report("Loading Reason table from file…");
            LoadTsvToTable(conn, _settings.ReasonFilePath, "Reason",
                new[] { "processName", "NgName", "Reason" });
            NormalizeTableColumns(conn, "Reason", new[] { "processName", "NgName" });
            progress?.Report("Reason table done.");
        }
        else
        {
            progress?.Report("[WARN] No Reason data (settings DB empty, file not found).");
        }

        // SetLineShiftColumnValue: LineShift = MATERIALNAME || '_' || PRODUCTION_LINE
        progress?.Report("Setting LineShift column…");
        SetLineShift(conn);
        progress?.Report("Processing complete.");
    }

    /// <summary>
    /// Tab 구분자 텍스트 파일 → SQLite 테이블 (첫 줄 헤더 스킵, Columns 파라미터 사용)
    /// clMakeTxtTable.MakeDataTable 동일 로직
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
        if (lines.Length < 2) return; // 헤더만 있거나 빈 파일

        // Create table
        string colDefs = string.Join(", ", columns.Select(c => $"[{c}] TEXT"));
        using (var create = conn.CreateCommand())
        {
            create.CommandText = $"CREATE TABLE [{tableName}] ({colDefs})";
            create.ExecuteNonQuery();
        }

        // Insert (첫 줄 = 헤더 스킵)
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
    /// 지정 컬럼에 CONSTANT.Normalize 동일 정규화 적용 후 재저장
    /// </summary>
    private static void NormalizeTableColumns(
        SqliteConnection conn, string tableName, string[] columnNames)
    {
        // 한 컬럼씩 UPDATE로 정규화 적용 (SQLite에서 regex 함수 없으므로 로드 후 재저장)
        // 행이 많지 않은 마스터 테이블에만 사용
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
        // ADD COLUMN (이미 있으면 무시)
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
