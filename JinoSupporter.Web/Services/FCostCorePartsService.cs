using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

/// <summary>
/// MES072410 core-parts F-COST collector and parser.
/// Uses configured BMES credentials and a fresh cookie container; browser cookies or
/// copied SSO/session tokens are never accepted by this service.
/// </summary>
public sealed class FCostCorePartsService(
    NgRateSettingsService settings,
    AppActivityLogger activity)
{
    private const string BaseUrl = "https://bmes.bujeon.com";
    private readonly NgRateSettingsService _settings = settings;
    private readonly AppActivityLogger _activity = activity;

    public string GetRawDbPath()
        => Path.Combine(_settings.FCostDbSaveDirectory, "fcost_raw.db");

    public FCostCorePartsStatus GetStatus()
        => ReadStatus(GetRawDbPath());

    public FCostCorePartsKpiSnapshot? GetKpiSnapshot(DateTime atOrBefore)
        => ReadKpiSnapshot(GetRawDbPath(), atOrBefore);

    public FCostCorePartsKpiSnapshot? GetKpiRangeSnapshot(DateTime start, DateTime end)
        => ReadKpiRangeSnapshot(GetRawDbPath(), start, end);

    public async Task<FCostCorePartsBackfillResult> BackfillAsync(
        DateTime start,
        DateTime end,
        bool force = false,
        DateTime? forceFromDate = null,
        TimeSpan? forceRefreshTtl = null,
        int delayMs = 1200,
        int queryIntervalDays = 1,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        start = start.Date;
        end = end.Date;
        if (end > DateTime.Today) end = DateTime.Today;
        if (start > end)
            throw new ArgumentException("Start date must be before or equal to end date.");
        if (queryIntervalDays < 1)
            throw new ArgumentOutOfRangeException(
                nameof(queryIntervalDays),
                "Query interval must be at least one day.");

        Directory.CreateDirectory(_settings.FCostDbSaveDirectory);
        string dbPath = GetRawDbPath();
        await Task.Run(() => EnsureDatabase(dbPath), cancellationToken);
        List<DateTime> queryDates = BuildQueryDates(start, end, queryIntervalDays);

        var result = new FCostCorePartsBackfillResult
        {
            DbPath = dbPath,
            StartDate = start,
            EndDate = end,
        };

        int totalDays = queryDates.Count;
        var datesToFetch = new List<(DateTime Day, int Ordinal)>();
        for (int dateIndex = 0; dateIndex < queryDates.Count; dateIndex++)
        {
            DateTime day = queryDates[dateIndex];
            cancellationToken.ThrowIfCancellationRequested();
            int ordinal = dateIndex + 1;
            result.AttemptedDays++;

            bool isRefreshWindow =
                forceFromDate is { } refreshFrom &&
                day >= refreshFrom.Date;
            bool cachedPullIsUsable = !force && PullSucceeded(
                dbPath,
                day,
                isRefreshWindow ? forceRefreshTtl : null);
            bool forceDay =
                force ||
                (isRefreshWindow &&
                 (forceRefreshTtl is null || !cachedPullIsUsable));
            if (!forceDay && cachedPullIsUsable)
            {
                result.SkippedDays++;
                Report(progress, $"[{ordinal}/{totalDays}] {day:yyyy-MM-dd}: skip (already OK)");
                continue;
            }

            datesToFetch.Add((day, ordinal));
        }

        if (datesToFetch.Count == 0)
        {
            Report(
                progress,
                $"MES072410 cache complete: skipped={result.SkippedDays:N0}, no BMES login required");
            return result;
        }

        string loginId = _settings.LoginId;
        string password = _settings.Password;
        if (string.IsNullOrWhiteSpace(loginId) || string.IsNullOrWhiteSpace(password))
        {
            result.Failures.Add("BMES credentials are not configured.");
            return result;
        }

        using var handler = new HttpClientHandler
        {
            UseCookies = true,
            CookieContainer = new System.Net.CookieContainer(),
        };
        using var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(300) };
        client.DefaultRequestHeaders.Add("X-Requested-With", "XMLHttpRequest");

        Report(progress, "MES072410: reading login token");
        string token = await GetTokenAsync(client, cancellationToken);
        if (string.IsNullOrEmpty(token))
        {
            result.Failures.Add("Failed to read BMES login token.");
            return result;
        }

        Report(progress, "MES072410: logging in with configured credentials");
        if (!await LoginAsync(client, token, loginId, password, cancellationToken))
        {
            result.Failures.Add("BMES login failed.");
            return result;
        }

        for (int fetchIndex = 0; fetchIndex < datesToFetch.Count; fetchIndex++)
        {
            (DateTime day, int ordinal) = datesToFetch[fetchIndex];
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                Report(progress, $"[{ordinal}/{totalDays}] {day:yyyy-MM-dd}: fetch");
                FCostCorePartsParseResult parsed =
                    await FetchDateAsync(client, day, cancellationToken);
                await Task.Run(
                    () => SaveParsedToDatabase(dbPath, day, parsed),
                    cancellationToken);
                result.FetchedDays++;
                result.TotalRows += parsed.Rows.Count;
                Report(progress, $"{day:yyyy-MM-dd}: saved {parsed.Rows.Count:N0} rows");
            }
            catch (Exception ex)
            {
                result.FailedDays++;
                string message = $"{day:yyyy-MM-dd}: {ex.Message}";
                result.Failures.Add(message);
                await Task.Run(
                    () => SaveFailure(dbPath, day, message),
                    cancellationToken);
                Report(progress, "[WARN] " + message);
            }

            if (delayMs > 0 && fetchIndex < datesToFetch.Count - 1)
                await Task.Delay(delayMs, cancellationToken);
        }

        return result;
    }

    public static FCostCorePartsParseResult ParseSearchListJson(string json)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(json);

        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement root = document.RootElement;
        if (!TryGetProperty(root, "data", out JsonElement data) ||
            data.ValueKind != JsonValueKind.Object)
        {
            throw new JsonException("MES072410 response does not contain data.");
        }

        if (!TryGetProperty(data, "contents", out JsonElement contents) ||
            contents.ValueKind != JsonValueKind.Array)
        {
            throw new JsonException("MES072410 response does not contain data.contents.");
        }

        var rows = new List<FCostCorePartsRow>(contents.GetArrayLength());
        int columnCount = 0;
        foreach (JsonElement item in contents.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Object)
                continue;

            var row = new FCostCorePartsRow
            {
                RawJson = item.GetRawText(),
                FaccoTx = ReadText(item, "FACCO_TX"),
                CptypTx = ReadText(item, "CPTYP_TX"),
                ZtypeTx = ReadText(item, "ZTYPE_TX"),
                Ztote = ReadText(item, "ZTOTE"),
                Zsort01 = ReadText(item, "ZSORT_01"),
                Zsort02 = ReadText(item, "ZSORT_02"),
                Zsort03 = ReadText(item, "ZSORT_03"),
                Zsort04 = ReadText(item, "ZSORT_04"),
                Cptyp = ReadText(item, "CPTYP"),
                Tsort = ReadText(item, "TSORT"),
                Zsort = ReadText(item, "ZSORT"),
                Facco = ReadText(item, "FACCO"),
                Ztype = ReadText(item, "ZTYPE"),
                Chk = ReadText(item, "CHK"),
            };

            for (int index = 1; index <= 14; index++)
            {
                string propertyName = $"COL{index:D4}";
                if (TryGetProperty(item, propertyName, out JsonElement value))
                {
                    columnCount = Math.Max(columnCount, index);
                    row.SetCol(index, ReadNullableDouble(value));
                }
            }

            rows.Add(row);
        }

        ForwardFillTriplets(rows);
        List<FCostCorePartsColumn> columns = ParseTrustedColumnMetadata(data);
        if (columns.Count > 0)
            columnCount = Math.Max(columnCount, columns.Max(column => column.ColIndex));

        return new FCostCorePartsParseResult(rows, columns, columnCount);
    }

    public static FCostCorePartsKpiSnapshot BuildKpiSnapshot(
        DateTime queryDate,
        FCostCorePartsParseResult parsed)
    {
        ArgumentNullException.ThrowIfNull(parsed);

        List<(int Index, FCostCorePartsRow Row, int Priority)> anchors = parsed.Rows
            .Select((row, index) => (Index: index, Row: row, Priority: TotalMarkerPriority(row)))
            .Where(candidate =>
                string.Equals(candidate.Row.Ztype, "INAMT", StringComparison.OrdinalIgnoreCase))
            .ToList();

        List<(int Index, FCostCorePartsRow Row, int Priority)> markedAnchors = anchors
            .Where(candidate => candidate.Priority > 0)
            .OrderByDescending(candidate => candidate.Priority)
            .ThenBy(candidate => candidate.Index)
            .ToList();
        (int Index, FCostCorePartsRow Row, int Priority)? selected =
            markedAnchors.Count > 0 ? markedAnchors[0] : null;
        if (selected is null && anchors.Count == 1)
            selected = anchors[0];

        FCostCorePartsRow? costRow = null;
        FCostCorePartsRow? rateRow = null;
        if (selected is { } total)
        {
            for (int index = total.Index + 1; index < parsed.Rows.Count; index++)
            {
                FCostCorePartsRow row = parsed.Rows[index];
                if (string.Equals(row.Ztype, "INAMT", StringComparison.OrdinalIgnoreCase))
                    break;
                if (costRow is null &&
                    string.Equals(row.Ztype, "FCOST", StringComparison.OrdinalIgnoreCase))
                {
                    costRow = row;
                }
                else if (rateRow is null &&
                         string.Equals(row.Ztype, "FRATE", StringComparison.OrdinalIgnoreCase))
                {
                    rateRow = row;
                }
            }
        }

        IReadOnlyList<FCostCorePartsKpiPeriod> periods = parsed.Columns
            .OrderBy(column => column.ColIndex)
            .Select(column => new FCostCorePartsKpiPeriod(
                column.ColIndex,
                column.Code,
                column.Header,
                column.PDate,
                column.Kind,
                costRow?.GetCol(column.ColIndex),
                rateRow?.GetCol(column.ColIndex)))
            .ToList();

        return new FCostCorePartsKpiSnapshot(queryDate.Date, periods);
    }

    /// <summary>
    /// Idempotently replaces one query date in the endpoint-specific tables.
    /// Exposed so parser/storage behavior can be verified without a BMES request.
    /// </summary>
    public static void SaveParsedToDatabase(
        string dbPath,
        DateTime queryDate,
        FCostCorePartsParseResult parsed)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(dbPath);
        ArgumentNullException.ThrowIfNull(parsed);

        string? directory = Path.GetDirectoryName(Path.GetFullPath(dbPath));
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);
        EnsureDatabase(dbPath);

        using var connection = OpenConnection(dbPath);
        using SqliteTransaction transaction = connection.BeginTransaction();
        string dateText = queryDate.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        string fetchedAt = DateTime.Now.ToString("O", CultureInfo.InvariantCulture);

        DeleteDateRows(connection, transaction, dateText);
        InsertRows(connection, transaction, dateText, fetchedAt, parsed.Rows);
        InsertColumns(connection, transaction, dateText, parsed.Columns);
        UpsertPull(
            connection,
            transaction,
            dateText,
            "OK",
            fetchedAt,
            parsed.Rows.Count,
            parsed.ColumnCount,
            string.Empty);
        transaction.Commit();
    }

    public static FCostCorePartsStatus ReadStatus(string dbPath)
    {
        var status = new FCostCorePartsStatus
        {
            DbPath = dbPath,
            Exists = File.Exists(dbPath),
        };
        if (!status.Exists)
            return status;

        try
        {
            using var connection = OpenConnection(dbPath, readOnly: true);
            if (!TableExists(connection, "MES072410Pulls"))
                return status;

            using (var command = connection.CreateCommand())
            {
                command.CommandText =
                    """
                    SELECT
                        COUNT(*),
                        COALESCE(SUM(CASE WHEN Status = 'OK' THEN 1 ELSE 0 END), 0),
                        COALESCE(SUM(CASE WHEN Status <> 'OK' THEN 1 ELSE 0 END), 0),
                        COALESCE(MIN(QueryDate), ''),
                        COALESCE(MAX(QueryDate), '')
                    FROM MES072410Pulls;
                    """;
                using SqliteDataReader reader = command.ExecuteReader();
                if (reader.Read())
                {
                    status.PullCount = reader.GetInt32(0);
                    status.SuccessCount = reader.GetInt32(1);
                    status.FailedCount = reader.GetInt32(2);
                    status.FirstDate = reader.GetString(3);
                    status.LastDate = reader.GetString(4);
                }
            }

            if (TableExists(connection, "MES072410Rows"))
            {
                using var command = connection.CreateCommand();
                command.CommandText = "SELECT COUNT(*) FROM MES072410Rows;";
                status.TotalRows = Convert.ToInt32(command.ExecuteScalar() ?? 0);
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

    public static FCostCorePartsKpiSnapshot? ReadKpiRangeSnapshot(
        string dbPath,
        DateTime start,
        DateTime end)
    {
        start = start.Date;
        end = end.Date;
        if (start > end || string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath))
            return null;

        var storedDates = new List<DateTime>();
        using (var connection = OpenConnection(dbPath, readOnly: true))
        {
            if (!TableExists(connection, "MES072410Pulls") ||
                !TableExists(connection, "MES072410Rows") ||
                !TableExists(connection, "MES072410Columns"))
            {
                return null;
            }

            using var command = connection.CreateCommand();
            command.CommandText =
                """
                SELECT p.QueryDate
                FROM MES072410Pulls p
                WHERE p.QueryDate BETWEEN @startDate AND @endDate
                  AND EXISTS (
                      SELECT 1
                      FROM MES072410Columns c
                      WHERE c.QueryDate = p.QueryDate
                  )
                  AND EXISTS (
                      SELECT 1
                      FROM MES072410Rows r
                      WHERE r.QueryDate = p.QueryDate
                  )
                ORDER BY p.QueryDate;
                """;
            command.Parameters.AddWithValue(
                "@startDate",
                start.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
            command.Parameters.AddWithValue(
                "@endDate",
                end.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
            using SqliteDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                if (DateTime.TryParseExact(
                        reader.GetString(0),
                        "yyyy-MM-dd",
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out DateTime storedDate))
                {
                    storedDates.Add(storedDate);
                }
            }
        }

        var periodsByKey =
            new Dictionary<string, FCostCorePartsKpiPeriod>(StringComparer.OrdinalIgnoreCase);
        DateTime latestDate = DateTime.MinValue;
        foreach (DateTime storedDate in storedDates)
        {
            FCostCorePartsKpiSnapshot? snapshot = ReadKpiSnapshot(dbPath, storedDate);
            if (snapshot is null || snapshot.QueryDate != storedDate)
                continue;

            latestDate = storedDate;
            foreach (FCostCorePartsKpiPeriod period in snapshot.Periods)
            {
                // Iteration is oldest → newest, so a later snapshot replaces a partial
                // week/month value emitted earlier in that same reporting period.
                periodsByKey[CorePartsPeriodIdentity(period)] = period;
            }
        }

        return latestDate == DateTime.MinValue
            ? null
            : new FCostCorePartsKpiSnapshot(
                latestDate,
                periodsByKey.Values
                    .OrderBy(period => period.Kind, StringComparer.Ordinal)
                    .ThenBy(period => period.PDate, StringComparer.Ordinal)
                    .ThenBy(period => period.ColIndex)
                    .ToList());
    }

    public static FCostCorePartsKpiSnapshot? ReadKpiSnapshot(
        string dbPath,
        DateTime atOrBefore)
    {
        if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath))
            return null;

        using var connection = OpenConnection(dbPath, readOnly: true);
        if (!TableExists(connection, "MES072410Pulls") ||
            !TableExists(connection, "MES072410Rows") ||
            !TableExists(connection, "MES072410Columns"))
        {
            return null;
        }

        string queryDate;
        using (var command = connection.CreateCommand())
        {
            command.CommandText =
                """
                SELECT p.QueryDate
                FROM MES072410Pulls p
                WHERE p.QueryDate <= @endDate
                  AND EXISTS (
                      SELECT 1
                      FROM MES072410Columns c
                      WHERE c.QueryDate = p.QueryDate
                  )
                  AND EXISTS (
                      SELECT 1
                      FROM MES072410Rows r
                      WHERE r.QueryDate = p.QueryDate
                  )
                ORDER BY p.QueryDate DESC
                LIMIT 1;
                """;
            command.Parameters.AddWithValue(
                "@endDate",
                atOrBefore.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
            queryDate = Convert.ToString(
                command.ExecuteScalar(),
                CultureInfo.InvariantCulture) ?? string.Empty;
        }
        if (queryDate.Length == 0)
            return null;

        var columns = new List<FCostCorePartsColumn>();
        using (var command = connection.CreateCommand())
        {
            command.CommandText =
                """
                SELECT ColIndex, Code, Header, PDate, Kind
                FROM MES072410Columns
                WHERE QueryDate = @queryDate
                ORDER BY ColIndex;
                """;
            command.Parameters.AddWithValue("@queryDate", queryDate);
            using SqliteDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                columns.Add(new FCostCorePartsColumn(
                    reader.GetInt32(0),
                    reader.GetString(1),
                    reader.GetString(2),
                    reader.GetString(3),
                    reader.GetString(4)));
            }
        }

        var rows = new List<FCostCorePartsRow>();
        using (var command = connection.CreateCommand())
        {
            command.CommandText =
                """
                SELECT
                    FACCO_TX, CPTYP_TX, ZTYPE_TX, ZTOTE,
                    CPTYP, FACCO, ZTYPE,
                    Col0001, Col0002, Col0003, Col0004, Col0005, Col0006, Col0007,
                    Col0008, Col0009, Col0010, Col0011, Col0012, Col0013, Col0014
                FROM MES072410Rows
                WHERE QueryDate = @queryDate
                ORDER BY RowNo;
                """;
            command.Parameters.AddWithValue("@queryDate", queryDate);
            using SqliteDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                var row = new FCostCorePartsRow
                {
                    FaccoTx = ReadDbText(reader, 0),
                    CptypTx = ReadDbText(reader, 1),
                    ZtypeTx = ReadDbText(reader, 2),
                    Ztote = ReadDbText(reader, 3),
                    Cptyp = ReadDbText(reader, 4),
                    Facco = ReadDbText(reader, 5),
                    Ztype = ReadDbText(reader, 6),
                };
                for (int columnIndex = 1; columnIndex <= 14; columnIndex++)
                {
                    int ordinal = columnIndex + 6;
                    row.SetCol(
                        columnIndex,
                        reader.IsDBNull(ordinal) ? null : reader.GetDouble(ordinal));
                }
                rows.Add(row);
            }
        }

        DateTime parsedDate = DateTime.TryParseExact(
            queryDate,
            "yyyy-MM-dd",
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out DateTime storedDate)
            ? storedDate
            : atOrBefore.Date;
        int columnCount = columns.Count > 0
            ? columns.Max(column => column.ColIndex)
            : 0;
        return BuildKpiSnapshot(
            parsedDate,
            new FCostCorePartsParseResult(rows, columns, columnCount));
    }

    private static List<DateTime> BuildQueryDates(
        DateTime start,
        DateTime end,
        int intervalDays)
    {
        if (intervalDays == 1)
        {
            return Enumerable.Range(0, (end - start).Days + 1)
                .Select(offset => start.AddDays(offset))
                .ToList();
        }

        var dates = new HashSet<DateTime> { start, end };
        for (DateTime cursor = end;
             cursor > start && cursor >= DateTime.MinValue.AddDays(intervalDays);
             cursor = cursor.AddDays(-intervalDays))
        {
            dates.Add(cursor);
        }
        return dates.OrderBy(date => date).ToList();
    }

    private static string CorePartsPeriodIdentity(FCostCorePartsKpiPeriod period)
    {
        string periodCode = !string.IsNullOrWhiteSpace(period.PDate)
            ? period.PDate
            : !string.IsNullOrWhiteSpace(period.Code)
                ? period.Code
                : period.Header;
        return period.Kind + "|" + periodCode;
    }

    private async Task<FCostCorePartsParseResult> FetchDateAsync(
        HttpClient client,
        DateTime queryDate,
        CancellationToken cancellationToken)
    {
        string url =
            $"{BaseUrl}/MES072410/SearchList?perPage=" +
            $"&Condition.SDATE={queryDate:yyyy-MM-dd}" +
            "&Condition.ZGUBN=D" +
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

        using HttpResponseMessage response = await client.GetAsync(url, cancellationToken);
        string json = await response.Content.ReadAsStringAsync(cancellationToken);
        response.EnsureSuccessStatusCode();
        return ParseSearchListJson(json);
    }

    private void Report(IProgress<string>? progress, string message)
    {
        _activity.Log("FCostCoreParts", message);
        progress?.Report(message);
    }

    private static async Task<string> GetTokenAsync(
        HttpClient client,
        CancellationToken cancellationToken)
    {
        try
        {
            string html = await client.GetStringAsync(BaseUrl, cancellationToken);
            Match match = Regex.Match(
                html,
                @"<input[^>]+name=""__RequestVerificationToken""[^>]+value=""([^""]+)""",
                RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                match = Regex.Match(
                    html,
                    @"<input[^>]+value=""([^""]+)""[^>]+name=""__RequestVerificationToken""",
                    RegexOptions.IgnoreCase);
            }
            return match.Success ? match.Groups[1].Value : string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static async Task<bool> LoginAsync(
        HttpClient client,
        string token,
        string loginId,
        string password,
        CancellationToken cancellationToken)
    {
        using var content = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["UserInfo[USRID]"] = loginId,
            ["UserInfo[PWNO]"] = password,
            ["UserInfo[LANG]"] = "EN",
            ["UserInfo[FACCO]"] = "GN",
            ["UserInfo[STYPE]"] = "P",
            ["UserInfo[VTYPE]"] = "P",
            ["__RequestVerificationToken"] = token,
        });
        try
        {
            using HttpResponseMessage response = await client.PostAsync(
                BaseUrl + "/MES000000/LoginCheck",
                content,
                cancellationToken);
            string body = await response.Content.ReadAsStringAsync(cancellationToken);
            return response.IsSuccessStatusCode &&
                   body.Contains("\"Result\":\"M\"", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static List<FCostCorePartsColumn> ParseTrustedColumnMetadata(JsonElement data)
    {
        var result = new List<FCostCorePartsColumn>();
        if (!TryGetProperty(data, "BottomGridColumnList", out JsonElement metadata) ||
            metadata.ValueKind != JsonValueKind.Array)
        {
            return result;
        }

        foreach (JsonElement item in metadata.EnumerateArray())
        {
            string columnName = ReadText(item, "ZCOLN");
            int index = ParseColumnIndex(columnName);
            if (index is < 1 or > 14)
                continue;

            string code = ReadText(item, "ZPVTT");
            string header = ReadText(item, "ZPVTT_TX");
            string pdate = ReadText(item, "PDATE");
            if (string.IsNullOrWhiteSpace(code) &&
                string.IsNullOrWhiteSpace(header) &&
                string.IsNullOrWhiteSpace(pdate))
            {
                continue;
            }

            result.Add(new FCostCorePartsColumn(
                index,
                code,
                header,
                pdate,
                InferPeriodKind(code, pdate)));
        }

        return result
            .GroupBy(column => column.ColIndex)
            .Select(group => group.First())
            .OrderBy(column => column.ColIndex)
            .ToList();
    }

    private static string InferPeriodKind(string code, string pdate)
    {
        string candidate = string.IsNullOrWhiteSpace(code) ? pdate : code;
        if (candidate.Contains('-', StringComparison.Ordinal))
            return "Week";
        string digits = new(candidate.Where(char.IsDigit).ToArray());
        return digits.Length switch
        {
            8 => "Day",
            6 => "Month",
            _ => "Unknown",
        };
    }

    private static int ParseColumnIndex(string columnName)
    {
        string digits = new(columnName.Where(char.IsDigit).ToArray());
        return int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out int index)
            ? index
            : 0;
    }

    private static void ForwardFillTriplets(IReadOnlyList<FCostCorePartsRow> rows)
    {
        FCostCorePartsRow? anchor = null;
        foreach (FCostCorePartsRow row in rows)
        {
            if (string.Equals(row.Ztype, "INAMT", StringComparison.OrdinalIgnoreCase))
            {
                anchor = row;
                continue;
            }

            bool continuation =
                string.Equals(row.Ztype, "FCOST", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(row.Ztype, "FRATE", StringComparison.OrdinalIgnoreCase);
            if (!continuation)
            {
                anchor = null;
                continue;
            }
            if (anchor is null)
                continue;

            if (string.IsNullOrWhiteSpace(row.FaccoTx)) row.FaccoTx = anchor.FaccoTx;
            if (string.IsNullOrWhiteSpace(row.CptypTx)) row.CptypTx = anchor.CptypTx;
            if (string.IsNullOrWhiteSpace(row.Facco)) row.Facco = anchor.Facco;
            if (string.IsNullOrWhiteSpace(row.Cptyp)) row.Cptyp = anchor.Cptyp;
        }
    }

    private static int TotalMarkerPriority(FCostCorePartsRow row)
    {
        foreach (string text in new[]
                 {
                     row.FaccoTx,
                     row.CptypTx,
                     row.Facco,
                     row.Cptyp,
                 })
        {
            string normalized = text.Trim();
            if (normalized.Equals("TOTAL", StringComparison.OrdinalIgnoreCase) ||
                normalized.Equals("합계", StringComparison.Ordinal) ||
                normalized.Equals("전체", StringComparison.Ordinal))
            {
                return 2;
            }
        }

        string totalFlag = row.Ztote.Trim();
        if (totalFlag.Length == 0 ||
            totalFlag.Equals("0", StringComparison.OrdinalIgnoreCase) ||
            totalFlag.Equals("N", StringComparison.OrdinalIgnoreCase) ||
            totalFlag.Equals("NO", StringComparison.OrdinalIgnoreCase) ||
            totalFlag.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        return 1;
    }

    private static string ReadDbText(SqliteDataReader reader, int ordinal)
        => reader.IsDBNull(ordinal) ? string.Empty : reader.GetString(ordinal);

    private static string ReadText(JsonElement element, string propertyName)
    {
        if (!TryGetProperty(element, propertyName, out JsonElement value) ||
            value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
        {
            return string.Empty;
        }

        return value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? string.Empty
            : value.GetRawText();
    }

    private static double? ReadNullableDouble(JsonElement value)
    {
        if (value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        if (value.ValueKind == JsonValueKind.Number)
            return value.TryGetDouble(out double numeric) ? numeric : null;
        if (value.ValueKind != JsonValueKind.String)
            return null;

        string text = value.GetString()?.Trim() ?? string.Empty;
        if (text.Length == 0)
            return null;
        return double.TryParse(
            text,
            NumberStyles.Float | NumberStyles.AllowThousands,
            CultureInfo.InvariantCulture,
            out double parsed)
            ? parsed
            : null;
    }

    private static bool TryGetProperty(
        JsonElement element,
        string propertyName,
        out JsonElement value)
    {
        if (element.ValueKind == JsonValueKind.Object &&
            element.TryGetProperty(propertyName, out value))
        {
            return true;
        }

        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (JsonProperty property in element.EnumerateObject())
            {
                if (string.Equals(property.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    value = property.Value;
                    return true;
                }
            }
        }

        value = default;
        return false;
    }

    private static void EnsureDatabase(string dbPath)
    {
        using var connection = OpenConnection(dbPath);
        using var command = connection.CreateCommand();
        command.CommandText =
            """
            CREATE TABLE IF NOT EXISTS MES072410Pulls (
                QueryDate   TEXT PRIMARY KEY,
                Status      TEXT NOT NULL,
                FetchedAt   TEXT NOT NULL,
                RowCount    INTEGER NOT NULL,
                ColumnCount INTEGER NOT NULL,
                ErrorMessage TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS MES072410Rows (
                QueryDate TEXT NOT NULL,
                RowNo INTEGER NOT NULL,
                FetchedAt TEXT NOT NULL,
                RawJson TEXT NOT NULL,
                FACCO_TX TEXT,
                CPTYP_TX TEXT,
                ZTYPE_TX TEXT,
                ZTOTE TEXT,
                ZSORT_01 TEXT,
                ZSORT_02 TEXT,
                ZSORT_03 TEXT,
                ZSORT_04 TEXT,
                CPTYP TEXT,
                TSORT TEXT,
                ZSORT TEXT,
                FACCO TEXT,
                ZTYPE TEXT,
                CHK TEXT,
                Col0001 REAL, Col0002 REAL, Col0003 REAL, Col0004 REAL,
                Col0005 REAL, Col0006 REAL, Col0007 REAL, Col0008 REAL,
                Col0009 REAL, Col0010 REAL, Col0011 REAL, Col0012 REAL,
                Col0013 REAL, Col0014 REAL,
                PRIMARY KEY (QueryDate, RowNo)
            );

            CREATE TABLE IF NOT EXISTS MES072410Columns (
                QueryDate TEXT NOT NULL,
                ColIndex INTEGER NOT NULL,
                Code TEXT NOT NULL,
                Header TEXT NOT NULL,
                PDate TEXT NOT NULL,
                Kind TEXT NOT NULL,
                PRIMARY KEY (QueryDate, ColIndex)
            );
            """;
        command.ExecuteNonQuery();
    }

    private static SqliteConnection OpenConnection(string dbPath, bool readOnly = false)
    {
        var builder = new SqliteConnectionStringBuilder
        {
            DataSource = dbPath,
            Mode = readOnly ? SqliteOpenMode.ReadOnly : SqliteOpenMode.ReadWriteCreate,
        };
        var connection = new SqliteConnection(builder.ToString());
        connection.Open();
        return connection;
    }

    private static bool TableExists(SqliteConnection connection, string tableName)
    {
        using var command = connection.CreateCommand();
        command.CommandText =
            "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@name;";
        command.Parameters.AddWithValue("@name", tableName);
        return Convert.ToInt64(command.ExecuteScalar() ?? 0L) > 0;
    }

    private static bool PullSucceeded(
        string dbPath,
        DateTime queryDate,
        TimeSpan? maxAge = null)
    {
        if (!File.Exists(dbPath))
            return false;
        try
        {
            using var connection = OpenConnection(dbPath, readOnly: true);
            if (!TableExists(connection, "MES072410Pulls"))
                return false;
            using var command = connection.CreateCommand();
            command.CommandText =
                "SELECT Status, FetchedAt FROM MES072410Pulls WHERE QueryDate=@date;";
            command.Parameters.AddWithValue(
                "@date",
                queryDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
            using var reader = command.ExecuteReader();
            if (!reader.Read())
                return false;

            string status = reader.IsDBNull(0) ? string.Empty : reader.GetString(0);
            if (!string.Equals(status, "OK", StringComparison.OrdinalIgnoreCase))
                return false;
            if (maxAge is not TimeSpan ttl)
                return true;

            string fetchedText = reader.IsDBNull(1) ? string.Empty : reader.GetString(1);
            return DateTimeOffset.TryParse(
                       fetchedText,
                       CultureInfo.InvariantCulture,
                       DateTimeStyles.RoundtripKind,
                       out DateTimeOffset fetchedAt) &&
                   DateTimeOffset.Now - fetchedAt <= ttl;
        }
        catch
        {
            return false;
        }
    }

    private static void DeleteDateRows(
        SqliteConnection connection,
        SqliteTransaction transaction,
        string queryDate)
    {
        foreach (string table in new[] { "MES072410Rows", "MES072410Columns" })
        {
            using var command = connection.CreateCommand();
            command.Transaction = transaction;
            command.CommandText = $"DELETE FROM {table} WHERE QueryDate=@date;";
            command.Parameters.AddWithValue("@date", queryDate);
            command.ExecuteNonQuery();
        }
    }

    private static void InsertRows(
        SqliteConnection connection,
        SqliteTransaction transaction,
        string queryDate,
        string fetchedAt,
        IReadOnlyList<FCostCorePartsRow> rows)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText =
            """
            INSERT INTO MES072410Rows
            (QueryDate, RowNo, FetchedAt, RawJson,
             FACCO_TX, CPTYP_TX, ZTYPE_TX, ZTOTE,
             ZSORT_01, ZSORT_02, ZSORT_03, ZSORT_04,
             CPTYP, TSORT, ZSORT, FACCO, ZTYPE, CHK,
             Col0001, Col0002, Col0003, Col0004, Col0005, Col0006, Col0007,
             Col0008, Col0009, Col0010, Col0011, Col0012, Col0013, Col0014)
            VALUES
            (@date, @rowNo, @fetchedAt, @rawJson,
             @faccoTx, @cptypTx, @ztypeTx, @ztote,
             @zsort01, @zsort02, @zsort03, @zsort04,
             @cptyp, @tsort, @zsort, @facco, @ztype, @chk,
             @c01, @c02, @c03, @c04, @c05, @c06, @c07,
             @c08, @c09, @c10, @c11, @c12, @c13, @c14);
            """;

        SqliteParameter Add(string name, SqliteType type) => command.Parameters.Add(name, type);
        SqliteParameter pDate = Add("@date", SqliteType.Text);
        SqliteParameter pRowNo = Add("@rowNo", SqliteType.Integer);
        SqliteParameter pFetchedAt = Add("@fetchedAt", SqliteType.Text);
        SqliteParameter pRawJson = Add("@rawJson", SqliteType.Text);
        SqliteParameter pFaccoTx = Add("@faccoTx", SqliteType.Text);
        SqliteParameter pCptypTx = Add("@cptypTx", SqliteType.Text);
        SqliteParameter pZtypeTx = Add("@ztypeTx", SqliteType.Text);
        SqliteParameter pZtote = Add("@ztote", SqliteType.Text);
        SqliteParameter pZsort01 = Add("@zsort01", SqliteType.Text);
        SqliteParameter pZsort02 = Add("@zsort02", SqliteType.Text);
        SqliteParameter pZsort03 = Add("@zsort03", SqliteType.Text);
        SqliteParameter pZsort04 = Add("@zsort04", SqliteType.Text);
        SqliteParameter pCptyp = Add("@cptyp", SqliteType.Text);
        SqliteParameter pTsort = Add("@tsort", SqliteType.Text);
        SqliteParameter pZsort = Add("@zsort", SqliteType.Text);
        SqliteParameter pFacco = Add("@facco", SqliteType.Text);
        SqliteParameter pZtype = Add("@ztype", SqliteType.Text);
        SqliteParameter pChk = Add("@chk", SqliteType.Text);
        var pColumns = new SqliteParameter[14];
        for (int index = 0; index < pColumns.Length; index++)
            pColumns[index] = Add($"@c{index + 1:D2}", SqliteType.Real);

        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            FCostCorePartsRow row = rows[rowIndex];
            pDate.Value = queryDate;
            pRowNo.Value = rowIndex + 1;
            pFetchedAt.Value = fetchedAt;
            pRawJson.Value = row.RawJson;
            pFaccoTx.Value = DbText(row.FaccoTx);
            pCptypTx.Value = DbText(row.CptypTx);
            pZtypeTx.Value = DbText(row.ZtypeTx);
            pZtote.Value = DbText(row.Ztote);
            pZsort01.Value = DbText(row.Zsort01);
            pZsort02.Value = DbText(row.Zsort02);
            pZsort03.Value = DbText(row.Zsort03);
            pZsort04.Value = DbText(row.Zsort04);
            pCptyp.Value = DbText(row.Cptyp);
            pTsort.Value = DbText(row.Tsort);
            pZsort.Value = DbText(row.Zsort);
            pFacco.Value = DbText(row.Facco);
            pZtype.Value = DbText(row.Ztype);
            pChk.Value = DbText(row.Chk);
            for (int columnIndex = 1; columnIndex <= 14; columnIndex++)
                pColumns[columnIndex - 1].Value = row.GetCol(columnIndex) is double value
                    ? value
                    : DBNull.Value;
            command.ExecuteNonQuery();
        }
    }

    private static void InsertColumns(
        SqliteConnection connection,
        SqliteTransaction transaction,
        string queryDate,
        IReadOnlyList<FCostCorePartsColumn> columns)
    {
        if (columns.Count == 0)
            return;

        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText =
            """
            INSERT INTO MES072410Columns
                (QueryDate, ColIndex, Code, Header, PDate, Kind)
            VALUES
                (@date, @index, @code, @header, @pdate, @kind);
            """;
        SqliteParameter pDate = command.Parameters.Add("@date", SqliteType.Text);
        SqliteParameter pIndex = command.Parameters.Add("@index", SqliteType.Integer);
        SqliteParameter pCode = command.Parameters.Add("@code", SqliteType.Text);
        SqliteParameter pHeader = command.Parameters.Add("@header", SqliteType.Text);
        SqliteParameter pPdate = command.Parameters.Add("@pdate", SqliteType.Text);
        SqliteParameter pKind = command.Parameters.Add("@kind", SqliteType.Text);

        foreach (FCostCorePartsColumn column in columns)
        {
            pDate.Value = queryDate;
            pIndex.Value = column.ColIndex;
            pCode.Value = column.Code;
            pHeader.Value = column.Header;
            pPdate.Value = column.PDate;
            pKind.Value = column.Kind;
            command.ExecuteNonQuery();
        }
    }

    private static void UpsertPull(
        SqliteConnection connection,
        SqliteTransaction transaction,
        string queryDate,
        string status,
        string fetchedAt,
        int rowCount,
        int columnCount,
        string errorMessage)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText =
            """
            INSERT INTO MES072410Pulls
                (QueryDate, Status, FetchedAt, RowCount, ColumnCount, ErrorMessage)
            VALUES
                (@date, @status, @fetchedAt, @rowCount, @columnCount, @error)
            ON CONFLICT(QueryDate) DO UPDATE SET
                Status=excluded.Status,
                FetchedAt=excluded.FetchedAt,
                RowCount=excluded.RowCount,
                ColumnCount=excluded.ColumnCount,
                ErrorMessage=excluded.ErrorMessage;
            """;
        command.Parameters.AddWithValue("@date", queryDate);
        command.Parameters.AddWithValue("@status", status);
        command.Parameters.AddWithValue("@fetchedAt", fetchedAt);
        command.Parameters.AddWithValue("@rowCount", rowCount);
        command.Parameters.AddWithValue("@columnCount", columnCount);
        command.Parameters.AddWithValue("@error", errorMessage);
        command.ExecuteNonQuery();
    }

    private static void SaveFailure(string dbPath, DateTime queryDate, string message)
    {
        EnsureDatabase(dbPath);
        using var connection = OpenConnection(dbPath);
        using SqliteTransaction transaction = connection.BeginTransaction();
        UpsertPull(
            connection,
            transaction,
            queryDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            "ERROR",
            DateTime.Now.ToString("O", CultureInfo.InvariantCulture),
            0,
            0,
            message);
        transaction.Commit();
    }

    private static object DbText(string value)
        => string.IsNullOrEmpty(value) ? DBNull.Value : value;
}

public sealed record FCostCorePartsParseResult(
    IReadOnlyList<FCostCorePartsRow> Rows,
    IReadOnlyList<FCostCorePartsColumn> Columns,
    int ColumnCount);

public sealed record FCostCorePartsColumn(
    int ColIndex,
    string Code,
    string Header,
    string PDate,
    string Kind);

public sealed record FCostCorePartsKpiSnapshot(
    DateTime QueryDate,
    IReadOnlyList<FCostCorePartsKpiPeriod> Periods);

public sealed record FCostCorePartsKpiPeriod(
    int ColIndex,
    string Code,
    string Header,
    string PDate,
    string Kind,
    double? TotalCostUsd,
    double? TotalRatePercent);

public sealed class FCostCorePartsRow
{
    private readonly double?[] _columns = new double?[14];

    public string RawJson { get; set; } = string.Empty;
    public string FaccoTx { get; set; } = string.Empty;
    public string CptypTx { get; set; } = string.Empty;
    public string ZtypeTx { get; set; } = string.Empty;
    public string Ztote { get; set; } = string.Empty;
    public string Zsort01 { get; set; } = string.Empty;
    public string Zsort02 { get; set; } = string.Empty;
    public string Zsort03 { get; set; } = string.Empty;
    public string Zsort04 { get; set; } = string.Empty;
    public string Cptyp { get; set; } = string.Empty;
    public string Tsort { get; set; } = string.Empty;
    public string Zsort { get; set; } = string.Empty;
    public string Facco { get; set; } = string.Empty;
    public string Ztype { get; set; } = string.Empty;
    public string Chk { get; set; } = string.Empty;

    public double? Col0001 { get => GetCol(1); set => SetCol(1, value); }
    public double? Col0002 { get => GetCol(2); set => SetCol(2, value); }
    public double? Col0003 { get => GetCol(3); set => SetCol(3, value); }
    public double? Col0004 { get => GetCol(4); set => SetCol(4, value); }
    public double? Col0005 { get => GetCol(5); set => SetCol(5, value); }
    public double? Col0006 { get => GetCol(6); set => SetCol(6, value); }
    public double? Col0007 { get => GetCol(7); set => SetCol(7, value); }
    public double? Col0008 { get => GetCol(8); set => SetCol(8, value); }
    public double? Col0009 { get => GetCol(9); set => SetCol(9, value); }
    public double? Col0010 { get => GetCol(10); set => SetCol(10, value); }
    public double? Col0011 { get => GetCol(11); set => SetCol(11, value); }
    public double? Col0012 { get => GetCol(12); set => SetCol(12, value); }
    public double? Col0013 { get => GetCol(13); set => SetCol(13, value); }
    public double? Col0014 { get => GetCol(14); set => SetCol(14, value); }

    public double? GetCol(int index)
        => index is >= 1 and <= 14 ? _columns[index - 1] : null;

    public void SetCol(int index, double? value)
    {
        if (index is >= 1 and <= 14)
            _columns[index - 1] = value;
    }
}

public sealed class FCostCorePartsBackfillResult
{
    public string DbPath { get; set; } = string.Empty;
    public DateTime StartDate { get; set; }
    public DateTime EndDate { get; set; }
    public int AttemptedDays { get; set; }
    public int FetchedDays { get; set; }
    public int SkippedDays { get; set; }
    public int FailedDays { get; set; }
    public int TotalRows { get; set; }
    public List<string> Failures { get; } = [];
}

public sealed class FCostCorePartsStatus
{
    public string DbPath { get; set; } = string.Empty;
    public bool Exists { get; set; }
    public int PullCount { get; set; }
    public int SuccessCount { get; set; }
    public int FailedCount { get; set; }
    public int TotalRows { get; set; }
    public string FirstDate { get; set; } = string.Empty;
    public string LastDate { get; set; } = string.Empty;
}
