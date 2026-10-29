using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace PortalGNLAB.Web.Services;

/// <summary>
/// MES050032 IPG defect response client/parser.
/// Authentication always starts from configured credentials and a new cookie session;
/// copied browser cookies and SSO/session tokens are not inputs to this service.
/// </summary>
public sealed class IpgDefectService(
    NgRateSettingsService settings,
    AppActivityLogger activity)
{
    private static readonly TimeSpan MutableCacheTtl = TimeSpan.FromMinutes(15);
    private readonly NgRateSettingsService _settings = settings;
    private readonly AppActivityLogger _activity = activity;

    public async Task<IpgDefectParseResult> FetchAsync(
        DateTime queryDate,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (TryReadSearchListCache(queryDate, "W", progress) is { } cached)
            return cached;

        using BmesWebSession session = await OpenSessionAsync(progress, cancellationToken);
        return await FetchSearchListAsync(
            session,
            queryDate,
            "W",
            progress,
            cancellationToken);
    }

    public async Task<IpgDefectKpiSnapshot> FetchKpiRangeAsync(
        DateTime start,
        DateTime end,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        start = start.Date;
        end = end.Date;
        if (end > DateTime.Today) end = DateTime.Today;
        if (start > end)
            throw new ArgumentException("Start date must be before or equal to end date.");

        var periodsByKey =
            new Dictionary<string, IpgDefectKpiPeriod>(StringComparer.OrdinalIgnoreCase);
        double? annualAveragePpm = null;
        IReadOnlyList<DateTime> weeklyQueryDates = BuildWeeklyQueryDates(start, end);

        BmesWebSession? session = null;
        async Task<BmesWebSession> GetSessionAsync()
        {
            session ??= await OpenSessionAsync(progress, cancellationToken);
            return session;
        }

        try
        {
            for (int index = 0; index < weeklyQueryDates.Count; index++)
            {
                DateTime queryDate = weeklyQueryDates[index];
                try
                {
                    Report(
                        progress,
                        $"MES050032 weekly [{index + 1}/{weeklyQueryDates.Count}]: {queryDate:yyyy-MM-dd}");
                    IpgDefectParseResult parsed = await FetchSearchListCachedAsync(
                        GetSessionAsync,
                        queryDate,
                        "W",
                        progress,
                        cancellationToken);
                    IpgDefectKpiSnapshot snapshot = BuildKpiSnapshot(parsed, "Week");
                    annualAveragePpm = snapshot.AnnualAveragePpm;
                    foreach (IpgDefectKpiPeriod period in snapshot.Periods)
                        periodsByKey[IpgPeriodIdentity(period)] = period;
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    Report(
                        progress,
                        $"[WARN] MES050032 weekly {queryDate:yyyy-MM-dd}: {ex.Message}");
                }
            }

            try
            {
                Report(progress, $"MES050032 monthly: {end:yyyy-MM-dd}");
                IpgDefectParseResult parsed = await FetchSearchListCachedAsync(
                    GetSessionAsync,
                    end,
                    "M",
                    progress,
                    cancellationToken);
                IpgDefectKpiSnapshot monthly = BuildKpiSnapshot(parsed, "Month");
                foreach (IpgDefectKpiPeriod period in monthly.Periods)
                    periodsByKey[IpgPeriodIdentity(period)] = period;
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                // A monthly response is optional until BMES actually returns trustworthy
                // header metadata. No inferred monthly value is generated on failure.
                Report(progress, "[WARN] MES050032 monthly: " + ex.Message);
            }
        }
        finally
        {
            session?.Dispose();
        }

        return new IpgDefectKpiSnapshot(
            periodsByKey.Values.ToList(),
            annualAveragePpm);
    }

    public static string BuildSearchListPayload(DateTime queryDate)
        => BuildSearchListPayload(queryDate, "W");

    public static string BuildSearchListPayload(DateTime queryDate, string periodCode)
        => BmesGateway.BuildIpgSearchListPayload(queryDate, periodCode);


    public static IpgDefectParseResult ParseSearchListJson(string json)
        => BmesGateway.ParseIpgSearchList(json);


    public static IpgDefectKpiSnapshot BuildKpiSnapshot(
        IpgDefectParseResult parsed,
        string periodKind = "Week")
    {
        ArgumentNullException.ThrowIfNull(parsed);
        string normalizedPeriodKind = NormalizePeriodKind(periodKind);

        IReadOnlyList<IpgDefectKpiPeriod> periods = parsed.Columns
            .OrderBy(column => column.ColIndex)
            .Select(column =>
            {
                double[] values = parsed.Rows
                    .Select(row => row.GetCol(column.ColIndex))
                    .Where(value => value.HasValue)
                    .Select(value => value!.Value)
                    .ToArray();
                return new IpgDefectKpiPeriod(
                    column.ColIndex,
                    column.ColumnName,
                    column.Header,
                    normalizedPeriodKind,
                    values.Length > 0 ? values.Average() : null);
            })
            .ToList();

        double[] annualValues = parsed.Rows
            .Where(row => row.Avegr.HasValue)
            .Select(row => row.Avegr!.Value)
            .ToArray();
        return new IpgDefectKpiSnapshot(
            periods,
            annualValues.Length > 0 ? annualValues.Average() : null);
    }

    public static IReadOnlyList<DateTime> BuildWeeklyQueryDates(DateTime start, DateTime end)
        => BuildQueryDates(start.Date, end.Date, intervalDays: 42);

    public static bool IsCachedSearchListReusable(
        DateTime queryDate,
        DateTime fetchedAt,
        DateTime now,
        TimeSpan mutableTtl)
    {
        if (mutableTtl <= TimeSpan.Zero)
            return false;

        DateTime queryDay = queryDate.Date;
        DateTime currentDay = now.Date;
        bool isRecentMutableQuery = queryDay >= currentDay.AddDays(-1);
        return !isRecentMutableQuery || fetchedAt >= now.Subtract(mutableTtl);
    }

    private async Task<BmesWebSession> OpenSessionAsync(
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        BmesLoginOutcome outcome = await BmesWebSession.OpenAsync(
            _settings.LoginId,
            _settings.Password,
            new BmesWebSessionOptions
            {
                Timeout = TimeSpan.FromSeconds(300),
                TokenMessage = "MES050032: reading login token",
                LoginMessage = "MES050032: logging in with configured credentials",
            },
            new Progress<string>(message => Report(progress, message)),
            cancellationToken);
        if (outcome.Session is not null)
            return outcome.Session;
        throw new InvalidOperationException(outcome.Failure switch
        {
            BmesLoginFailure.MissingCredentials => "BMES credentials are not configured.",
            BmesLoginFailure.Token => "Failed to read BMES login token.",
            _ => "BMES login failed.",
        });
    }


    private async Task<IpgDefectParseResult> FetchSearchListAsync(
        BmesWebSession session,
        DateTime queryDate,
        string periodCode,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        string normalizedPeriodCode = NormalizePeriodCode(periodCode);
        string json = await BmesGateway.FetchIpgDefectJson(
            session,
            queryDate,
            normalizedPeriodCode,
            cancellationToken);
        IpgDefectParseResult parsed = ParseSearchListJson(json);
        ReportParsed(progress, normalizedPeriodCode, parsed);
        TrySaveSearchListCache(queryDate, normalizedPeriodCode, json, parsed, progress);
        return parsed;
    }

    private async Task<IpgDefectParseResult> FetchSearchListCachedAsync(
        Func<Task<BmesWebSession>> sessionFactory,
        DateTime queryDate,
        string periodCode,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        string normalizedPeriodCode = NormalizePeriodCode(periodCode);
        if (TryReadSearchListCache(queryDate, normalizedPeriodCode, progress) is { } cached)
            return cached;

        BmesWebSession session = await sessionFactory();
        return await FetchSearchListAsync(
            session,
            queryDate,
            normalizedPeriodCode,
            progress,
            cancellationToken);
    }

    private void ReportParsed(
        IProgress<string>? progress,
        string normalizedPeriodCode,
        IpgDefectParseResult parsed)
    {
        Report(
            progress,
            $"MES050032 {normalizedPeriodCode}: parsed {parsed.Rows.Count:N0} rows · {parsed.Columns.Count:N0} headers");
    }

    private static List<DateTime> BuildQueryDates(
        DateTime start,
        DateTime end,
        int intervalDays)
    {
        start = start.Date;
        end = end.Date;

        var dates = new HashSet<DateTime> { start, end };
        for (DateTime cursor = end.Date;
             cursor > start.Date && cursor >= DateTime.MinValue.AddDays(intervalDays);
             cursor = cursor.AddDays(-intervalDays))
        {
            dates.Add(cursor);
        }
        return dates.OrderBy(date => date).ToList();
    }

    private static string NormalizePeriodCode(string periodCode)
        => BmesGateway.NormalizeIpgPeriodCode(periodCode);

    private static string NormalizePeriodKind(string periodKind) =>
        NormalizePeriodCode(periodKind) == "M" ? "Month" : "Week";

    private static string IpgPeriodIdentity(IpgDefectKpiPeriod period)
    {
        string periodCode = string.IsNullOrWhiteSpace(period.Header)
            ? period.ColumnName
            : period.Header;
        return period.Kind + "|" + periodCode;
    }

    private void Report(IProgress<string>? progress, string message)
    {
        _activity.Log("IpgDefect", message);
        progress?.Report(message);
    }

    private string GetRawDbPath()
        => Path.Combine(_settings.FCostDbSaveDirectory, "fcost_raw.db");

    private IpgDefectParseResult? TryReadSearchListCache(
        DateTime queryDate,
        string periodCode,
        IProgress<string>? progress)
    {
        string normalizedPeriodCode = NormalizePeriodCode(periodCode);
        string dbPath = GetRawDbPath();
        if (!File.Exists(dbPath))
            return null;

        try
        {
            using SqliteConnection connection = OpenCacheConnection(dbPath, readOnly: true);
            if (!TableExists(connection, "MES050032SearchListCache"))
                return null;

            string queryDateText = queryDate.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            using SqliteCommand command = connection.CreateCommand();
            command.CommandText =
                """
                SELECT FetchedAt, ResponseJson
                FROM MES050032SearchListCache
                WHERE QueryDate = @queryDate AND PeriodCode = @periodCode;
                """;
            command.Parameters.AddWithValue("@queryDate", queryDateText);
            command.Parameters.AddWithValue("@periodCode", normalizedPeriodCode);

            using SqliteDataReader reader = command.ExecuteReader();
            if (!reader.Read())
                return null;

            string fetchedAtText = ReadDbText(reader, 0);
            string json = ReadDbText(reader, 1);
            if (!DateTime.TryParse(
                    fetchedAtText,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.RoundtripKind,
                    out DateTime fetchedAt) ||
                !IsCachedSearchListReusable(
                    queryDate,
                    fetchedAt,
                    DateTime.Now,
                    MutableCacheTtl))
            {
                return null;
            }

            IpgDefectParseResult parsed = ParseSearchListJson(json);
            Report(
                progress,
                $"MES050032 {normalizedPeriodCode}: cache hit {queryDateText} (cached {fetchedAt:yyyy-MM-dd HH:mm})");
            ReportParsed(progress, normalizedPeriodCode, parsed);
            return parsed;
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            Report(
                progress,
                $"[WARN] MES050032 {normalizedPeriodCode} cache read {queryDate:yyyy-MM-dd}: {ex.Message}");
            return null;
        }
    }

    private void TrySaveSearchListCache(
        DateTime queryDate,
        string periodCode,
        string json,
        IpgDefectParseResult parsed,
        IProgress<string>? progress)
    {
        try
        {
            string saveDir = _settings.FCostDbSaveDirectory;
            Directory.CreateDirectory(saveDir);
            string dbPath = Path.Combine(saveDir, "fcost_raw.db");
            EnsureCacheDatabase(dbPath);

            using SqliteConnection connection = OpenCacheConnection(dbPath, readOnly: false);
            using SqliteCommand command = connection.CreateCommand();
            command.CommandText =
                """
                INSERT INTO MES050032SearchListCache
                    (QueryDate, PeriodCode, FetchedAt, RowCount, ColumnCount, ResponseJson)
                VALUES
                    (@queryDate, @periodCode, @fetchedAt, @rowCount, @columnCount, @responseJson)
                ON CONFLICT(QueryDate, PeriodCode) DO UPDATE SET
                    FetchedAt=excluded.FetchedAt,
                    RowCount=excluded.RowCount,
                    ColumnCount=excluded.ColumnCount,
                    ResponseJson=excluded.ResponseJson;
                """;
            command.Parameters.AddWithValue(
                "@queryDate",
                queryDate.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
            command.Parameters.AddWithValue("@periodCode", NormalizePeriodCode(periodCode));
            command.Parameters.AddWithValue(
                "@fetchedAt",
                DateTime.Now.ToString("O", CultureInfo.InvariantCulture));
            command.Parameters.AddWithValue("@rowCount", parsed.Rows.Count);
            command.Parameters.AddWithValue("@columnCount", parsed.ColumnCount);
            command.Parameters.AddWithValue("@responseJson", json);
            command.ExecuteNonQuery();
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            Report(
                progress,
                $"[WARN] MES050032 {periodCode} cache save {queryDate:yyyy-MM-dd}: {ex.Message}");
        }
    }

    private static void EnsureCacheDatabase(string dbPath)
    {
        using SqliteConnection connection = OpenCacheConnection(dbPath, readOnly: false);
        using SqliteCommand command = connection.CreateCommand();
        command.CommandText =
            """
            PRAGMA journal_mode=WAL;
            PRAGMA busy_timeout=5000;

            CREATE TABLE IF NOT EXISTS MES050032SearchListCache (
                QueryDate TEXT NOT NULL,
                PeriodCode TEXT NOT NULL,
                FetchedAt TEXT NOT NULL,
                RowCount INTEGER NOT NULL,
                ColumnCount INTEGER NOT NULL,
                ResponseJson TEXT NOT NULL,
                PRIMARY KEY (QueryDate, PeriodCode)
            );

            CREATE INDEX IF NOT EXISTS IX_MES050032SearchListCache_FetchedAt
                ON MES050032SearchListCache(FetchedAt);
            """;
        command.ExecuteNonQuery();
    }

    private static SqliteConnection OpenCacheConnection(string dbPath, bool readOnly)
    {
        var builder = new SqliteConnectionStringBuilder
        {
            DataSource = dbPath,
            Mode = readOnly ? SqliteOpenMode.ReadOnly : SqliteOpenMode.ReadWriteCreate,
        };
        var connection = new SqliteConnection(builder.ToString());
        connection.Open();
        using SqliteCommand pragma = connection.CreateCommand();
        pragma.CommandText = "PRAGMA busy_timeout=5000;";
        pragma.ExecuteNonQuery();
        return connection;
    }

    private static bool TableExists(SqliteConnection connection, string tableName)
    {
        using SqliteCommand command = connection.CreateCommand();
        command.CommandText =
            "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@name;";
        command.Parameters.AddWithValue("@name", tableName);
        return Convert.ToInt64(command.ExecuteScalar() ?? 0L) > 0;
    }

    private static string ReadDbText(SqliteDataReader reader, int ordinal) =>
        reader.IsDBNull(ordinal) ? string.Empty : reader.GetString(ordinal);
}

public sealed record IpgDefectParseResult(
    IReadOnlyList<IpgDefectRow> Rows,
    IReadOnlyList<IpgDefectColumn> Columns,
    int ColumnCount);

public sealed record IpgDefectColumn(
    int ColIndex,
    string ColumnName,
    string Header);

public sealed record IpgDefectKpiSnapshot(
    IReadOnlyList<IpgDefectKpiPeriod> Periods,
    double? AnnualAveragePpm);

public sealed record IpgDefectKpiPeriod(
    int ColIndex,
    string ColumnName,
    string Header,
    string Kind,
    double? AveragePpm);

public sealed class IpgDefectRow
{
    private readonly double?[] _columns = new double?[14];

    public string RawJson { get; set; } = string.Empty;
    public string Cptyp { get; set; } = string.Empty;
    public string CptypTx { get; set; } = string.Empty;
    public string FaccoTx { get; set; } = string.Empty;
    public string ModnoTx { get; set; } = string.Empty;
    public string ZlrtpTx { get; set; } = string.Empty;
    public string MatnrTx { get; set; } = string.Empty;
    public string Facco { get; set; } = string.Empty;
    public string Modno { get; set; } = string.Empty;
    public string Zsort { get; set; } = string.Empty;
    public string Zlrtp { get; set; } = string.Empty;
    public double? Avegr { get; set; }

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
