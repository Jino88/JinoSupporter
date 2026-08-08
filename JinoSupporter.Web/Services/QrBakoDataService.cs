using System.Data;
using System.Globalization;
using Microsoft.Data.SqlClient;

namespace JinoSupporter.Web.Services;

public sealed class QrBakoDataService(AppPathsService appPaths, AppActivityLogger activity)
{
    private const int FallbackMaxRows = 1000;
    private const int HardMaxRows = 20000;
    private const int FallbackTimeoutSeconds = 60;
    private const int MaxTimeoutSeconds = 300;

    private readonly AppPathsService _appPaths = appPaths;
    private readonly AppActivityLogger _activity = activity;

    public QrBakoDataConnectionSnapshot Snapshot => BuildSnapshot(_appPaths.Current);

    public bool IsConfigured
    {
        get
        {
            var s = Snapshot;
            return !string.IsNullOrWhiteSpace(s.Server)
                && !string.IsNullOrWhiteSpace(s.Database)
                && !string.IsNullOrWhiteSpace(s.UserId)
                && !string.IsNullOrWhiteSpace(s.Password);
        }
    }

    public int DefaultMaxRows => NormalizeMaxRows(_appPaths.Current.QrBakoDataDefaultMaxRows);

    public async Task TestConnectionAsync(CancellationToken cancellationToken = default)
    {
        var snapshot = RequireConfiguredSnapshot();
        await using var conn = new SqlConnection(BuildConnectionString(snapshot));
        await conn.OpenAsync(cancellationToken);

        await using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT TOP (1) 1 FROM [dbo].[BKTD] WITH (NOLOCK);";
        cmd.CommandTimeout = NormalizeTimeoutSeconds(snapshot.TimeoutSeconds);
        await cmd.ExecuteScalarAsync(cancellationToken);
    }

    public async Task<QrBakoDataResult> FetchAsync(
        QrBakoDataQuery query,
        CancellationToken cancellationToken = default)
    {
        var snapshot = RequireConfiguredSnapshot();
        int maxRows = NormalizeMaxRows(query.MaxRows <= 0 ? snapshot.DefaultMaxRows : query.MaxRows);

        await using var conn = new SqlConnection(BuildConnectionString(snapshot));
        await conn.OpenAsync(cancellationToken);

        List<string> tableColumns = await FetchTableColumnNamesAsync(
            conn,
            NormalizeTimeoutSeconds(snapshot.TimeoutSeconds),
            cancellationToken);
        string? sortColumn = ResolveTestTimeColumn(tableColumns);

        List<QrBakoDateSummary> dates = sortColumn is null
            ? []
            : await FetchDateSummariesAsync(
                conn,
                sortColumn,
                NormalizeTimeoutSeconds(snapshot.TimeoutSeconds),
                cancellationToken);
        DateOnly? selectedDate = query.TestDate ?? dates.FirstOrDefault()?.Date;

        long? totalRows = selectedDate is null || sortColumn is null
            ? 0
            : await TryCountRowsAsync(
                conn,
                sortColumn,
                selectedDate.Value,
                NormalizeTimeoutSeconds(snapshot.TimeoutSeconds),
                cancellationToken);

        await using var cmd = conn.CreateCommand();
        cmd.CommandText = sortColumn is null || selectedDate is null
            ? """
            SELECT TOP (@MaxRows) *
            FROM [dbo].[BKTD] WITH (NOLOCK);
            """
            : $"""
            SELECT TOP (@MaxRows) *
            FROM [dbo].[BKTD] WITH (NOLOCK)
            WHERE {QuoteSqlServerIdentifier(sortColumn)} >= @DateStart
              AND {QuoteSqlServerIdentifier(sortColumn)} < @DateEnd
            ORDER BY {QuoteSqlServerIdentifier(sortColumn)} DESC;
            """;
        cmd.CommandTimeout = NormalizeTimeoutSeconds(snapshot.TimeoutSeconds);
        cmd.Parameters.Add(new SqlParameter("@MaxRows", SqlDbType.Int) { Value = maxRows });
        if (sortColumn is not null && selectedDate is not null)
        {
            DateTime dateStart = selectedDate.Value.ToDateTime(TimeOnly.MinValue);
            cmd.Parameters.Add(new SqlParameter("@DateStart", SqlDbType.DateTime2) { Value = dateStart });
            cmd.Parameters.Add(new SqlParameter("@DateEnd", SqlDbType.DateTime2) { Value = dateStart.AddDays(1) });
        }

        var columns = new List<QrBakoDataColumn>();
        var rows = new List<QrBakoDataRow>();

        await using var reader = await cmd.ExecuteReaderAsync(
            CommandBehavior.SequentialAccess | CommandBehavior.SingleResult,
            cancellationToken);

        while (reader.FieldCount == 0 && await reader.NextResultAsync(cancellationToken))
        {
            // Skip statements that do not return a resultset.
        }

        if (reader.FieldCount > 0)
        {
            for (int i = 0; i < reader.FieldCount; i++)
            {
                string name = string.IsNullOrWhiteSpace(reader.GetName(i))
                    ? $"Column{i + 1}"
                    : reader.GetName(i).Trim();
                columns.Add(new QrBakoDataColumn(name, reader.GetDataTypeName(i)));
            }

            while (await reader.ReadAsync(cancellationToken))
            {
                var values = new List<string>(reader.FieldCount);
                for (int i = 0; i < reader.FieldCount; i++)
                    values.Add(FormatCell(reader, i));
                rows.Add(new QrBakoDataRow(values));
            }
        }

        _activity.Log(
            "QR BAKO DATA",
            $"Fetched dbo.BKTD date={selectedDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? "none"}, rows={rows.Count:N0}, maxRows={maxRows:N0}, total={(totalRows?.ToString("N0", CultureInfo.InvariantCulture) ?? "unknown")}");

        return new QrBakoDataResult
        {
            Server = snapshot.Server,
            Port = snapshot.Port,
            Database = snapshot.Database,
            TableName = "dbo.BKTD",
            FetchedAt = DateTime.Now,
            MaxRows = maxRows,
            TotalRows = totalRows,
            SortColumnName = sortColumn ?? string.Empty,
            SelectedDate = selectedDate,
            Dates = dates,
            WarningMessage = sortColumn is null
                ? "TestTime column was not found in dbo.BKTD, so the date list could not be created."
                : string.Empty,
            Columns = columns,
            Rows = rows,
        };
    }

    private static async Task<List<QrBakoDateSummary>> FetchDateSummariesAsync(
        SqlConnection conn,
        string sortColumn,
        int timeoutSeconds,
        CancellationToken cancellationToken)
    {
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = $"""
            SELECT CONVERT(date, {QuoteSqlServerIdentifier(sortColumn)}) AS [TestDate],
                   COUNT_BIG(1) AS [RecordCount]
            FROM [dbo].[BKTD] WITH (NOLOCK)
            WHERE {QuoteSqlServerIdentifier(sortColumn)} IS NOT NULL
            GROUP BY CONVERT(date, {QuoteSqlServerIdentifier(sortColumn)})
            ORDER BY [TestDate] DESC;
            """;
        cmd.CommandTimeout = timeoutSeconds;

        var dates = new List<QrBakoDateSummary>();
        await using var reader = await cmd.ExecuteReaderAsync(CommandBehavior.SequentialAccess, cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            DateTime date = reader.GetDateTime(0);
            long recordCount = reader.GetInt64(1);
            dates.Add(new QrBakoDateSummary(DateOnly.FromDateTime(date), recordCount));
        }

        return dates;
    }

    private static async Task<List<string>> FetchTableColumnNamesAsync(
        SqlConnection conn,
        int timeoutSeconds,
        CancellationToken cancellationToken)
    {
        try
        {
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT c.name
                FROM sys.columns AS c
                WHERE c.object_id = OBJECT_ID(N'[dbo].[BKTD]')
                ORDER BY c.column_id;
                """;
            cmd.CommandTimeout = timeoutSeconds;

            var columns = new List<string>();
            await using var reader = await cmd.ExecuteReaderAsync(CommandBehavior.SequentialAccess, cancellationToken);
            while (await reader.ReadAsync(cancellationToken))
            {
                if (!reader.IsDBNull(0))
                    columns.Add(reader.GetString(0));
            }
            return columns;
        }
        catch
        {
            return [];
        }
    }

    private static async Task<long?> TryCountRowsAsync(
        SqlConnection conn,
        string sortColumn,
        DateOnly date,
        int timeoutSeconds,
        CancellationToken cancellationToken)
    {
        try
        {
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = $"""
                SELECT COUNT_BIG(1)
                FROM [dbo].[BKTD] WITH (NOLOCK)
                WHERE {QuoteSqlServerIdentifier(sortColumn)} >= @DateStart
                  AND {QuoteSqlServerIdentifier(sortColumn)} < @DateEnd;
                """;
            cmd.CommandTimeout = timeoutSeconds;
            DateTime dateStart = date.ToDateTime(TimeOnly.MinValue);
            cmd.Parameters.Add(new SqlParameter("@DateStart", SqlDbType.DateTime2) { Value = dateStart });
            cmd.Parameters.Add(new SqlParameter("@DateEnd", SqlDbType.DateTime2) { Value = dateStart.AddDays(1) });
            object? value = await cmd.ExecuteScalarAsync(cancellationToken);
            return value is null or DBNull ? null : Convert.ToInt64(value, CultureInfo.InvariantCulture);
        }
        catch
        {
            return null;
        }
    }

    private QrBakoDataConnectionSnapshot RequireConfiguredSnapshot()
    {
        var snapshot = Snapshot;
        if (string.IsNullOrWhiteSpace(snapshot.Server))
            throw new InvalidOperationException("QR BAKO DB server is not configured.");
        if (string.IsNullOrWhiteSpace(snapshot.Database))
            throw new InvalidOperationException("QR BAKO DB database is not configured.");
        if (string.IsNullOrWhiteSpace(snapshot.UserId))
            throw new InvalidOperationException("QR BAKO DB user ID is not configured.");
        if (string.IsNullOrWhiteSpace(snapshot.Password))
            throw new InvalidOperationException("QR BAKO DB password is not configured on the server.");
        return snapshot;
    }

    private static QrBakoDataConnectionSnapshot BuildSnapshot(AppPathsConfig c) => new()
    {
        Server = c.QrBakoDataServer,
        Port = c.QrBakoDataPort,
        Database = c.QrBakoDataDatabase,
        UserId = c.QrBakoDataUserId,
        Password = c.QrBakoDataPassword,
        TimeoutSeconds = NormalizeTimeoutSeconds(c.QrBakoDataTimeoutSeconds),
        DefaultMaxRows = NormalizeMaxRows(c.QrBakoDataDefaultMaxRows),
        Encrypt = c.QrBakoDataEncrypt,
        TrustServerCertificate = c.QrBakoDataTrustServerCertificate,
    };

    private static string BuildConnectionString(QrBakoDataConnectionSnapshot c)
    {
        var builder = new SqlConnectionStringBuilder
        {
            DataSource = $"{c.Server.Trim()},{Math.Clamp(c.Port, 1, 65535)}",
            InitialCatalog = c.Database.Trim(),
            UserID = c.UserId.Trim(),
            Password = c.Password,
            ConnectTimeout = NormalizeTimeoutSeconds(c.TimeoutSeconds),
            Encrypt = c.Encrypt,
            TrustServerCertificate = c.TrustServerCertificate,
            MultipleActiveResultSets = false,
            ApplicationName = "JinoSupporter.Web QR BAKO DATA",
        };
        builder["ApplicationIntent"] = "ReadOnly";
        return builder.ConnectionString;
    }

    private static int NormalizeTimeoutSeconds(int value) =>
        Math.Clamp(value <= 0 ? FallbackTimeoutSeconds : value, 1, MaxTimeoutSeconds);

    private static int NormalizeMaxRows(int value) =>
        Math.Clamp(value <= 0 ? FallbackMaxRows : value, 1, HardMaxRows);

    private static string? ResolveTestTimeColumn(IReadOnlyList<string> columns)
    {
        if (columns.Count == 0)
            return null;

        string? exact = columns.FirstOrDefault(c =>
            string.Equals(c.Trim(), "Test Time", StringComparison.OrdinalIgnoreCase));
        if (!string.IsNullOrWhiteSpace(exact))
            return exact;

        string? normalized = columns.FirstOrDefault(c =>
            string.Equals(NormalizeColumnName(c), "testtime", StringComparison.OrdinalIgnoreCase));
        if (!string.IsNullOrWhiteSpace(normalized))
            return normalized;

        return columns.FirstOrDefault(c =>
        {
            string name = NormalizeColumnName(c);
            return name.Contains("test", StringComparison.OrdinalIgnoreCase)
                && name.Contains("time", StringComparison.OrdinalIgnoreCase);
        });
    }

    private static string NormalizeColumnName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        return new string(value
            .Where(char.IsLetterOrDigit)
            .Select(char.ToLowerInvariant)
            .ToArray());
    }

    private static string QuoteSqlServerIdentifier(string value) =>
        "[" + value.Replace("]", "]]", StringComparison.Ordinal) + "]";

    private static string FormatCell(SqlDataReader reader, int ordinal)
    {
        if (reader.IsDBNull(ordinal))
            return string.Empty;

        object value = reader.GetValue(ordinal);
        return value switch
        {
            DateTime dt => dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.CurrentCulture),
            DateTimeOffset dto => dto.ToString("yyyy-MM-dd HH:mm:ss zzz", CultureInfo.CurrentCulture),
            TimeSpan ts => ts.ToString(),
            byte[] bytes => FormatBytes(bytes),
            IFormattable f => f.ToString(null, CultureInfo.CurrentCulture) ?? string.Empty,
            _ => Convert.ToString(value, CultureInfo.CurrentCulture) ?? string.Empty,
        };
    }

    private static string FormatBytes(byte[] bytes)
    {
        if (bytes.Length == 0) return string.Empty;
        if (bytes.Length <= 32) return "0x" + Convert.ToHexString(bytes);
        return "0x" + Convert.ToHexString(bytes.AsSpan(0, 32)) + $"... ({bytes.Length:N0} bytes)";
    }
}

public sealed class QrBakoDataConnectionSnapshot
{
    public string Server { get; init; } = string.Empty;
    public int Port { get; init; } = 1430;
    public string Database { get; init; } = "TCDB";
    public string UserId { get; init; } = "TCDB";
    public string Password { get; init; } = string.Empty;
    public int TimeoutSeconds { get; init; } = 60;
    public int DefaultMaxRows { get; init; } = 1000;
    public bool Encrypt { get; init; }
    public bool TrustServerCertificate { get; init; } = true;
}

public sealed class QrBakoDataQuery
{
    public int MaxRows { get; init; } = 1000;
    public DateOnly? TestDate { get; init; }
}

public sealed class QrBakoDataResult
{
    public string Server { get; init; } = string.Empty;
    public int Port { get; init; }
    public string Database { get; init; } = string.Empty;
    public string TableName { get; init; } = "dbo.BKTD";
    public DateTime FetchedAt { get; init; }
    public int MaxRows { get; init; }
    public long? TotalRows { get; init; }
    public string SortColumnName { get; init; } = string.Empty;
    public DateOnly? SelectedDate { get; init; }
    public string WarningMessage { get; init; } = string.Empty;
    public List<QrBakoDateSummary> Dates { get; init; } = [];
    public List<QrBakoDataColumn> Columns { get; init; } = [];
    public List<QrBakoDataRow> Rows { get; init; } = [];
    public bool HitMaxRows => Rows.Count >= MaxRows;
}

public sealed record QrBakoDataColumn(string Name, string DataType);

public sealed record QrBakoDataRow(IReadOnlyList<string> Values);

public sealed record QrBakoDateSummary(DateOnly Date, long RowCount);
