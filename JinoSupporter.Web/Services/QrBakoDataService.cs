using System.Data;
using System.Globalization;
using Microsoft.Data.SqlClient;

namespace JinoSupporter.Web.Services;

public sealed class QrBakoDataService(AppPathsService appPaths, AppActivityLogger activity)
{
    private const int FallbackMaxRows = 1000;
    private const int HardMaxRows = 20000;
    private const int MaxSelectedDates = 512;
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
        string? productIdColumn = ResolveExactColumn(tableColumns, "ProductID");
        string? testResultColumn = ResolveExactColumn(tableColumns, "TestResult");

        List<QrBakoDateSummary> dates = sortColumn is null
            ? []
            : await FetchDateSummariesAsync(
                conn,
                sortColumn,
                NormalizeTimeoutSeconds(snapshot.TimeoutSeconds),
                cancellationToken);
        var availableDates = dates.Select(item => item.Date).ToHashSet();
        List<DateOnly> requestedDates = query.TestDates
            .Where(availableDates.Contains)
            .Distinct()
            .OrderByDescending(date => date)
            .ToList();
        if (requestedDates.Count == 0 && dates.FirstOrDefault() is { } latestDate)
            requestedDates.Add(latestDate.Date);

        bool dateSelectionLimited = requestedDates.Count > MaxSelectedDates;
        List<DateOnly> selectedDates = requestedDates
            .Take(MaxSelectedDates)
            .ToList();
        var selectedDateSet = selectedDates.ToHashSet();
        long? totalRows = sortColumn is null
            ? 0
            : dates.Where(item => selectedDateSet.Contains(item.Date)).Sum(item => item.RowCount);
        ProductAggregate? productAggregate = productIdColumn is null || sortColumn is null
            ? null
            : await FetchProductAggregateAsync(
                conn,
                sortColumn,
                productIdColumn,
                testResultColumn,
                selectedDates,
                NormalizeTimeoutSeconds(snapshot.TimeoutSeconds),
                cancellationToken);
        long? validTotalRows = productAggregate?.ValidRows ?? totalRows;
        long? excludedRows = totalRows is long allRows && validTotalRows is long validRows
            ? Math.Max(0, allRows - validRows)
            : null;
        QrBakoDataSummary? summary = productAggregate is not null && testResultColumn is not null
            ? new QrBakoDataSummary(productAggregate.InputCount, productAggregate.NgCount)
            : null;

        await using var cmd = conn.CreateCommand();
        var predicates = new List<string>();
        if (sortColumn is not null && selectedDates.Count > 0)
            predicates.Add(AddSelectedDatePredicate(cmd, sortColumn, selectedDates));
        if (productIdColumn is not null)
            predicates.Add(ProductIdLengthPredicate(productIdColumn));

        string whereClause = predicates.Count == 0
            ? string.Empty
            : "WHERE " + string.Join(" AND ", predicates);
        string orderClause = sortColumn is null
            ? string.Empty
            : $"ORDER BY {QuoteSqlServerIdentifier(sortColumn)} DESC";
        cmd.CommandText = $"""
            SELECT TOP (@MaxRows) *
            FROM [dbo].[BKTD] WITH (NOLOCK)
            {whereClause}
            {orderClause};
            """;
        cmd.CommandTimeout = NormalizeTimeoutSeconds(snapshot.TimeoutSeconds);
        cmd.Parameters.Add(new SqlParameter("@MaxRows", SqlDbType.Int) { Value = maxRows });

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

            int sortColumnOrdinal = sortColumn is null
                ? -1
                : columns.FindIndex(column =>
                    string.Equals(column.Name, sortColumn, StringComparison.OrdinalIgnoreCase));

            while (await reader.ReadAsync(cancellationToken))
            {
                var values = new List<string>(reader.FieldCount);
                DateTimeOffset? testTime = null;
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    object? rawValue = reader.IsDBNull(i) ? null : reader.GetValue(i);
                    values.Add(FormatCell(rawValue));
                    if (i == sortColumnOrdinal)
                        testTime = ConvertTestTime(rawValue);
                }
                rows.Add(new QrBakoDataRow(values, testTime));
            }


            if (sortColumnOrdinal >= 0)
                rows.Sort(static (left, right) => Nullable.Compare(right.TestTime, left.TestTime));
        }

        _activity.Log(
            "QR BAKO DATA",
            $"Fetched dbo.BKTD dates={SelectedDatesLogText(selectedDates)}, rows={rows.Count:N0}, maxRows={maxRows:N0}, validTotal={(validTotalRows?.ToString("N0", CultureInfo.InvariantCulture) ?? "unknown")}, excluded={(excludedRows?.ToString("N0", CultureInfo.InvariantCulture) ?? "unknown")}");

        var warnings = new List<string>();
        if (sortColumn is null)
            warnings.Add("TestTime column was not found in dbo.BKTD, so the date list could not be created.");
        if (productIdColumn is null)
            warnings.Add("ProductID column was not found in dbo.BKTD, so invalid QR values could not be excluded.");
        if (testResultColumn is null)
            warnings.Add("TestResult column was not found in dbo.BKTD, so the QR summary could not be calculated.");
        if (dateSelectionLimited)
            warnings.Add($"날짜는 한 번에 최대 {MaxSelectedDates:N0}개까지 조회하며, 가장 최근 날짜부터 적용했습니다.");

        return new QrBakoDataResult
        {
            Server = snapshot.Server,
            Port = snapshot.Port,
            Database = snapshot.Database,
            TableName = "dbo.BKTD",
            FetchedAt = DateTime.Now,
            MaxRows = maxRows,
            TotalRows = totalRows,
            ValidTotalRows = validTotalRows,
            ExcludedRows = excludedRows,
            SortColumnName = sortColumn ?? string.Empty,
            SelectedDates = selectedDates,
            Dates = dates,
            WarningMessage = string.Join(' ', warnings),
            Summary = summary,
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

    private static string AddSelectedDatePredicate(
        SqlCommand cmd,
        string sortColumn,
        IReadOnlyCollection<DateOnly> selectedDates)
    {
        List<DateRange> ranges = BuildDateRanges(selectedDates);
        var predicates = new List<string>(ranges.Count);
        for (int i = 0; i < ranges.Count; i++)
        {
            string startName = $"@DateStart{i}";
            string endName = $"@DateEnd{i}";
            predicates.Add($"({QuoteSqlServerIdentifier(sortColumn)} >= {startName} AND {QuoteSqlServerIdentifier(sortColumn)} < {endName})");
            cmd.Parameters.Add(new SqlParameter(startName, SqlDbType.DateTime2)
            {
                Value = ranges[i].Start.ToDateTime(TimeOnly.MinValue),
            });
            cmd.Parameters.Add(new SqlParameter(endName, SqlDbType.DateTime2)
            {
                Value = ranges[i].EndExclusive.ToDateTime(TimeOnly.MinValue),
            });
        }

        return "(" + string.Join(" OR ", predicates) + ")";
    }

    private static async Task<ProductAggregate> FetchProductAggregateAsync(
        SqlConnection conn,
        string sortColumn,
        string productIdColumn,
        string? testResultColumn,
        IReadOnlyCollection<DateOnly> selectedDates,
        int timeoutSeconds,
        CancellationToken cancellationToken)
    {
        await using var cmd = conn.CreateCommand();
        var predicates = new List<string>();
        if (selectedDates.Count > 0)
            predicates.Add(AddSelectedDatePredicate(cmd, sortColumn, selectedDates));
        predicates.Add(ProductIdLengthPredicate(productIdColumn));

        string cleanProductId = $"LTRIM(RTRIM({QuoteSqlServerIdentifier(productIdColumn)}))";
        string testTimeIdentifier = QuoteSqlServerIdentifier(sortColumn);
        // A retest replaces the prior verdict: use the newest TestTime row, matching the table's representative row.
        string ngExpression = testResultColumn is null
            ? "CONVERT(int, 0)"
            : $"CASE WHEN UPPER(LTRIM(RTRIM({QuoteSqlServerIdentifier(testResultColumn)}))) IN (N'NG', N'N/G', N'FAIL') THEN 1 ELSE 0 END";
        cmd.CommandText = $"""
            SELECT COALESCE(SUM([ProductRows]), 0) AS [ValidRows],
                   COUNT_BIG(1) AS [InputCount],
                   COALESCE(SUM(CONVERT(bigint, [IsNg])), 0) AS [NgCount]
            FROM
            (
                SELECT [ProductRows], [IsNg]
                FROM
                (
                    SELECT COUNT_BIG(1) OVER (PARTITION BY {cleanProductId}) AS [ProductRows],
                           {ngExpression} AS [IsNg],
                           ROW_NUMBER() OVER
                           (
                               PARTITION BY {cleanProductId}
                               ORDER BY {testTimeIdentifier} DESC
                           ) AS [LatestRank]
                    FROM [dbo].[BKTD] WITH (NOLOCK)
                    WHERE {string.Join(" AND ", predicates)}
                ) AS [RankedProductRows]
                WHERE [LatestRank] = 1
            ) AS [LatestProducts];
            """;
        cmd.CommandTimeout = timeoutSeconds;
        await using var reader = await cmd.ExecuteReaderAsync(CommandBehavior.SingleRow, cancellationToken);
        if (!await reader.ReadAsync(cancellationToken))
            return new ProductAggregate(0, 0, 0);

        return new ProductAggregate(
            reader.IsDBNull(0) ? 0 : reader.GetInt64(0),
            reader.IsDBNull(1) ? 0 : reader.GetInt64(1),
            reader.IsDBNull(2) ? 0 : reader.GetInt64(2));
    }

    private static string ProductIdLengthPredicate(string productIdColumn) =>
        $"LEN(LTRIM(RTRIM({QuoteSqlServerIdentifier(productIdColumn)}))) = 17";

    private static List<DateRange> BuildDateRanges(IReadOnlyCollection<DateOnly> selectedDates)
    {
        List<DateOnly> ordered = selectedDates.Distinct().OrderBy(date => date).ToList();
        var ranges = new List<DateRange>();
        if (ordered.Count == 0)
            return ranges;

        DateOnly rangeStart = ordered[0];
        DateOnly rangeEnd = ordered[0];
        for (int i = 1; i < ordered.Count; i++)
        {
            if (ordered[i] == rangeEnd.AddDays(1))
            {
                rangeEnd = ordered[i];
                continue;
            }

            ranges.Add(new DateRange(rangeStart, rangeEnd.AddDays(1)));
            rangeStart = ordered[i];
            rangeEnd = ordered[i];
        }

        ranges.Add(new DateRange(rangeStart, rangeEnd.AddDays(1)));
        return ranges;
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

    private static string? ResolveExactColumn(IReadOnlyList<string> columns, string expectedName)
    {
        string normalizedExpected = NormalizeColumnName(expectedName);
        return columns.FirstOrDefault(column =>
            string.Equals(NormalizeColumnName(column), normalizedExpected, StringComparison.OrdinalIgnoreCase));
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

    private static string SelectedDatesLogText(IReadOnlyList<DateOnly> dates)
    {
        if (dates.Count == 0)
            return "none";
        if (dates.Count == 1)
            return dates[0].ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        return $"{dates[^1]:yyyy-MM-dd}..{dates[0]:yyyy-MM-dd} ({dates.Count:N0})";
    }

    private static DateTimeOffset? ConvertTestTime(object? value)
    {
        return value switch
        {
            DateTime dt => new DateTimeOffset(dt),
            DateTimeOffset dto => dto,
            _ when DateTimeOffset.TryParse(
                Convert.ToString(value, CultureInfo.InvariantCulture),
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeLocal,
                out DateTimeOffset parsed) => parsed,
            _ => null,
        };
    }

    private static string FormatCell(object? value)
    {
        if (value is null or DBNull)
            return string.Empty;

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

    private readonly record struct DateRange(DateOnly Start, DateOnly EndExclusive);
    private sealed record ProductAggregate(long ValidRows, long InputCount, long NgCount);
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
    public IReadOnlyCollection<DateOnly> TestDates { get; init; } = [];
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
    public long? ValidTotalRows { get; init; }
    public long? ExcludedRows { get; init; }
    public string SortColumnName { get; init; } = string.Empty;
    public List<DateOnly> SelectedDates { get; init; } = [];
    public string WarningMessage { get; init; } = string.Empty;
    public QrBakoDataSummary? Summary { get; init; }
    public List<QrBakoDateSummary> Dates { get; init; } = [];
    public List<QrBakoDataColumn> Columns { get; init; } = [];
    public List<QrBakoDataRow> Rows { get; init; } = [];
    public bool HitMaxRows => Rows.Count >= MaxRows;
}

public sealed record QrBakoDataColumn(string Name, string DataType);

public sealed record QrBakoDataSummary(long InputCount, long NgCount)
{
    // This app reports defect rates in ppm; without any input the rate is undefined and the UI shows "-".
    public double? NgRatePpm => InputCount <= 0
        ? null
        : NgCount / (double)InputCount * 1_000_000d;
}

public sealed record QrBakoDataRow(IReadOnlyList<string> Values, DateTimeOffset? TestTime = null)
{
    public DateOnly? TestDate => TestTime is DateTimeOffset value
        ? DateOnly.FromDateTime(value.DateTime)
        : null;
}

public sealed record QrBakoDateSummary(DateOnly Date, long RowCount);
