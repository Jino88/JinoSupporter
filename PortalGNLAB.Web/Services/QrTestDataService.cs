using System.Data;
using System.Globalization;
using Microsoft.Data.SqlClient;

namespace PortalGNLAB.Web.Services;

/// <summary>
/// Reads one QR inspection table (BKTD / ALTD / GSTD) from TCDB. The connection is
/// shared across the three tables — only the table definition changes — and every
/// statement is a read: the connection asks for <c>ApplicationIntent=ReadOnly</c>.
/// </summary>
public sealed class QrTestDataService(AppPathsService appPaths, AppActivityLogger activity)
{
    // A test date runs to roughly 1,300 rows, so a 1,000-row cap truncated even a single day and
    // the counts on the page described only the newest part of it. 10,000 covers about a week of
    // one station in one read; the grid virtualizes what it renders, so the cost is the fetch and
    // the rows held per circuit, not the DOM.
    private const int FallbackMaxRows = 10000;
    private const int HardMaxRows = 20000;

    /// <summary>
    /// Ceiling for an <see cref="QrTestDataQuery.AllRows"/> read. The export is meant to carry
    /// the whole selection rather than the page the grid shows, but the workbook writer holds
    /// every cell in memory, so a stray 500-date selection must not decide how much the server
    /// allocates. At roughly 1,500 rows a test date this is about 60 dates in one file.
    /// </summary>
    public const int ExportMaxRows = 100_000;
    private const int MaxSelectedDates = 512;
    private const int FallbackTimeoutSeconds = 60;
    private const int MaxTimeoutSeconds = 300;

    private readonly AppPathsService _appPaths = appPaths;
    private readonly AppActivityLogger _activity = activity;

    public QrTestDataConnectionSnapshot Snapshot => BuildSnapshot(_appPaths.Current);

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

    /// <summary>Opens a read-only connection for callers that build their own SELECT (see <see cref="QrSummaryService"/>).</summary>
    internal async Task<SqlConnection> OpenConnectionAsync(CancellationToken cancellationToken)
    {
        var snapshot = RequireConfiguredSnapshot();
        var conn = new SqlConnection(BuildConnectionString(snapshot));
        try
        {
            await conn.OpenAsync(cancellationToken);
            return conn;
        }
        catch
        {
            await conn.DisposeAsync();
            throw;
        }
    }

    internal int CommandTimeoutSeconds => NormalizeTimeoutSeconds(Snapshot.TimeoutSeconds);

    public async Task TestConnectionAsync(
        QrTestTableDef? table = null,
        CancellationToken cancellationToken = default)
    {
        QrTestTableDef target = table ?? QrTestTables.Bako;
        var snapshot = RequireConfiguredSnapshot();
        await using var conn = new SqlConnection(BuildConnectionString(snapshot));
        await conn.OpenAsync(cancellationToken);

        await using var cmd = conn.CreateCommand();
        cmd.CommandText = $"SELECT TOP (1) 1 FROM {target.Table} WITH (NOLOCK);";
        cmd.CommandTimeout = NormalizeTimeoutSeconds(snapshot.TimeoutSeconds);
        await cmd.ExecuteScalarAsync(cancellationToken);
    }

    public async Task<QrTestDataResult> FetchAsync(
        QrTestDataQuery query,
        CancellationToken cancellationToken = default)
    {
        QrTestTableDef table = query.Table ?? QrTestTables.Bako;
        var snapshot = RequireConfiguredSnapshot();
        int maxRows = query.AllRows
            ? ExportMaxRows
            : NormalizeMaxRows(query.MaxRows <= 0 ? snapshot.DefaultMaxRows : query.MaxRows);

        await using var conn = new SqlConnection(BuildConnectionString(snapshot));
        await conn.OpenAsync(cancellationToken);

        List<string> tableColumns = await FetchTableColumnNamesAsync(
            conn,
            table,
            NormalizeTimeoutSeconds(snapshot.TimeoutSeconds),
            cancellationToken);
        string? sortColumn = ResolveTestTimeColumn(table, tableColumns);
        string? productIdColumn = ResolveExactColumn(tableColumns, "ProductID");
        string? testResultColumn = ResolveExactColumn(tableColumns, "TestResult");

        List<QrTestDateSummary> dates = sortColumn is null
            ? []
            : await FetchDateSummariesAsync(
                conn,
                table,
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
                table,
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

        // Both counts are already known here — the date list and the product aggregate run before
        // the row read — so the cap can be fitted to the selection instead of being guessed at.
        // One row of headroom keeps HitMaxRows false when everything fits, which is what tells the
        // caller its numbers cover the whole selection rather than a truncated page of it.
        if (query.AutoMaxRows && !query.AllRows && validTotalRows is long neededRows)
            maxRows = (int)Math.Clamp(neededRows + 1, 1, HardMaxRows);
        await using var cmd = conn.CreateCommand();
        var predicates = new List<string>();
        if (sortColumn is not null && selectedDates.Count > 0)
            predicates.Add(AddSelectedDatePredicate(cmd, sortColumn, selectedDates));
        if (productIdColumn is not null)
            predicates.Add(ProductIdLengthPredicate(table, productIdColumn));

        string whereClause = predicates.Count == 0
            ? string.Empty
            : "WHERE " + string.Join(" AND ", predicates);
        string orderClause = sortColumn is null
            ? string.Empty
            : $"ORDER BY {QuoteSqlServerIdentifier(sortColumn)} DESC";
        cmd.CommandText = $"""
            SELECT TOP (@MaxRows) *
            FROM {table.Table} WITH (NOLOCK)
            {whereClause}
            {orderClause};
            """;
        cmd.CommandTimeout = NormalizeTimeoutSeconds(snapshot.TimeoutSeconds);
        cmd.Parameters.Add(new SqlParameter("@MaxRows", SqlDbType.Int) { Value = maxRows });

        var columns = new List<QrTestDataColumn>();
        var rows = new List<QrTestDataRow>();

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
                columns.Add(new QrTestDataColumn(name, reader.GetDataTypeName(i)));
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
                rows.Add(new QrTestDataRow(values, testTime));
            }


            if (sortColumnOrdinal >= 0)
                rows.Sort(static (left, right) => Nullable.Compare(right.TestTime, left.TestTime));
        }

        QrTestDataSummary? summary = productAggregate is not null && testResultColumn is not null
            ? rows.Count >= maxRows
                ? BuildVisibleSummary(table, rows, columns, productIdColumn!, testResultColumn)
                : new QrTestDataSummary(productAggregate.InputCount, productAggregate.NgCount)
            : null;

        _activity.Log(
            table.MenuLabel,
            $"Fetched {table.DisplayTable} dates={SelectedDatesLogText(selectedDates)}, rows={rows.Count:N0}, maxRows={maxRows:N0}, validTotal={(validTotalRows?.ToString("N0", CultureInfo.InvariantCulture) ?? "unknown")}, excluded={(excludedRows?.ToString("N0", CultureInfo.InvariantCulture) ?? "unknown")}");

        var warnings = new List<string>();
        if (sortColumn is null)
            warnings.Add($"The test time column was not found in {table.DisplayTable}, so the date list could not be built.");
        if (productIdColumn is null)
            warnings.Add($"The ProductID column was not found in {table.DisplayTable}, so invalid QR values could not be excluded.");
        if (testResultColumn is null)
            warnings.Add($"The TestResult column was not found in {table.DisplayTable}, so the QR summary could not be calculated.");
        if (dateSelectionLimited)
            warnings.Add($"At most {MaxSelectedDates:N0} dates are queried at once; the most recent ones were applied.");

        return new QrTestDataResult
        {
            Table = table,
            Server = snapshot.Server,
            Port = snapshot.Port,
            Database = snapshot.Database,
            TableName = table.DisplayTable,
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

    private static async Task<List<QrTestDateSummary>> FetchDateSummariesAsync(
        SqlConnection conn,
        QrTestTableDef table,
        string sortColumn,
        int timeoutSeconds,
        CancellationToken cancellationToken)
    {
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = $"""
            SELECT CONVERT(date, {QuoteSqlServerIdentifier(sortColumn)}) AS [TestDate],
                   COUNT_BIG(1) AS [RecordCount]
            FROM {table.Table} WITH (NOLOCK)
            WHERE {QuoteSqlServerIdentifier(sortColumn)} IS NOT NULL
            GROUP BY CONVERT(date, {QuoteSqlServerIdentifier(sortColumn)})
            ORDER BY [TestDate] DESC;
            """;
        cmd.CommandTimeout = timeoutSeconds;

        var dates = new List<QrTestDateSummary>();
        await using var reader = await cmd.ExecuteReaderAsync(CommandBehavior.SequentialAccess, cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            DateTime date = reader.GetDateTime(0);
            long recordCount = reader.GetInt64(1);
            dates.Add(new QrTestDateSummary(DateOnly.FromDateTime(date), recordCount));
        }

        return dates;
    }

    private static async Task<List<string>> FetchTableColumnNamesAsync(
        SqlConnection conn,
        QrTestTableDef table,
        int timeoutSeconds,
        CancellationToken cancellationToken)
    {
        try
        {
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT c.name
                FROM sys.columns AS c
                WHERE c.object_id = OBJECT_ID(@TableName)
                ORDER BY c.column_id;
                """;
            cmd.CommandTimeout = timeoutSeconds;
            cmd.Parameters.Add(new SqlParameter("@TableName", SqlDbType.NVarChar, 256) { Value = table.Table });

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
        QrTestTableDef table,
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
        predicates.Add(ProductIdLengthPredicate(table, productIdColumn));

        string cleanProductId = $"LTRIM(RTRIM({QuoteSqlServerIdentifier(productIdColumn)}))";
        string testTimeIdentifier = QuoteSqlServerIdentifier(sortColumn);
        // A retest replaces the prior verdict: use the newest test time row, matching the table's representative row.
        string ngExpression = testResultColumn is null
            ? "CONVERT(int, 0)"
            : $"CASE WHEN UPPER(LTRIM(RTRIM({QuoteSqlServerIdentifier(testResultColumn)}))) IN ({AddFailValueParameters(cmd, table)}) THEN 1 ELSE 0 END";
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
                    FROM {table.Table} WITH (NOLOCK)
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

    /// <summary>Binds this table's fail verdicts as parameters and returns the <c>IN (...)</c> body.</summary>
    private static string AddFailValueParameters(SqlCommand cmd, QrTestTableDef table)
    {
        var names = new List<string>(table.FailValues.Count);
        for (int i = 0; i < table.FailValues.Count; i++)
        {
            string name = $"@Ng{i}";
            names.Add(name);
            cmd.Parameters.Add(new SqlParameter(name, SqlDbType.NVarChar, 40)
            {
                Value = table.FailValues[i].ToUpperInvariant(),
            });
        }
        return string.Join(", ", names);
    }

    private static string ProductIdLengthPredicate(QrTestTableDef table, string productIdColumn) =>
        $"LEN(LTRIM(RTRIM({QuoteSqlServerIdentifier(productIdColumn)}))) = {table.ProductIdLength}";

    private static QrTestDataSummary BuildVisibleSummary(
        QrTestTableDef table,
        IReadOnlyList<QrTestDataRow> rows,
        IReadOnlyList<QrTestDataColumn> columns,
        string productIdColumn,
        string testResultColumn)
    {
        int productIdOrdinal = FindColumnOrdinal(columns, productIdColumn);
        int testResultOrdinal = FindColumnOrdinal(columns, testResultColumn);
        if (productIdOrdinal < 0 || testResultOrdinal < 0)
            return new QrTestDataSummary(0, 0);

        var seenProductIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        long ngCount = 0;
        // Rows are newest-first, so the first occurrence is the table representative and the current retest verdict.
        // If TOP limits the table, summarize that same visible scope so the cards and representative rows stay aligned.
        foreach (QrTestDataRow row in rows)
        {
            string productId = row.Values[productIdOrdinal].Trim();
            if (!seenProductIds.Add(productId))
                continue;

            if (table.IsFail(row.Values[testResultOrdinal]))
                ngCount++;
        }

        return new QrTestDataSummary(seenProductIds.Count, ngCount);
    }

    private static int FindColumnOrdinal(IReadOnlyList<QrTestDataColumn> columns, string columnName)
    {
        for (int index = 0; index < columns.Count; index++)
        {
            if (string.Equals(columns[index].Name, columnName, StringComparison.OrdinalIgnoreCase))
                return index;
        }

        return -1;
    }

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

    private QrTestDataConnectionSnapshot RequireConfiguredSnapshot()
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

    private static QrTestDataConnectionSnapshot BuildSnapshot(AppPathsConfig c) => new()
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

    private static string BuildConnectionString(QrTestDataConnectionSnapshot c)
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
            ApplicationName = "PortalGNLAB.Web QR DATA",
        };
        builder["ApplicationIntent"] = "ReadOnly";
        return builder.ConnectionString;
    }

    private static int NormalizeTimeoutSeconds(int value) =>
        Math.Clamp(value <= 0 ? FallbackTimeoutSeconds : value, 1, MaxTimeoutSeconds);

    private static int NormalizeMaxRows(int value) =>
        Math.Clamp(value <= 0 ? FallbackMaxRows : value, 1, HardMaxRows);

    /// <summary>
    /// Picks the timestamp column: the table's declared candidates first (BKTD/ALTD use
    /// <c>TestTime</c>, GSTD uses <c>OTIME</c>), then a name heuristic so a renamed column
    /// still resolves instead of blanking the date list.
    /// </summary>
    private static string? ResolveTestTimeColumn(QrTestTableDef table, IReadOnlyList<string> columns)
    {
        if (columns.Count == 0)
            return null;

        foreach (string candidate in table.TimeColumns)
        {
            string? match = ResolveExactColumn(columns, candidate);
            if (!string.IsNullOrWhiteSpace(match))
                return match;
        }

        return columns.FirstOrDefault(c =>
        {
            string name = NormalizeColumnName(c);
            return name.Contains("time", StringComparison.OrdinalIgnoreCase)
                || name.Contains("date", StringComparison.OrdinalIgnoreCase);
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

public sealed class QrTestDataConnectionSnapshot
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

public sealed class QrTestDataQuery
{
    public QrTestTableDef? Table { get; init; }
    public int MaxRows { get; init; } = 1000;
    public IReadOnlyCollection<DateOnly> TestDates { get; init; } = [];

    /// <summary>
    /// Sizes the cap to the selected dates instead of taking <see cref="MaxRows"/>, so the grid
    /// holds the whole selection without the cap being re-picked every time the dates change.
    /// Still bounded by the service's hard ceiling; a selection past it is truncated as before.
    /// </summary>
    public bool AutoMaxRows { get; init; }

    /// <summary>
    /// Ignores <see cref="MaxRows"/> and reads every row of the selected dates, up to
    /// <see cref="QrTestDataService.ExportMaxRows"/>. For the workbook export, which has to
    /// carry the whole period rather than the page the grid happens to be showing.
    /// </summary>
    public bool AllRows { get; init; }
}

public sealed class QrTestDataResult
{
    public QrTestTableDef Table { get; init; } = QrTestTables.Bako;
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
    public QrTestDataSummary? Summary { get; init; }
    public List<QrTestDateSummary> Dates { get; init; } = [];
    public List<QrTestDataColumn> Columns { get; init; } = [];
    public List<QrTestDataRow> Rows { get; init; } = [];
    public bool HitMaxRows => Rows.Count >= MaxRows;
}

public sealed record QrTestDataColumn(string Name, string DataType);

public sealed record QrTestDataSummary(long InputCount, long NgCount)
{
    // This app reports defect rates in ppm; without any input the rate is undefined and the UI shows "-".
    public double? NgRatePpm => InputCount <= 0
        ? null
        : NgCount / (double)InputCount * 1_000_000d;
}

public sealed record QrTestDataRow(IReadOnlyList<string> Values, DateTimeOffset? TestTime = null)
{
    public DateOnly? TestDate => TestTime is DateTimeOffset value
        ? DateOnly.FromDateTime(value.DateTime)
        : null;
}

public sealed record QrTestDateSummary(DateOnly Date, long RowCount);
