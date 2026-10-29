using System.Data;
using System.Globalization;
using Microsoft.Data.SqlClient;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Joins the three QR inspection tables into one row per ProductID.
/// Each station contributes the values from its own latest test in the selected range —
/// a retest supersedes the earlier attempt, so only the newest row per ProductID counts.
/// Read-only: this service issues SELECT statements and nothing else.
/// </summary>
public sealed class QrSummaryService(QrTestDataService qrData, AppActivityLogger activity)
{
    private const int FallbackMaxRows = 2000;
    private const int HardMaxRows = 20000;

    /// <summary>
    /// Ceiling in QR for an <see cref="QrSummaryQuery.AllRows"/> read. Each QR can carry up to
    /// <see cref="MaxRoundsPerProduct"/> rounds and the workbook writer holds every cell in
    /// memory, so the export is bounded even when a wide date range is asked for.
    /// </summary>
    public const int ExportMaxRows = 50_000;

    /// <summary>
    /// How many test rounds per ProductID the QR group column can unfold. Round 1 is the
    /// newest test of each station; a QR retested more often than this keeps only the newest
    /// rounds, so one pathological ProductID cannot flood the grid.
    /// </summary>
    private const int MaxRoundsPerProduct = 20;

    private readonly QrTestDataService _qrData = qrData;
    private readonly AppActivityLogger _activity = activity;

    public bool IsConfigured => _qrData.IsConfigured;

    public QrTestDataConnectionSnapshot Snapshot => _qrData.Snapshot;

    public async Task<QrSummaryResult> FetchAsync(
        QrSummaryQuery query,
        CancellationToken cancellationToken = default)
    {
        int maxRows = query.AllRows
            ? ExportMaxRows
            : Math.Clamp(query.MaxRows <= 0 ? FallbackMaxRows : query.MaxRows, 1, HardMaxRows);

        await using SqlConnection conn = await _qrData.OpenConnectionAsync(cancellationToken);
        int timeout = _qrData.CommandTimeoutSeconds;

        QrSummaryBounds bounds = await FetchBoundsAsync(conn, timeout, cancellationToken);

        // With no data at all there is nothing to range over; hand back an empty result
        // rather than querying a nonsense window.
        if (bounds.MaxDate is not DateOnly maxDate)
        {
            return new QrSummaryResult
            {
                From = query.From,
                To = query.To,
                MaxRows = maxRows,
                Bounds = bounds,
                WarningMessage = "None of the three tables hold any test data to query.",
            };
        }

        DateOnly from = query.From == default ? maxDate.AddDays(-6) : query.From;
        DateOnly to = query.To == default ? maxDate : query.To;
        if (to < from)
            (from, to) = (to, from);

        var rows = new List<QrSummaryRow>();
        long totalProducts = 0;

        await using (SqlCommand cmd = conn.CreateCommand())
        {
            cmd.CommandText = SummarySql;
            cmd.CommandTimeout = timeout;
            cmd.Parameters.Add(new SqlParameter("@From", SqlDbType.DateTime2)
            {
                Value = from.ToDateTime(TimeOnly.MinValue),
            });
            cmd.Parameters.Add(new SqlParameter("@ToExclusive", SqlDbType.DateTime2)
            {
                Value = to.AddDays(1).ToDateTime(TimeOnly.MinValue),
            });
            cmd.Parameters.Add(new SqlParameter("@MaxRows", SqlDbType.Int) { Value = maxRows });
            cmd.Parameters.Add(new SqlParameter("@MaxRounds", SqlDbType.Int) { Value = MaxRoundsPerProduct });
            cmd.Parameters.Add(new SqlParameter("@QrLength", SqlDbType.Int)
            {
                Value = QrTestTables.Bako.ProductIdLength,
            });

            await using SqlDataReader reader = await cmd.ExecuteReaderAsync(
                CommandBehavior.SingleResult,
                cancellationToken);

            // Resolve ordinals once; the row loop runs up to HardMaxRows times.
            var o = new Ordinals(reader);

            // Rows arrive grouped by ProductID and ordered newest round first, so the first
            // row of each ProductID is the one the grid shows and the rest fold under it.
            QrSummaryRow? product = null;

            while (await reader.ReadAsync(cancellationToken))
            {
                if (totalProducts == 0)
                    totalProducts = reader.GetInt64(o.TotalProducts);

                var row = new QrSummaryRow
                {
                    ProductId = ReadText(reader, o.ProductId),
                    Round = reader.GetInt32(o.Round),
                    Airleak = BuildCell(
                        QrTestTables.Airleak,
                        ReadText(reader, o.AirleakResult),
                        ReadTime(reader, o.AirleakTime),
                        [Item("Value", ReadText(reader, o.AirleakValue), ReadText(reader, o.AirleakResult), QrTestTables.Airleak)]),
                    Bako = BuildCell(
                        QrTestTables.Bako,
                        ReadText(reader, o.BakoResult),
                        ReadTime(reader, o.BakoTime),
                        [
                            Item("SPL1", ReadText(reader, o.BakoSpl1), ReadText(reader, o.BakoSpl1Result), QrTestTables.Bako),
                            Item("SPL2", ReadText(reader, o.BakoSpl2), ReadText(reader, o.BakoSpl2Result), QrTestTables.Bako),
                            Item("THD", ReadText(reader, o.BakoThd), ReadText(reader, o.BakoThdResult), QrTestTables.Bako),
                            Item("ImpLimit", ReadText(reader, o.BakoImpLimit), ReadText(reader, o.BakoImpLimitResult), QrTestTables.Bako),
                            Item("Imp2kHz", ReadText(reader, o.BakoImp2kHz), ReadText(reader, o.BakoImp2kHzResult), QrTestTables.Bako),
                        ]),
                    Gauss = BuildCell(
                        QrTestTables.Gauss,
                        ReadText(reader, o.GaussResult),
                        ReadTime(reader, o.GaussTime),
                        [
                            // GSTD carries no per-value verdict, so both readings inherit the row's TestResult.
                            Item("Value", ReadText(reader, o.GaussValue), ReadText(reader, o.GaussResult), QrTestTables.Gauss),
                            Item("Value 2", ReadText(reader, o.GaussValue2), ReadText(reader, o.GaussResult), QrTestTables.Gauss),
                        ]),
                };

                if (product is null || !string.Equals(product.ProductId, row.ProductId, StringComparison.Ordinal))
                {
                    product = row;
                    rows.Add(row);
                }
                else
                {
                    product.Older.Add(row);
                }
            }
        }

        _activity.Log(
            "QR SUMMARY",
            $"Fetched ALTD/BKTD/GSTD {from:yyyy-MM-dd}..{to:yyyy-MM-dd}, products={rows.Count:N0}/{totalProducts:N0}, maxRows={maxRows:N0}");

        return new QrSummaryResult
        {
            From = from,
            To = to,
            MaxRows = maxRows,
            Bounds = bounds,
            TotalProducts = totalProducts,
            Rows = rows,
            FetchedAt = DateTime.Now,
        };
    }

    private static async Task<QrSummaryBounds> FetchBoundsAsync(
        SqlConnection conn,
        int timeout,
        CancellationToken cancellationToken)
    {
        await using SqlCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT (SELECT MAX([TestTime]) FROM [dbo].[ALTD] WITH (NOLOCK)) AS [AirleakMax],
                   (SELECT MAX([TestTime]) FROM [dbo].[BKTD] WITH (NOLOCK)) AS [BakoMax],
                   (SELECT MAX([OTIME])    FROM [dbo].[GSTD] WITH (NOLOCK)) AS [GaussMax];
            """;
        cmd.CommandTimeout = timeout;

        await using SqlDataReader reader = await cmd.ExecuteReaderAsync(CommandBehavior.SingleRow, cancellationToken);
        if (!await reader.ReadAsync(cancellationToken))
            return new QrSummaryBounds(null, null, null);

        return new QrSummaryBounds(
            ReadDate(reader, 0),
            ReadDate(reader, 1),
            ReadDate(reader, 2));
    }

    /// <summary>
    /// The three tables are normalised into one shape, ranked once, then pivoted with
    /// conditional aggregation.
    ///
    /// The obvious shape — a CTE per table outer-joined onto the union of their keys —
    /// costs 52s over a 30-day window here, because the optimiser re-evaluates each
    /// windowed CTE per probe row. Scanning once and grouping instead runs in ~36ms.
    /// Keep it single-pass.
    ///
    /// One row comes back per (ProductID, Round): round 1 pairs each station's newest test,
    /// round 2 the second newest, and so on. The grid shows round 1 and hides the rest behind
    /// the QR group column, so a retested QR no longer silently drops its earlier attempts.
    /// </summary>
    private const string SummarySql = """
        WITH [Raw] AS
        (
            SELECT CONVERT(tinyint, 1)                             AS [Station],
                   LTRIM(RTRIM(a.[ProductID]))                     AS [ProductID],
                   a.[TestTime]                                    AS [TestTime],
                   CONVERT(bigint, a.[id])                         AS [Tie],
                   CONVERT(nvarchar(80), a.[TestResult])           AS [Result],
                   -- ALTD often leaves the numeric pressure null and puts the reading in description ("1.2 Pa/s").
                   CONVERT(nvarchar(80), CASE WHEN a.[leak_pressure_val] IS NULL
                                              THEN a.[description]
                                              ELSE CONCAT(CONVERT(nvarchar(40), a.[leak_pressure_val]), N' ', a.[leak_pressure_unit])
                                         END)                      AS [V1],
                   CONVERT(nvarchar(80), a.[TestResult])           AS [R1],
                   CONVERT(nvarchar(80), NULL)                     AS [V2],
                   CONVERT(nvarchar(80), NULL)                     AS [R2],
                   CONVERT(nvarchar(80), NULL)                     AS [V3],
                   CONVERT(nvarchar(80), NULL)                     AS [R3],
                   CONVERT(nvarchar(80), NULL)                     AS [V4],
                   CONVERT(nvarchar(80), NULL)                     AS [R4],
                   CONVERT(nvarchar(80), NULL)                     AS [V5],
                   CONVERT(nvarchar(80), NULL)                     AS [R5]
            FROM [dbo].[ALTD] AS a WITH (NOLOCK)
            WHERE a.[TestTime] >= @From
              AND a.[TestTime] <  @ToExclusive
              AND LEN(LTRIM(RTRIM(a.[ProductID]))) = @QrLength

            UNION ALL

            SELECT CONVERT(tinyint, 2),
                   LTRIM(RTRIM(b.[ProductID])),
                   b.[TestTime],
                   CONVERT(bigint, b.[Id]),
                   CONVERT(nvarchar(80), b.[TestResult]),
                   CONVERT(nvarchar(80), b.[SPL1_Value]),     CONVERT(nvarchar(80), b.[SPL1_Result]),
                   CONVERT(nvarchar(80), b.[SPL2_Value]),     CONVERT(nvarchar(80), b.[SPL2_Result]),
                   CONVERT(nvarchar(80), b.[THD_Value]),      CONVERT(nvarchar(80), b.[THD_Result]),
                   CONVERT(nvarchar(80), b.[ImpLimit_Value]), CONVERT(nvarchar(80), b.[ImpLimit_Result]),
                   CONVERT(nvarchar(80), b.[Imp2kHz_Value]),  CONVERT(nvarchar(80), b.[Imp2kHz_Result])
            FROM [dbo].[BKTD] AS b WITH (NOLOCK)
            WHERE b.[TestTime] >= @From
              AND b.[TestTime] <  @ToExclusive
              AND LEN(LTRIM(RTRIM(b.[ProductID]))) = @QrLength

            UNION ALL

            SELECT CONVERT(tinyint, 3),
                   LTRIM(RTRIM(g.[ProductID])),
                   g.[OTIME],
                   CONVERT(bigint, 0),
                   CONVERT(nvarchar(80), g.[TestResult]),
                   CONVERT(nvarchar(80), g.[Value]),  CONVERT(nvarchar(80), g.[TestResult]),
                   CONVERT(nvarchar(80), g.[Value2]), CONVERT(nvarchar(80), g.[TestResult]),
                   CONVERT(nvarchar(80), NULL), CONVERT(nvarchar(80), NULL),
                   CONVERT(nvarchar(80), NULL), CONVERT(nvarchar(80), NULL),
                   CONVERT(nvarchar(80), NULL), CONVERT(nvarchar(80), NULL)
            FROM [dbo].[GSTD] AS g WITH (NOLOCK)
            WHERE g.[OTIME] >= @From
              AND g.[OTIME] <  @ToExclusive
              AND LEN(LTRIM(RTRIM(g.[ProductID]))) = @QrLength
        ),
        [Ranked] AS
        (
            SELECT [Raw].*,
                   ROW_NUMBER() OVER (PARTITION BY [ProductID], [Station]
                                      ORDER BY [TestTime] DESC, [Tie] DESC) AS [Round]
            FROM [Raw]
        ),
        [Pivoted] AS
        (
            SELECT [ProductID]                                      AS [ProductID],
                   [Round]                                          AS [Round],
                   -- Window over the aggregate: the QR's newest time across all of its rounds,
                   -- so every round of one ProductID sorts together at the QR's own position.
                   MAX(MAX([TestTime])) OVER (PARTITION BY [ProductID]) AS [ProductLatest],
                   MAX(CASE WHEN [Station] = 1 THEN [Result]   END) AS [AirleakResult],
                   MAX(CASE WHEN [Station] = 1 THEN [TestTime] END) AS [AirleakTime],
                   MAX(CASE WHEN [Station] = 1 THEN [V1]       END) AS [AirleakValue],
                   MAX(CASE WHEN [Station] = 2 THEN [Result]   END) AS [BakoResult],
                   MAX(CASE WHEN [Station] = 2 THEN [TestTime] END) AS [BakoTime],
                   MAX(CASE WHEN [Station] = 2 THEN [V1]       END) AS [BakoSpl1],
                   MAX(CASE WHEN [Station] = 2 THEN [R1]       END) AS [BakoSpl1Result],
                   MAX(CASE WHEN [Station] = 2 THEN [V2]       END) AS [BakoSpl2],
                   MAX(CASE WHEN [Station] = 2 THEN [R2]       END) AS [BakoSpl2Result],
                   MAX(CASE WHEN [Station] = 2 THEN [V3]       END) AS [BakoThd],
                   MAX(CASE WHEN [Station] = 2 THEN [R3]       END) AS [BakoThdResult],
                   MAX(CASE WHEN [Station] = 2 THEN [V4]       END) AS [BakoImpLimit],
                   MAX(CASE WHEN [Station] = 2 THEN [R4]       END) AS [BakoImpLimitResult],
                   MAX(CASE WHEN [Station] = 2 THEN [V5]       END) AS [BakoImp2kHz],
                   MAX(CASE WHEN [Station] = 2 THEN [R5]       END) AS [BakoImp2kHzResult],
                   MAX(CASE WHEN [Station] = 3 THEN [Result]   END) AS [GaussResult],
                   MAX(CASE WHEN [Station] = 3 THEN [TestTime] END) AS [GaussTime],
                   MAX(CASE WHEN [Station] = 3 THEN [V1]       END) AS [GaussValue],
                   MAX(CASE WHEN [Station] = 3 THEN [V2]       END) AS [GaussValue2]
            FROM [Ranked]
            WHERE [Round] <= @MaxRounds
            GROUP BY [ProductID], [Round]
        ),
        [Ordered] AS
        (
            SELECT [Pivoted].*,
                   DENSE_RANK() OVER (ORDER BY [ProductLatest] DESC, [ProductID] ASC)  AS [ProductRank],
                   -- Same total ordering reversed. TOP (@MaxRows) cannot report how many QR
                   -- were left behind, but for a dense rank and its mirror the two always sum
                   -- to the distinct QR count plus one — no second scan of the three tables.
                   DENSE_RANK() OVER (ORDER BY [ProductLatest] ASC,  [ProductID] DESC) AS [ProductRankRev]
            FROM [Pivoted]
        )
        SELECT [ProductID],
               -- ROW_NUMBER() is bigint; the reader wants a plain round counter.
               CONVERT(int, [Round])                                 AS [Round],
               CONVERT(bigint, [ProductRank] + [ProductRankRev] - 1) AS [TotalProducts],
               [AirleakResult], [AirleakTime], [AirleakValue],
               [BakoResult], [BakoTime],
               [BakoSpl1], [BakoSpl1Result], [BakoSpl2], [BakoSpl2Result],
               [BakoThd], [BakoThdResult], [BakoImpLimit], [BakoImpLimitResult],
               [BakoImp2kHz], [BakoImp2kHzResult],
               [GaussResult], [GaussTime], [GaussValue], [GaussValue2]
        FROM [Ordered]
        WHERE [ProductRank] <= @MaxRows
        ORDER BY [ProductRank] ASC, [Round] ASC;
        """;

    private static QrSummaryItem Item(string label, string value, string result, QrTestTableDef table) =>
        new(label, value.Trim(), Classify(table, result));

    private static QrVerdict Classify(QrTestTableDef table, string result) =>
        table.IsFail(result) ? QrVerdict.Fail
        : table.IsPass(result) ? QrVerdict.Pass
        : QrVerdict.Unknown;

    private static QrSummaryCell? BuildCell(
        QrTestTableDef table,
        string result,
        DateTimeOffset? testTime,
        IReadOnlyList<QrSummaryItem> items)
    {
        if (testTime is null && string.IsNullOrWhiteSpace(result))
            return null;

        return new QrSummaryCell(table.Station, Classify(table, result), result.Trim(), testTime, items);
    }

    private static string ReadText(SqlDataReader reader, int ordinal)
    {
        if (reader.IsDBNull(ordinal))
            return string.Empty;

        object value = reader.GetValue(ordinal);
        return value switch
        {
            string s => s,
            DateTime dt => dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.CurrentCulture),
            IFormattable f => f.ToString(null, CultureInfo.InvariantCulture) ?? string.Empty,
            _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty,
        };
    }

    private static DateTimeOffset? ReadTime(SqlDataReader reader, int ordinal) =>
        reader.IsDBNull(ordinal) ? null : new DateTimeOffset(reader.GetDateTime(ordinal));

    private static DateOnly? ReadDate(SqlDataReader reader, int ordinal) =>
        reader.IsDBNull(ordinal) ? null : DateOnly.FromDateTime(reader.GetDateTime(ordinal));

    /// <summary>Column ordinals resolved once per query instead of per row.</summary>
    private sealed class Ordinals(SqlDataReader reader)
    {
        public int ProductId { get; } = reader.GetOrdinal("ProductID");
        public int Round { get; } = reader.GetOrdinal("Round");
        public int TotalProducts { get; } = reader.GetOrdinal("TotalProducts");
        public int AirleakResult { get; } = reader.GetOrdinal("AirleakResult");
        public int AirleakTime { get; } = reader.GetOrdinal("AirleakTime");
        public int AirleakValue { get; } = reader.GetOrdinal("AirleakValue");
        public int BakoResult { get; } = reader.GetOrdinal("BakoResult");
        public int BakoTime { get; } = reader.GetOrdinal("BakoTime");
        public int BakoSpl1 { get; } = reader.GetOrdinal("BakoSpl1");
        public int BakoSpl1Result { get; } = reader.GetOrdinal("BakoSpl1Result");
        public int BakoSpl2 { get; } = reader.GetOrdinal("BakoSpl2");
        public int BakoSpl2Result { get; } = reader.GetOrdinal("BakoSpl2Result");
        public int BakoThd { get; } = reader.GetOrdinal("BakoThd");
        public int BakoThdResult { get; } = reader.GetOrdinal("BakoThdResult");
        public int BakoImpLimit { get; } = reader.GetOrdinal("BakoImpLimit");
        public int BakoImpLimitResult { get; } = reader.GetOrdinal("BakoImpLimitResult");
        public int BakoImp2kHz { get; } = reader.GetOrdinal("BakoImp2kHz");
        public int BakoImp2kHzResult { get; } = reader.GetOrdinal("BakoImp2kHzResult");
        public int GaussResult { get; } = reader.GetOrdinal("GaussResult");
        public int GaussTime { get; } = reader.GetOrdinal("GaussTime");
        public int GaussValue { get; } = reader.GetOrdinal("GaussValue");
        public int GaussValue2 { get; } = reader.GetOrdinal("GaussValue2");
    }
}

public enum QrVerdict
{
    Unknown,
    Pass,
    Fail,
}

public sealed class QrSummaryQuery
{
    /// <summary>Inclusive start of the window; <c>default</c> means "latest date minus 6 days".</summary>
    public DateOnly From { get; init; }

    /// <summary>Inclusive end of the window; <c>default</c> means "latest date with data".</summary>
    public DateOnly To { get; init; }

    public int MaxRows { get; init; } = 2000;

    /// <summary>
    /// Ignores <see cref="MaxRows"/> and reads every QR in the window, up to
    /// <see cref="QrSummaryService.ExportMaxRows"/>. For the workbook export, which has to carry
    /// the whole period rather than the page the grid happens to be showing.
    /// </summary>
    public bool AllRows { get; init; }
}

public sealed record QrSummaryBounds(DateOnly? AirleakMax, DateOnly? BakoMax, DateOnly? GaussMax)
{
    /// <summary>Newest date any of the three stations has data for, or null when all are empty.</summary>
    public DateOnly? MaxDate
    {
        get
        {
            DateOnly? max = null;
            foreach (DateOnly? candidate in new[] { AirleakMax, BakoMax, GaussMax })
            {
                if (candidate is DateOnly value && (max is null || value > max.Value))
                    max = value;
            }
            return max;
        }
    }
}

/// <summary>One measured item inside a station cell, with its own pass/fail verdict.</summary>
public sealed record QrSummaryItem(string Label, string Value, QrVerdict Verdict);

public sealed record QrSummaryCell(
    string Station,
    QrVerdict Verdict,
    string Result,
    DateTimeOffset? TestTime,
    IReadOnlyList<QrSummaryItem> Items)
{
    public QrSummaryItem? ItemAt(int index) =>
        index >= 0 && index < Items.Count ? Items[index] : null;
}

public sealed class QrSummaryRow
{
    public string ProductId { get; init; } = string.Empty;

    /// <summary>1 for the newest test of each station, 2 for the one before it, and so on.</summary>
    public int Round { get; init; } = 1;

    public QrSummaryCell? Airleak { get; init; }
    public QrSummaryCell? Bako { get; init; }
    public QrSummaryCell? Gauss { get; init; }

    /// <summary>
    /// Earlier rounds of the same ProductID, newest first. Only the round the grid shows
    /// carries them; the older rounds themselves are always empty.
    /// </summary>
    public List<QrSummaryRow> Older { get; } = [];

    public bool HasOlder => Older.Count > 0;

    /// <summary>
    /// PASS only when all three stations report OK on this round. A station that never tested
    /// this QR counts against it, so an incomplete set reads NG — see <see cref="FinalNote"/>
    /// for which station is responsible.
    /// </summary>
    public bool IsFinalPass =>
        Airleak?.Verdict == QrVerdict.Pass &&
        Bako?.Verdict == QrVerdict.Pass &&
        Gauss?.Verdict == QrVerdict.Pass;

    /// <summary>Why the row is NG: the stations that failed, then the ones with no record.</summary>
    public string FinalNote
    {
        get
        {
            if (IsFinalPass)
                return "AIRLEAK · BAKO · GAUSS all OK";

            var failed = new List<string>();
            var missing = new List<string>();
            foreach ((string station, QrSummaryCell? cell) in new[]
                     {
                         ("AIRLEAK", Airleak), ("BAKO", Bako), ("GAUSS", Gauss),
                     })
            {
                if (cell is null)
                    missing.Add(station);
                else if (cell.Verdict != QrVerdict.Pass)
                    failed.Add(station);
            }

            var parts = new List<string>();
            if (failed.Count > 0)
                parts.Add($"NG: {string.Join(", ", failed)}");
            if (missing.Count > 0)
                parts.Add($"No test record: {string.Join(", ", missing)}");
            return string.Join(" · ", parts);
        }
    }

    public IEnumerable<QrSummaryCell> Cells
    {
        get
        {
            if (Airleak is not null) yield return Airleak;
            if (Bako is not null) yield return Bako;
            if (Gauss is not null) yield return Gauss;
        }
    }

    public int TestedStationCount => Cells.Count();
    public bool HasAnyFail => Cells.Any(c => c.Verdict == QrVerdict.Fail);
    public bool IsCompleteSet => Airleak is not null && Bako is not null && Gauss is not null;

    public DateTimeOffset? LatestTime => Cells.Max(c => c.TestTime);
}

public sealed class QrSummaryResult
{
    public DateOnly From { get; init; }
    public DateOnly To { get; init; }
    public int MaxRows { get; init; }
    public long TotalProducts { get; init; }
    public QrSummaryBounds Bounds { get; init; } = new(null, null, null);
    public DateTime FetchedAt { get; init; }
    public string WarningMessage { get; init; } = string.Empty;
    public List<QrSummaryRow> Rows { get; init; } = [];

    public bool HitMaxRows => Rows.Count >= MaxRows && TotalProducts > Rows.Count;
}
