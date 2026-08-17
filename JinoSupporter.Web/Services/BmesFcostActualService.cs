using System.Data;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;
using Microsoft.Data.SqlClient;

namespace JinoSupporter.Web.Services;

public sealed partial class BmesFcostActualService(
    AppActivityLogger activity,
    NgRateSettingsService settings)
{
    private const int DefaultTimeoutSeconds = 300;
    private const int MaxTimeoutSeconds = 600;
    private const int DefaultMaxRows = 5000;
    private const int HardMaxRows = 50000;
    private const int RawBreakdownCacheSchemaVersion = 1;
    private static readonly TimeSpan RawBreakdownRecentMutableWindow = TimeSpan.FromDays(7);
    private static readonly TimeSpan RawBreakdownRecentCacheTtl = TimeSpan.FromMinutes(15);
    private static readonly JsonSerializerOptions RawBreakdownCacheJsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    private readonly AppActivityLogger _activity = activity;
    private readonly NgRateSettingsService _settings = settings;

    public async Task TestConnectionAsync(
        BmesFcostDbConnection connection,
        CancellationToken cancellationToken = default)
    {
        await using var conn = new SqlConnection(BuildConnectionString(connection));
        await conn.OpenAsync(cancellationToken);
    }

    public async Task<BmesFcostActualResult> FetchAsync(
        BmesFcostActualQuery query,
        CancellationToken cancellationToken = default)
    {
        BmesFcostResolvedPeriod period = ResolvePeriod(query.WorkPeriod);
        int maxRows = NormalizeMaxRows(query.MaxRows);

        await using var conn = new SqlConnection(BuildConnectionString(query.Connection));
        await conn.OpenAsync(cancellationToken);

        await using var cmd = conn.CreateCommand();
        cmd.CommandText = FCostActualSql;
        cmd.CommandTimeout = NormalizeTimeoutSeconds(query.Connection.TimeoutSeconds);
        cmd.Parameters.Add(new SqlParameter("@StartDate", SqlDbType.Date) { Value = period.StartDate });
        cmd.Parameters.Add(new SqlParameter("@EndDateExclusive", SqlDbType.Date) { Value = period.EndDateExclusive });
        cmd.Parameters.Add(new SqlParameter("@Fact", SqlDbType.NVarChar, 20) { Value = NormalizeFilter(query.Fact) });
        cmd.Parameters.Add(new SqlParameter("@Plant", SqlDbType.NVarChar, 20) { Value = NormalizeFilter(query.Plant) });
        cmd.Parameters.Add(new SqlParameter("@Line", SqlDbType.NVarChar, 40) { Value = NormalizeFilter(query.Line) });
        cmd.Parameters.Add(new SqlParameter("@ProductCode", SqlDbType.NVarChar, 80) { Value = NormalizeFilter(query.ProductCode) });
        cmd.Parameters.Add(new SqlParameter("@MaterialCode", SqlDbType.NVarChar, 80) { Value = NormalizeFilter(query.MaterialCode) });
        string searchText = NormalizeFilter(query.SearchText);
        cmd.Parameters.Add(new SqlParameter("@SearchText", SqlDbType.NVarChar, 200) { Value = searchText });
        cmd.Parameters.Add(new SqlParameter("@SearchLike", SqlDbType.NVarChar, 220) { Value = "%" + EscapeLike(searchText) + "%" });
        cmd.Parameters.Add(new SqlParameter("@IncludeZeroFcost", SqlDbType.Bit) { Value = query.IncludeZeroFcost });
        cmd.Parameters.Add(new SqlParameter("@MaxRows", SqlDbType.Int) { Value = maxRows });

        var rows = new List<BmesFcostActualRow>();
        await using var reader = await cmd.ExecuteReaderAsync(CommandBehavior.SequentialAccess, cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            rows.Add(new BmesFcostActualRow
            {
                Fact = ReadString(reader, "Fact"),
                Plant = ReadString(reader, "Plant"),
                Line = ReadString(reader, "Line"),
                ProductCode = ReadString(reader, "ProductCode"),
                ProductName = ReadString(reader, "ProductName"),
                MaterialCode = ReadString(reader, "MaterialCode"),
                MaterialName = ReadString(reader, "MaterialName"),
                WorkPeriod = period.DisplayText,
                StandardPriceVnd = ReadDecimal(reader, "StandardPriceVnd"),
                ActualInputPriceVnd = ReadDecimal(reader, "ActualInputPriceVnd"),
                FCostVnd = ReadDecimal(reader, "FCostVnd"),
                SourceRows = ReadInt64(reader, "SourceRows"),
            });
        }

        _activity.Log(
            "BMES Test2",
            $"FCOST raw {period.DisplayText} {query.Fact.Trim()} {query.Plant.Trim()}: {rows.Count:N0} row(s)");

        return new BmesFcostActualResult
        {
            WorkPeriod = period.DisplayText,
            PeriodKind = period.Kind,
            StartDate = period.StartDate,
            EndDateExclusive = period.EndDateExclusive,
            MaxRows = maxRows,
            Rows = rows,
        };
    }

    public async Task<BmesFcostRawBreakdownResult> FetchRawBreakdownAsync(
        BmesFcostRawBreakdownQuery query,
        CancellationToken cancellationToken = default)
    {
        List<BmesFcostRawBreakdownPeriod> periods = query.Periods
            .Where(p => !string.IsNullOrWhiteSpace(p.Key) && p.StartDate < p.EndDateExclusive)
            .OrderBy(p => p.Ordinal)
            .ToList();
        List<BmesFcostRawBreakdownLineShift> lineShifts = query.LineShifts
            .Where(l => !string.IsNullOrWhiteSpace(l.LineShift))
            .GroupBy(l => l.LineShift.Trim(), StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First() with { LineShift = g.Key })
            .ToList();
        List<LineShiftSqlKey> lineSqlKeys = BuildLineShiftSqlKeys(lineShifts);

        if (periods.Count == 0 || lineShifts.Count == 0 || lineSqlKeys.Count == 0)
        {
            return new BmesFcostRawBreakdownResult
            {
                Periods = periods,
                Rows = [],
            };
        }

        RawBreakdownCacheKey cacheKey = BuildRawBreakdownCacheKey(query, periods, lineShifts);
        RawBreakdownCachePolicy cachePolicy = BuildRawBreakdownCachePolicy(periods);
        BmesFcostRawBreakdownResult? cachedResult = await TryReadRawBreakdownCacheAsync(
            cacheKey,
            cachePolicy,
            cancellationToken);
        if (cachedResult is not null)
            return cachedResult;

        await using var conn = new SqlConnection(BuildConnectionString(query.Connection));
        await conn.OpenAsync(cancellationToken);

        List<BmesFcostExchangeRate> exchangeRates = await FetchExchangeRatesAsync(
            conn,
            periods,
            NormalizeTimeoutSeconds(query.Connection.TimeoutSeconds),
            cancellationToken);

        await using var cmd = conn.CreateCommand();
        cmd.CommandText = BuildRawBreakdownSql(periods.Count, lineSqlKeys.Count);
        cmd.CommandTimeout = NormalizeTimeoutSeconds(query.Connection.TimeoutSeconds);
        cmd.Parameters.Add(new SqlParameter("@Fact", SqlDbType.NVarChar, 20) { Value = NormalizeFilter(query.Fact) });
        cmd.Parameters.Add(new SqlParameter("@Plant", SqlDbType.NVarChar, 20) { Value = NormalizeFilter(query.Plant) });

        for (int i = 0; i < periods.Count; i++)
        {
            var period = periods[i];
            cmd.Parameters.Add(new SqlParameter($"@PeriodOrdinal{i}", SqlDbType.Int) { Value = period.Ordinal });
            cmd.Parameters.Add(new SqlParameter($"@PeriodKey{i}", SqlDbType.NVarChar, 40) { Value = period.Key });
            cmd.Parameters.Add(new SqlParameter($"@PeriodHeader{i}", SqlDbType.NVarChar, 40) { Value = period.Header });
            cmd.Parameters.Add(new SqlParameter($"@PeriodKind{i}", SqlDbType.NVarChar, 20) { Value = period.Kind });
            cmd.Parameters.Add(new SqlParameter($"@StartDate{i}", SqlDbType.Date) { Value = period.StartDate.Date });
            cmd.Parameters.Add(new SqlParameter($"@EndDateExclusive{i}", SqlDbType.Date) { Value = period.EndDateExclusive.Date });
        }

        for (int i = 0; i < lineSqlKeys.Count; i++)
        {
            cmd.Parameters.Add(new SqlParameter($"@LineShift{i}", SqlDbType.NVarChar, 240)
            {
                Value = lineSqlKeys[i].LineShift
            });
            cmd.Parameters.Add(new SqlParameter($"@LineVerid{i}", SqlDbType.NVarChar, 120)
            {
                Value = lineSqlKeys[i].VeridCandidate
            });
        }

        var mapping = lineShifts.ToDictionary(
            l => l.LineShift.Trim(),
            l => l,
            StringComparer.OrdinalIgnoreCase);
        var rows = new Dictionary<string, BmesFcostRawMaterialBreakdownRow>(StringComparer.OrdinalIgnoreCase);

        await using var reader = await cmd.ExecuteReaderAsync(CommandBehavior.SequentialAccess, cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            string periodKey = ReadString(reader, "PeriodKey");
            string lineShift = ReadString(reader, "LineShift");
            string materialCode = ReadString(reader, "MaterialCode");
            string materialName = ReadString(reader, "MaterialName");
            decimal fcost = ReadDecimal(reader, "FCostVnd") ?? 0;
            long sourceRows = ReadInt64(reader, "SourceRows");
            if (fcost <= 0 || !mapping.TryGetValue(lineShift, out var lineMap))
                continue;

            decimal? unitPrice = ReadDecimal(reader, "UnitPrice");
            string priceCurrency = ReadString(reader, "PriceCurrency");
            string priceUnit = ReadString(reader, "PriceUnit");
            decimal? unitPriceVnd = ReadDecimal(reader, "UnitPriceVnd");

            string rowKey = string.Join(
                '\t',
                lineMap.GroupName,
                lineMap.ModelName,
                materialCode,
                materialName);

            if (!rows.TryGetValue(rowKey, out var row))
            {
                row = new BmesFcostRawMaterialBreakdownRow
                {
                    GroupName = lineMap.GroupName,
                    ModelName = lineMap.ModelName,
                    MaterialCode = materialCode,
                    MaterialName = materialName,
                };
                rows[rowKey] = row;
            }

            row.FCostByPeriod[periodKey] = row.FCostByPeriod.GetValueOrDefault(periodKey) + fcost;
            ApplyPeriodPrice(row, periodKey, unitPrice, priceCurrency, priceUnit, unitPriceVnd);
            if (unitPriceVnd is > 0)
            {
                row.EquivalentQtyByPeriod[periodKey] =
                    row.EquivalentQtyByPeriod.GetValueOrDefault(periodKey) + fcost / unitPriceVnd.Value;
            }
            row.TotalFCostVnd += fcost;
            row.SourceRows += sourceRows;
        }
        await reader.DisposeAsync();

        var resultRows = rows.Values
            .Where(r => r.TotalFCostVnd > 0)
            .OrderBy(r => r.GroupName, StringComparer.Ordinal)
            .ThenBy(r => r.ModelName, StringComparer.Ordinal)
            .ThenByDescending(r => r.TotalFCostVnd)
            .ThenBy(r => r.MaterialName, StringComparer.Ordinal)
            .ToList();

        _activity.Log(
            "BMES F-COST Raw",
            $"Raw material breakdown periods={periods.Count:N0}, lineShifts={lineShifts.Count:N0}, rows={resultRows.Count:N0}");

        var result = new BmesFcostRawBreakdownResult
        {
            Periods = periods,
            ExchangeRates = exchangeRates,
            Rows = resultRows,
        };
        await StoreRawBreakdownCacheAsync(cacheKey, cachePolicy, result, cancellationToken);
        return result;
    }

    private async Task<BmesFcostRawBreakdownResult?> TryReadRawBreakdownCacheAsync(
        RawBreakdownCacheKey cacheKey,
        RawBreakdownCachePolicy cachePolicy,
        CancellationToken cancellationToken)
    {
        return await Task.Run(
            () =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return TryReadRawBreakdownCache(cacheKey, cachePolicy);
            },
            cancellationToken);
    }

    private BmesFcostRawBreakdownResult? TryReadRawBreakdownCache(
        RawBreakdownCacheKey cacheKey,
        RawBreakdownCachePolicy cachePolicy)
    {
        try
        {
            string dbPath = GetRawBreakdownCacheDbPath();
            if (string.IsNullOrWhiteSpace(dbPath))
            {
                LogRawBreakdownCache(cacheKey, "miss", "reason=no-db-path");
                return null;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(dbPath) ?? ".");
            using var conn = new SqliteConnection(BuildSqliteConnectionString(dbPath));
            conn.Open();
            EnsureRawBreakdownCacheTable(conn);

            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                """
                SELECT SchemaVersion, PayloadJson, ExpiresAtUtc
                FROM FCostRawBreakdownCache
                WHERE KeyHash = @key;
                """;
            cmd.Parameters.AddWithValue("@key", cacheKey.Hash);

            using var reader = cmd.ExecuteReader();
            if (!reader.Read())
            {
                LogRawBreakdownCache(cacheKey, "miss", CacheModeDetail(cachePolicy) + " reason=empty");
                return null;
            }

            int schemaVersion = reader.IsDBNull(0) ? 0 : Convert.ToInt32(reader.GetValue(0), CultureInfo.InvariantCulture);
            if (schemaVersion != RawBreakdownCacheSchemaVersion)
            {
                LogRawBreakdownCache(cacheKey, "miss", CacheModeDetail(cachePolicy) + $" reason=schema-{schemaVersion}");
                return null;
            }

            string payloadJson = reader.IsDBNull(1) ? string.Empty : reader.GetString(1);
            string expiresAtText = reader.IsDBNull(2) ? string.Empty : reader.GetString(2);
            if (!string.IsNullOrWhiteSpace(expiresAtText))
            {
                if (!TryParseCacheInstant(expiresAtText, out DateTimeOffset expiresAtUtc))
                {
                    TryDeleteRawBreakdownCache(conn, cacheKey.Hash);
                    LogRawBreakdownCache(cacheKey, "miss", CacheModeDetail(cachePolicy) + " reason=corrupt-expiry");
                    return null;
                }

                if (expiresAtUtc <= DateTimeOffset.UtcNow)
                {
                    TryDeleteRawBreakdownCache(conn, cacheKey.Hash);
                    LogRawBreakdownCache(cacheKey, "miss", CacheModeDetail(cachePolicy) + " reason=expired");
                    return null;
                }
            }

            BmesFcostRawBreakdownResult? result = JsonSerializer.Deserialize<BmesFcostRawBreakdownResult>(
                payloadJson,
                RawBreakdownCacheJsonOptions);
            if (result is null)
                throw new JsonException("Raw breakdown cache payload is empty.");
            ValidateRawBreakdownCachePayload(result);

            TryTouchRawBreakdownCache(conn, cacheKey.Hash);
            LogRawBreakdownCache(
                cacheKey,
                "hit",
                CacheModeDetail(cachePolicy) + $" rows={result.Rows.Count:N0}");
            return result;
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            LogRawBreakdownCache(cacheKey, "miss", "reason=unavailable " + CompactException(ex));
            return null;
        }
    }

    private async Task StoreRawBreakdownCacheAsync(
        RawBreakdownCacheKey cacheKey,
        RawBreakdownCachePolicy cachePolicy,
        BmesFcostRawBreakdownResult result,
        CancellationToken cancellationToken)
    {
        await Task.Run(
            () =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                StoreRawBreakdownCache(cacheKey, cachePolicy, result);
            },
            cancellationToken);
    }

    private void StoreRawBreakdownCache(
        RawBreakdownCacheKey cacheKey,
        RawBreakdownCachePolicy cachePolicy,
        BmesFcostRawBreakdownResult result)
    {
        try
        {
            string dbPath = GetRawBreakdownCacheDbPath();
            if (string.IsNullOrWhiteSpace(dbPath))
            {
                LogRawBreakdownCache(cacheKey, "store-skip", "reason=no-db-path");
                return;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(dbPath) ?? ".");
            using var conn = new SqliteConnection(BuildSqliteConnectionString(dbPath));
            conn.Open();
            EnsureRawBreakdownCacheTable(conn);

            DateTimeOffset nowUtc = DateTimeOffset.UtcNow;
            DateTimeOffset? expiresAtUtc = cachePolicy.UsesTtl
                ? nowUtc.Add(RawBreakdownRecentCacheTtl)
                : null;
            string payloadJson = JsonSerializer.Serialize(result, RawBreakdownCacheJsonOptions);

            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                """
                INSERT INTO FCostRawBreakdownCache
                    (KeyHash, SchemaVersion, CacheKeyJson, PayloadJson, CreatedAtUtc, ExpiresAtUtc,
                     LastHitAtUtc, HitCount, PeriodCount, LineShiftCount, RowCount)
                VALUES
                    (@key, @schema, @keyJson, @payload, @created, @expires,
                     NULL, 0, @periods, @lineShifts, @rows)
                ON CONFLICT(KeyHash) DO UPDATE SET
                    SchemaVersion = excluded.SchemaVersion,
                    CacheKeyJson = excluded.CacheKeyJson,
                    PayloadJson = excluded.PayloadJson,
                    CreatedAtUtc = excluded.CreatedAtUtc,
                    ExpiresAtUtc = excluded.ExpiresAtUtc,
                    LastHitAtUtc = NULL,
                    HitCount = 0,
                    PeriodCount = excluded.PeriodCount,
                    LineShiftCount = excluded.LineShiftCount,
                    RowCount = excluded.RowCount;
                """;
            cmd.Parameters.AddWithValue("@key", cacheKey.Hash);
            cmd.Parameters.AddWithValue("@schema", RawBreakdownCacheSchemaVersion);
            cmd.Parameters.AddWithValue("@keyJson", cacheKey.Json);
            cmd.Parameters.AddWithValue("@payload", payloadJson);
            cmd.Parameters.AddWithValue("@created", nowUtc.ToString("O", CultureInfo.InvariantCulture));
            cmd.Parameters.AddWithValue(
                "@expires",
                expiresAtUtc is null
                    ? DBNull.Value
                    : expiresAtUtc.Value.ToString("O", CultureInfo.InvariantCulture));
            cmd.Parameters.AddWithValue("@periods", cacheKey.PeriodCount);
            cmd.Parameters.AddWithValue("@lineShifts", cacheKey.LineShiftCount);
            cmd.Parameters.AddWithValue("@rows", result.Rows.Count);
            cmd.ExecuteNonQuery();

            string ttlText = expiresAtUtc is null
                ? "ttl=historical"
                : $"ttl={RawBreakdownRecentCacheTtl.TotalMinutes:N0}m expires={expiresAtUtc.Value:O}";
            LogRawBreakdownCache(cacheKey, "store", $"{ttlText} rows={result.Rows.Count:N0}");
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            LogRawBreakdownCache(cacheKey, "store-fail", CompactException(ex));
        }
    }

    private RawBreakdownCacheKey BuildRawBreakdownCacheKey(
        BmesFcostRawBreakdownQuery query,
        IReadOnlyList<BmesFcostRawBreakdownPeriod> periods,
        IReadOnlyList<BmesFcostRawBreakdownLineShift> lineShifts)
    {
        string fact = NormalizeFilter(query.Fact);
        string plant = NormalizeFilter(query.Plant);
        string database = string.IsNullOrWhiteSpace(query.Connection.Database)
            ? "BMES_LIV"
            : query.Connection.Database.Trim();

        var payload = new RawBreakdownCacheKeyPayload(
            RawBreakdownCacheSchemaVersion,
            database,
            fact,
            plant,
            periods
                .Select(p => new RawBreakdownCachePeriod(
                    p.Ordinal,
                    NormalizeFilter(p.Key),
                    NormalizeFilter(p.Header),
                    NormalizeFilter(p.Kind),
                    p.StartDate.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                    p.EndDateExclusive.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)))
                .ToList(),
            lineShifts
                .Select(l => new RawBreakdownCacheLineShift(
                    NormalizeFilter(l.GroupName),
                    NormalizeFilter(l.ModelName),
                    NormalizeFilter(l.LineShift)))
                .OrderBy(l => l.LineShift, StringComparer.OrdinalIgnoreCase)
                .ThenBy(l => l.GroupName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(l => l.ModelName, StringComparer.OrdinalIgnoreCase)
                .ToList());
        string json = JsonSerializer.Serialize(payload, RawBreakdownCacheJsonOptions);
        return new RawBreakdownCacheKey(
            Sha256Hex(json),
            json,
            fact,
            plant,
            payload.Periods.Count,
            payload.LineShifts.Count);
    }

    private static RawBreakdownCachePolicy BuildRawBreakdownCachePolicy(
        IReadOnlyList<BmesFcostRawBreakdownPeriod> periods)
    {
        DateTime mutableWindowStart = DateTime.Today.Subtract(RawBreakdownRecentMutableWindow);
        bool usesTtl = periods.Any(p => p.EndDateExclusive.Date.AddDays(-1) >= mutableWindowStart);
        return new RawBreakdownCachePolicy(usesTtl);
    }

    private string GetRawBreakdownCacheDbPath()
        => Path.Combine(_settings.FCostDbSaveDirectory, "fcost_raw.db");

    private static string BuildSqliteConnectionString(string dbPath)
    {
        var builder = new SqliteConnectionStringBuilder
        {
            DataSource = dbPath,
            Mode = SqliteOpenMode.ReadWriteCreate,
        };
        return builder.ToString();
    }

    private static void EnsureRawBreakdownCacheTable(SqliteConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            """
            PRAGMA journal_mode=WAL;
            PRAGMA busy_timeout=5000;

            CREATE TABLE IF NOT EXISTS FCostRawBreakdownCache (
                KeyHash TEXT PRIMARY KEY,
                SchemaVersion INTEGER NOT NULL,
                CacheKeyJson TEXT NOT NULL,
                PayloadJson TEXT NOT NULL,
                CreatedAtUtc TEXT NOT NULL,
                ExpiresAtUtc TEXT NULL,
                LastHitAtUtc TEXT NULL,
                HitCount INTEGER NOT NULL DEFAULT 0,
                PeriodCount INTEGER NOT NULL DEFAULT 0,
                LineShiftCount INTEGER NOT NULL DEFAULT 0,
                RowCount INTEGER NOT NULL DEFAULT 0
            );
            CREATE INDEX IF NOT EXISTS IX_FCostRawBreakdownCache_ExpiresAtUtc
                ON FCostRawBreakdownCache(ExpiresAtUtc);
            """;
        cmd.ExecuteNonQuery();
    }

    private static void TryTouchRawBreakdownCache(SqliteConnection conn, string keyHash)
    {
        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                """
                UPDATE FCostRawBreakdownCache
                SET LastHitAtUtc = @hit,
                    HitCount = HitCount + 1
                WHERE KeyHash = @key;
                """;
            cmd.Parameters.AddWithValue("@hit", DateTimeOffset.UtcNow.ToString("O", CultureInfo.InvariantCulture));
            cmd.Parameters.AddWithValue("@key", keyHash);
            cmd.ExecuteNonQuery();
        }
        catch
        {
            // Cache hit accounting must not turn a valid payload into a miss.
        }
    }

    private static void TryDeleteRawBreakdownCache(SqliteConnection conn, string keyHash)
    {
        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM FCostRawBreakdownCache WHERE KeyHash = @key;";
            cmd.Parameters.AddWithValue("@key", keyHash);
            cmd.ExecuteNonQuery();
        }
        catch
        {
            // Best-effort cleanup only; the next read will still validate before use.
        }
    }

    private static void ValidateRawBreakdownCachePayload(BmesFcostRawBreakdownResult result)
    {
        if (result.Periods is null ||
            result.ExchangeRates is null ||
            result.Rows is null)
        {
            throw new JsonException("Raw breakdown cache payload has null top-level collections.");
        }

        foreach (var row in result.Rows)
        {
            if (row.FCostByPeriod is null ||
                row.EquivalentQtyByPeriod is null ||
                row.PriceByPeriod is null)
            {
                throw new JsonException("Raw breakdown cache payload has null row dictionaries.");
            }
        }
    }

    private static bool TryParseCacheInstant(string value, out DateTimeOffset instant)
        => DateTimeOffset.TryParse(
            value,
            CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
            out instant);

    private void LogRawBreakdownCache(RawBreakdownCacheKey cacheKey, string action, string detail)
    {
        _activity.Log(
            "BMES F-COST Raw",
            $"cache {action} key={ShortHash(cacheKey.Hash)} fact={cacheKey.Fact} plant={cacheKey.Plant} periods={cacheKey.PeriodCount:N0} lineShifts={cacheKey.LineShiftCount:N0} {detail}");
    }

    private static string CacheModeDetail(RawBreakdownCachePolicy cachePolicy)
        => cachePolicy.UsesTtl ? "mode=recent" : "mode=historical";

    private static string Sha256Hex(string value)
        => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();

    private static string ShortHash(string hash)
        => hash.Length <= 12 ? hash : hash[..12];

    private static string CompactException(Exception ex)
    {
        string message = Regex.Replace(ex.Message, @"\s+", " ").Trim();
        if (message.Length > 120)
            message = message[..120] + "...";

        return string.IsNullOrWhiteSpace(message)
            ? ex.GetType().Name
            : $"{ex.GetType().Name}: {message}";
    }

    private sealed record RawBreakdownCacheKey(
        string Hash,
        string Json,
        string Fact,
        string Plant,
        int PeriodCount,
        int LineShiftCount);

    private sealed record RawBreakdownCachePolicy(bool UsesTtl);

    private sealed record RawBreakdownCacheKeyPayload(
        int SchemaVersion,
        string Database,
        string Fact,
        string Plant,
        List<RawBreakdownCachePeriod> Periods,
        List<RawBreakdownCacheLineShift> LineShifts);

    private sealed record RawBreakdownCachePeriod(
        int Ordinal,
        string Key,
        string Header,
        string Kind,
        string StartDate,
        string EndDateExclusive);

    private sealed record RawBreakdownCacheLineShift(
        string GroupName,
        string ModelName,
        string LineShift);

    public async Task<List<BmesBomMaterialCandidate>> FetchBomMaterialsAsync(
        BmesBomMaterialQuery query,
        CancellationToken cancellationToken = default)
    {
        string modelName = NormalizeFilter(query.ModelName);
        if (string.IsNullOrWhiteSpace(modelName))
            return [];

        await using var conn = new SqlConnection(BuildConnectionString(query.Connection));
        await conn.OpenAsync(cancellationToken);

        return await FetchBomMaterialsWithSqlAsync(
            conn,
            ActualMaterialCandidatesSql,
            "FCOST",
            query,
            cancellationToken);
    }

    /// <summary>
    /// Full MAST/BOMC BOM explosion for one model. Unlike <see cref="FetchBomMaterialsAsync"/>
    /// the rows keep their parent/level context, so the same material may appear more than
    /// once under different parents. <c>UsageQty</c> is the raw <c>BOMC.MENGE</c> of that one
    /// BOM line — it is not rolled up across levels, because BOMC carries no documented base
    /// quantity to normalise against.
    /// </summary>
    public async Task<List<BmesBomMaterialCandidate>> FetchBomTreeAsync(
        BmesBomMaterialQuery query,
        CancellationToken cancellationToken = default)
    {
        string modelName = NormalizeFilter(query.ModelName);
        if (string.IsNullOrWhiteSpace(modelName))
            return [];

        await using var conn = new SqlConnection(BuildConnectionString(query.Connection));
        await conn.OpenAsync(cancellationToken);

        var rows = await ReadBomRowsAsync(conn, BomMaterialsSql, "BOM", query, cancellationToken);

        // MAST can hold several BOM alternatives per material, so the same component can
        // arrive twice on one path; keep the first occurrence of each path.
        rows = rows
            .GroupBy(r => r.BomPath, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .ToList();

        return DecorateBomTree(rows);
    }

    /// <summary>
    /// Searches the actual BMES product master instead of assuming that a report model-group
    /// label is also a MATE product name. Only products that have an active BOM for the selected
    /// plant/date are returned. Individual words and number fragments are ranked independently,
    /// so a label such as "ASSY 338 RA1" can find a product such as "ASSY REAR-TAPE-338".
    /// </summary>
    public async Task<List<BmesBomModelCandidate>> SearchBomModelsAsync(
        BmesBomModelSearchQuery query,
        CancellationToken cancellationToken = default)
    {
        string searchText = NormalizeFilter(query.SearchText);
        int maxRows = Math.Clamp(query.MaxRows <= 0 ? 10 : query.MaxRows, 1, 30);
        List<string> tokens = BomModelTokenRegex()
            .Matches(searchText)
            .Select(match => match.Value.Trim())
            .Where(token => token.Length >= 2)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(token => token.Any(char.IsDigit))
            .ThenByDescending(token => token.Length)
            .Take(6)
            .ToList();

        if (searchText.Length < 2 || tokens.Count == 0)
            return [];

        const string productCodeSql = "RTRIM(CAST(m.MATNR AS nvarchar(80)))";
        const string productNameSql = "RTRIM(CAST(m.MAKTX AS nvarchar(200)))";

        var matchConditions = new List<string>();
        var scoreParts = new List<string>
        {
            $"CASE WHEN {productNameSql} = @SearchText THEN 1200 WHEN {productCodeSql} = @SearchText THEN 1100 ELSE 0 END",
            $"CASE WHEN {productNameSql} LIKE @SearchLike OR {productCodeSql} LIKE @SearchLike THEN 500 ELSE 0 END",
        };

        for (int i = 0; i < tokens.Count; i++)
        {
            string condition =
                $"({productNameSql} LIKE @TokenLike{i} OR {productCodeSql} LIKE @TokenLike{i})";
            matchConditions.Add(condition);

            // Digits usually identify the product family more accurately than generic words
            // such as ASSY, so make those matches dominate the BMES-side candidate ordering.
            int weight = tokens[i].Any(char.IsDigit) ? 180 : 80;
            scoreParts.Add($"CASE WHEN {condition} THEN {weight + Math.Min(tokens[i].Length, 20)} ELSE 0 END");
        }

        string matchSql = string.Join(Environment.NewLine + "            OR ", matchConditions);
        string scoreSql = string.Join(Environment.NewLine + "                + ", scoreParts);
        string sql =
            $"""
            WITH CandidateRows AS (
                SELECT TOP (@ScanRows)
                    {productCodeSql} AS ProductCode,
                    {productNameSql} AS ProductName,
                    {scoreSql} AS MatchScore
                FROM dbo.MATE AS m WITH (NOLOCK)
                WHERE ISNULL(m.MATNR, N'') <> N''
                  AND ISNULL(m.MAKTX, N'') <> N''
                  AND (
                        {matchSql}
                  )
                  AND EXISTS (
                      SELECT 1
                      FROM dbo.MAST AS mast WITH (NOLOCK)
                      JOIN dbo.BOMC AS bom WITH (NOLOCK)
                        ON bom.STLNR = mast.STLNR
                       AND bom.STLAL = mast.STLAL
                       AND bom.STLAN = mast.STLAN
                      WHERE RTRIM(CAST(mast.MATNR AS nvarchar(80))) = {productCodeSql}
                        AND (@Plant = N'' OR RTRIM(CAST(mast.WERKS AS nvarchar(20))) = @Plant)
                        AND ISNULL(bom.CMATE, N'') <> N''
                        AND ISNULL(bom.USEYN, N'Y') = N'Y'
                        AND (bom.SDATE IS NULL OR bom.SDATE <= @WorkDate)
                        AND (bom.EDATE IS NULL OR bom.EDATE >= @WorkDate)
                  )
                ORDER BY
                    MatchScore DESC,
                    LEN({productNameSql}),
                    {productNameSql},
                    {productCodeSql}
            )
            SELECT
                ProductCode,
                ProductName,
                MAX(MatchScore) AS MatchScore
            FROM CandidateRows
            GROUP BY ProductCode, ProductName
            ORDER BY
                MatchScore DESC,
                LEN(ProductName),
                ProductName,
                ProductCode;
            """;

        await using var conn = new SqlConnection(BuildConnectionString(query.Connection));
        await conn.OpenAsync(cancellationToken);

        await using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        cmd.CommandTimeout = NormalizeTimeoutSeconds(query.Connection.TimeoutSeconds);
        cmd.Parameters.Add(new SqlParameter("@SearchText", SqlDbType.NVarChar, 200) { Value = searchText });
        cmd.Parameters.Add(new SqlParameter("@SearchLike", SqlDbType.NVarChar, 220) { Value = "%" + EscapeLike(searchText) + "%" });
        cmd.Parameters.Add(new SqlParameter("@Plant", SqlDbType.NVarChar, 20) { Value = NormalizeFilter(query.Plant) });
        cmd.Parameters.Add(new SqlParameter("@WorkDate", SqlDbType.Date) { Value = query.WorkDate.Date });
        cmd.Parameters.Add(new SqlParameter("@ScanRows", SqlDbType.Int) { Value = Math.Clamp(maxRows * 30, 100, 600) });
        for (int i = 0; i < tokens.Count; i++)
        {
            cmd.Parameters.Add(new SqlParameter($"@TokenLike{i}", SqlDbType.NVarChar, 220)
            {
                Value = "%" + EscapeLike(tokens[i]) + "%",
            });
        }

        var candidates = new List<BmesBomModelCandidate>();
        await using var reader = await cmd.ExecuteReaderAsync(CommandBehavior.SequentialAccess, cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            candidates.Add(new BmesBomModelCandidate
            {
                ProductCode = ReadString(reader, "ProductCode"),
                ProductName = ReadString(reader, "ProductName"),
                MatchScore = ReadInt32(reader, "MatchScore"),
            });
        }

        return candidates
            .OrderByDescending(candidate => candidate.MatchScore)
            .ThenBy(candidate => candidate.ProductName.Length)
            .ThenBy(candidate => candidate.ProductName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(candidate => candidate.ProductCode, StringComparer.OrdinalIgnoreCase)
            .Take(maxRows)
            .ToList();
    }

    /// <summary>
    /// Downloads only the BMES finished-product code/name catalog for local searching. Test 5
    /// calls this explicitly when the user refreshes model names; normal typing never reaches
    /// SQL Server. Only products in the requested code families with a valid BOM for the selected
    /// plant/date are returned.
    /// </summary>
    public async Task<List<BmesBomModelCandidate>> FetchBomModelCatalogAsync(
        BmesBomModelCatalogQuery query,
        CancellationToken cancellationToken = default)
    {
        int maxRows = Math.Clamp(query.MaxRows <= 0 ? 20000 : query.MaxRows, 1, 50000);
        List<string> codePrefixes = query.ProductCodePrefixes
            .Select(NormalizeFilter)
            .Where(prefix => prefix.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        string productCodeSql = "RTRIM(CAST(m.MATNR AS nvarchar(80)))";
        string codePrefixFilter = codePrefixes.Count == 0
            ? "1 = 0"
            : string.Join(
                Environment.NewLine + "                    OR ",
                codePrefixes.Select((_, index) => $"{productCodeSql} LIKE @CodePrefixLike{index}"));

        string sql =
            $"""
            SELECT TOP (@MaxRows)
                RTRIM(CAST(m.MATNR AS nvarchar(80))) AS ProductCode,
                RTRIM(CAST(m.MAKTX AS nvarchar(200))) AS ProductName,
                CAST(0 AS int) AS MatchScore
            FROM dbo.MATE AS m WITH (NOLOCK)
            WHERE ISNULL(m.MATNR, N'') <> N''
              AND ISNULL(m.MAKTX, N'') <> N''
              AND (
                    {codePrefixFilter}
              )
              AND EXISTS (
                  SELECT 1
                  FROM dbo.MAST AS mast WITH (NOLOCK)
                  JOIN dbo.BOMC AS bom WITH (NOLOCK)
                    ON bom.STLNR = mast.STLNR
                   AND bom.STLAL = mast.STLAL
                   AND bom.STLAN = mast.STLAN
                  WHERE RTRIM(CAST(mast.MATNR AS nvarchar(80))) =
                        RTRIM(CAST(m.MATNR AS nvarchar(80)))
                    AND (@Plant = N'' OR RTRIM(CAST(mast.WERKS AS nvarchar(20))) = @Plant)
                    AND ISNULL(bom.CMATE, N'') <> N''
                    AND ISNULL(bom.USEYN, N'Y') = N'Y'
                    AND (bom.SDATE IS NULL OR bom.SDATE <= @WorkDate)
                    AND (bom.EDATE IS NULL OR bom.EDATE >= @WorkDate)
              )
            GROUP BY
                RTRIM(CAST(m.MATNR AS nvarchar(80))),
                RTRIM(CAST(m.MAKTX AS nvarchar(200)))
            ORDER BY
                ProductName,
                ProductCode;
            """;

        await using var conn = new SqlConnection(BuildConnectionString(query.Connection));
        await conn.OpenAsync(cancellationToken);

        await using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        cmd.CommandTimeout = NormalizeTimeoutSeconds(query.Connection.TimeoutSeconds);
        cmd.Parameters.Add(new SqlParameter("@Plant", SqlDbType.NVarChar, 20)
        {
            Value = NormalizeFilter(query.Plant),
        });
        cmd.Parameters.Add(new SqlParameter("@WorkDate", SqlDbType.Date)
        {
            Value = query.WorkDate.Date,
        });
        for (int i = 0; i < codePrefixes.Count; i++)
        {
            cmd.Parameters.Add(new SqlParameter($"@CodePrefixLike{i}", SqlDbType.NVarChar, 60)
            {
                Value = EscapeLike(codePrefixes[i]) + "%",
            });
        }
        cmd.Parameters.Add(new SqlParameter("@MaxRows", SqlDbType.Int) { Value = maxRows });

        var models = new List<BmesBomModelCandidate>();
        await using var reader = await cmd.ExecuteReaderAsync(CommandBehavior.SequentialAccess, cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            models.Add(new BmesBomModelCandidate
            {
                ProductCode = ReadString(reader, "ProductCode"),
                ProductName = ReadString(reader, "ProductName"),
            });
        }

        _activity.Log(
            "BMES Test5",
            $"Model-name catalog {NormalizeFilter(query.Plant)} {query.WorkDate:yyyy-MM-dd}: {models.Count:N0} row(s)");

        return models;
    }

    /// <summary>
    /// Fills in the parts the recursive CTE cannot produce: SQL Server forbids subqueries and
    /// outer joins inside a recursive member, so the tree comes back with code-only paths and
    /// no child flag.
    /// </summary>
    private static List<BmesBomMaterialCandidate> DecorateBomTree(List<BmesBomMaterialCandidate> rows)
    {
        var nameByCode = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in rows)
        {
            if (!string.IsNullOrWhiteSpace(row.MaterialCode) && !string.IsNullOrWhiteSpace(row.MaterialName))
                nameByCode[row.MaterialCode] = row.MaterialName;
            if (!string.IsNullOrWhiteSpace(row.ParentMaterialCode) && !string.IsNullOrWhiteSpace(row.ParentMaterialName))
                nameByCode[row.ParentMaterialCode] = row.ParentMaterialName;
            if (!string.IsNullOrWhiteSpace(row.ProductCode) && !string.IsNullOrWhiteSpace(row.ProductName))
                nameByCode[row.ProductCode] = row.ProductName;
        }

        var parentPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in rows)
        {
            int cut = row.BomPath.LastIndexOf('>');
            if (cut > 0)
                parentPaths.Add(row.BomPath[..cut]);
        }

        return rows
            .Select(row => new BmesBomMaterialCandidate
            {
                ProductCode = row.ProductCode,
                ProductName = row.ProductName,
                ParentMaterialCode = row.ParentMaterialCode,
                ParentMaterialName = row.ParentMaterialName,
                MaterialCode = row.MaterialCode,
                MaterialName = row.MaterialName,
                UsageQty = row.UsageQty,
                UsageUnit = row.UsageUnit,
                BomLevel = row.BomLevel,
                BomPath = row.BomPath,
                BomPathText = string.Join(
                    " > ",
                    row.BomPath
                        .Split('>', StringSplitOptions.RemoveEmptyEntries)
                        .Select(code => nameByCode.TryGetValue(code, out string? name) ? name : code)),
                HasChildren = parentPaths.Contains(row.BomPath),
                SourceRows = row.SourceRows,
                Source = row.Source,
            })
            .ToList();
    }

    private static async Task<List<BmesBomMaterialCandidate>> FetchBomMaterialsWithSqlAsync(
        SqlConnection conn,
        string sql,
        string source,
        BmesBomMaterialQuery query,
        CancellationToken cancellationToken)
    {
        var rows = await ReadBomRowsAsync(conn, sql, source, query, cancellationToken);

        return rows
            .Where(r => !string.IsNullOrWhiteSpace(r.MaterialCode) || !string.IsNullOrWhiteSpace(r.MaterialName))
            .GroupBy(r => NormalizeFilter(r.MaterialCode) + "\t" + NormalizeFilter(r.MaterialName), StringComparer.OrdinalIgnoreCase)
            .Select(g => g.OrderByDescending(r => r.SourceRows).First())
            .OrderBy(r => r.MaterialName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(r => r.MaterialCode, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static async Task<List<BmesBomMaterialCandidate>> ReadBomRowsAsync(
        SqlConnection conn,
        string sql,
        string source,
        BmesBomMaterialQuery query,
        CancellationToken cancellationToken)
    {
        string modelName = NormalizeFilter(query.ModelName);
        string plant = NormalizeFilter(query.Plant);
        int maxRows = Math.Clamp(query.MaxRows <= 0 ? 300 : query.MaxRows, 1, 2000);

        await using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        cmd.CommandTimeout = NormalizeTimeoutSeconds(query.Connection.TimeoutSeconds);
        cmd.Parameters.Add(new SqlParameter("@ModelSearch", SqlDbType.NVarChar, 200) { Value = modelName });
        cmd.Parameters.Add(new SqlParameter("@ModelLike", SqlDbType.NVarChar, 220) { Value = "%" + EscapeLike(modelName) + "%" });
        cmd.Parameters.Add(new SqlParameter("@Plant", SqlDbType.NVarChar, 20) { Value = plant });
        cmd.Parameters.Add(new SqlParameter("@MaxRows", SqlDbType.Int) { Value = maxRows });
        cmd.Parameters.Add(new SqlParameter("@MaxDepth", SqlDbType.Int) { Value = Math.Clamp(query.MaxDepth <= 0 ? 6 : query.MaxDepth, 1, 12) });
        cmd.Parameters.Add(new SqlParameter("@WorkDate", SqlDbType.Date) { Value = query.WorkDate.Date });

        var rows = new List<BmesBomMaterialCandidate>();
        await using var reader = await cmd.ExecuteReaderAsync(CommandBehavior.SequentialAccess, cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            rows.Add(new BmesBomMaterialCandidate
            {
                ProductCode = ReadString(reader, "ProductCode"),
                ProductName = ReadString(reader, "ProductName"),
                ParentMaterialCode = ReadString(reader, "ParentMaterialCode"),
                ParentMaterialName = ReadString(reader, "ParentMaterialName"),
                MaterialCode = ReadString(reader, "MaterialCode"),
                MaterialName = ReadString(reader, "MaterialName"),
                UsageQty = ReadDecimal(reader, "UsageQty"),
                UsageUnit = ReadString(reader, "UsageUnit"),
                BomLevel = ReadInt32(reader, "BomLevel"),
                BomPath = ReadString(reader, "BomPath"),
                BomPathText = ReadString(reader, "BomPathText"),
                HasChildren = ReadBool(reader, "HasChildren"),
                SourceRows = ReadInt64(reader, "SourceRows"),
                Source = source,
            });
        }

        return rows;
    }

    private static void ApplyPeriodPrice(
        BmesFcostRawMaterialBreakdownRow row,
        string periodKey,
        decimal? unitPrice,
        string priceCurrency,
        string priceUnit,
        decimal? unitPriceVnd)
    {
        if (unitPrice is null || string.IsNullOrWhiteSpace(priceCurrency))
            return;

        priceCurrency = priceCurrency.Trim().ToUpperInvariant();
        priceUnit = priceUnit.Trim();
        var next = new BmesFcostRawMaterialPeriodPrice
        {
            UnitPrice = unitPrice,
            Currency = priceCurrency,
            PriceUnit = priceUnit,
            UnitPriceVnd = unitPriceVnd,
        };

        if (!row.PriceByPeriod.TryGetValue(periodKey, out var current))
        {
            row.PriceByPeriod[periodKey] = next;
            return;
        }

        if (current.UnitPrice != next.UnitPrice ||
            current.UnitPriceVnd != next.UnitPriceVnd ||
            !string.Equals(current.Currency, next.Currency, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(current.PriceUnit, next.PriceUnit, StringComparison.OrdinalIgnoreCase))
        {
            current.IsMixed = true;
        }
    }

    private static async Task<List<BmesFcostExchangeRate>> FetchExchangeRatesAsync(
        SqlConnection conn,
        IReadOnlyList<BmesFcostRawBreakdownPeriod> periods,
        int timeoutSeconds,
        CancellationToken cancellationToken)
    {
        var standardDates = periods
            .Select(p => new DateTime(p.StartDate.Year, p.StartDate.Month, 1))
            .Distinct()
            .OrderBy(d => d)
            .ToList();
        if (standardDates.Count == 0)
            return [];

        string rateRows = string.Join(
            "," + Environment.NewLine,
            Enumerable.Range(0, standardDates.Count).Select(i => $"(@RateDate{i})"));

        await using var cmd = conn.CreateCommand();
        cmd.CommandTimeout = timeoutSeconds;
        cmd.CommandText =
            $"""
            WITH RateDates AS (
                SELECT v.StandardDate
                FROM (VALUES
                    {rateRows}
                ) AS v(StandardDate)
            )
            SELECT
                rd.StandardDate AS StandardDate,
                CAST(MAX(CASE WHEN RTRIM(t.FCURR) = N'USD'
                    THEN CAST(t.UKURS AS decimal(38, 10)) * CAST(t.TFACT AS decimal(38, 10)) / NULLIF(CAST(t.FFACT AS decimal(38, 10)), 0)
                    ELSE NULL END) AS decimal(38, 10)) AS KrwPerUsd,
                CAST(MAX(CASE WHEN RTRIM(t.FCURR) = N'VND'
                    THEN CAST(t.UKURS AS decimal(38, 10)) * CAST(t.TFACT AS decimal(38, 10)) / NULLIF(CAST(t.FFACT AS decimal(38, 10)), 0)
                    ELSE NULL END) AS decimal(38, 10)) AS KrwPerVnd
            FROM RateDates AS rd
            LEFT JOIN dbo.TCURR AS t WITH (NOLOCK)
                ON CONVERT(date, t.GDATU) = rd.StandardDate
               AND RTRIM(t.KURST) = N'BWCU'
               AND RTRIM(t.TCURR) = N'KRW'
               AND RTRIM(t.FCURR) IN (N'USD', N'VND')
            GROUP BY rd.StandardDate
            ORDER BY rd.StandardDate;
            """;

        for (int i = 0; i < standardDates.Count; i++)
        {
            cmd.Parameters.Add(new SqlParameter($"@RateDate{i}", SqlDbType.Date)
            {
                Value = standardDates[i].Date
            });
        }

        var ratesByDate = new Dictionary<DateTime, (decimal? KrwPerUsd, decimal? KrwPerVnd)>();
        await using (var reader = await cmd.ExecuteReaderAsync(CommandBehavior.SequentialAccess, cancellationToken))
        {
            while (await reader.ReadAsync(cancellationToken))
            {
                DateTime standardDate = reader.GetDateTime(reader.GetOrdinal("StandardDate")).Date;
                ratesByDate[standardDate] = (
                    ReadDecimal(reader, "KrwPerUsd"),
                    ReadDecimal(reader, "KrwPerVnd"));
            }
        }

        return periods
            .Select(period =>
            {
                DateTime standardDate = new(period.StartDate.Year, period.StartDate.Month, 1);
                ratesByDate.TryGetValue(standardDate, out var rate);
                return new BmesFcostExchangeRate(
                    period.Key,
                    standardDate,
                    rate.KrwPerUsd,
                    rate.KrwPerVnd);
            })
            .ToList();
    }

    private sealed record LineShiftSqlKey(string LineShift, string VeridCandidate);

    private static List<LineShiftSqlKey> BuildLineShiftSqlKeys(IEnumerable<BmesFcostRawBreakdownLineShift> lineShifts)
    {
        var result = new List<LineShiftSqlKey>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var lineShift in lineShifts)
        {
            string value = lineShift.LineShift.Trim();
            if (string.IsNullOrWhiteSpace(value))
                continue;

            foreach (string candidate in EnumerateLineShiftVeridCandidates(value))
            {
                string key = value + "\t" + candidate;
                if (seen.Add(key))
                    result.Add(new LineShiftSqlKey(value, candidate));
            }
        }

        return result;
    }

    private static IEnumerable<string> EnumerateLineShiftVeridCandidates(string lineShift)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (int i = lineShift.IndexOf('_'); i >= 0 && i < lineShift.Length - 1; i = lineShift.IndexOf('_', i + 1))
        {
            string candidate = lineShift[(i + 1)..].Trim();
            if (!string.IsNullOrWhiteSpace(candidate) && seen.Add(candidate))
                yield return candidate;
        }

        string full = lineShift.Trim();
        if (!string.IsNullOrWhiteSpace(full) && seen.Add(full))
            yield return full;
    }

    private static string BuildConnectionString(BmesFcostDbConnection connection)
    {
        if (string.IsNullOrWhiteSpace(connection.Server))
            throw new InvalidOperationException("Server is required.");
        if (string.IsNullOrWhiteSpace(connection.UserId))
            throw new InvalidOperationException("User ID is required.");

        var builder = new SqlConnectionStringBuilder
        {
            DataSource = $"{connection.Server.Trim()},{Math.Clamp(connection.Port, 1, 65535)}",
            InitialCatalog = string.IsNullOrWhiteSpace(connection.Database)
                ? "BMES_LIV"
                : connection.Database.Trim(),
            UserID = connection.UserId.Trim(),
            Password = connection.Password,
            ConnectTimeout = NormalizeTimeoutSeconds(connection.TimeoutSeconds),
            Encrypt = connection.Encrypt,
            TrustServerCertificate = connection.TrustServerCertificate,
            MultipleActiveResultSets = false,
            ApplicationName = "JinoSupporter.Web BMES FCost Test2",
            ApplicationIntent = ApplicationIntent.ReadOnly,
        };

        return builder.ConnectionString;
    }

    private static BmesFcostResolvedPeriod ResolvePeriod(string workPeriod)
    {
        string text = workPeriod.Trim();
        if (string.IsNullOrWhiteSpace(text))
            throw new InvalidOperationException("Work Date is required.");

        if (DateTime.TryParseExact(
                text,
                "yyyy-MM-dd",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out DateTime dashDate))
        {
            return new BmesFcostResolvedPeriod(
                dashDate.Date.ToString("yyyyMMdd", CultureInfo.InvariantCulture),
                "Day",
                dashDate.Date,
                dashDate.Date.AddDays(1));
        }

        if (DateTime.TryParseExact(
                text,
                "yyyyMMdd",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out DateTime compactDate))
        {
            return new BmesFcostResolvedPeriod(
                compactDate.Date.ToString("yyyyMMdd", CultureInfo.InvariantCulture),
                "Day",
                compactDate.Date,
                compactDate.Date.AddDays(1));
        }

        Match weekMatch = WeekPeriodRegex().Match(text);
        if (weekMatch.Success)
        {
            int year = int.Parse(weekMatch.Groups["year"].Value, CultureInfo.InvariantCulture);
            int week = int.Parse(weekMatch.Groups["week"].Value, CultureInfo.InvariantCulture);
            int maxWeek = ISOWeek.GetWeeksInYear(year);
            if (week < 1 || week > maxWeek)
                throw new InvalidOperationException($"Week must be between 1 and {maxWeek} for {year}.");

            DateTime start = ISOWeek.ToDateTime(year, week, DayOfWeek.Monday);
            return new BmesFcostResolvedPeriod(
                $"{year:0000}-{week:00}",
                "Week",
                start.Date,
                start.Date.AddDays(7));
        }

        Match monthMatch = MonthPeriodRegex().Match(text);
        if (monthMatch.Success)
        {
            int year = int.Parse(monthMatch.Groups["year"].Value, CultureInfo.InvariantCulture);
            int month = int.Parse(monthMatch.Groups["month"].Value, CultureInfo.InvariantCulture);
            if (month is < 1 or > 12)
                throw new InvalidOperationException("Month must be between 01 and 12.");

            var start = new DateTime(year, month, 1);
            return new BmesFcostResolvedPeriod(
                $"{year:0000}{month:00}",
                "Month",
                start,
                start.AddMonths(1));
        }

        throw new InvalidOperationException("Work Date must be yyyyMMdd, yyyy-MM-dd, yyyy-WW, or yyyyMM.");
    }

    private static string NormalizeFilter(string value) => value.Trim();

    private static int NormalizeTimeoutSeconds(int timeoutSeconds) =>
        timeoutSeconds <= 0
            ? DefaultTimeoutSeconds
            : Math.Clamp(timeoutSeconds, 1, MaxTimeoutSeconds);

    private static int NormalizeMaxRows(int maxRows) =>
        maxRows <= 0
            ? DefaultMaxRows
            : Math.Clamp(maxRows, 1, HardMaxRows);

    private static string EscapeLike(string value) =>
        value
            .Replace("[", "[[]", StringComparison.Ordinal)
            .Replace("%", "[%]", StringComparison.Ordinal)
            .Replace("_", "[_]", StringComparison.Ordinal);

    private static string ReadString(SqlDataReader reader, string columnName)
    {
        int ordinal = reader.GetOrdinal(columnName);
        if (reader.IsDBNull(ordinal))
            return string.Empty;

        return Convert.ToString(reader.GetValue(ordinal), CultureInfo.InvariantCulture)?.Trim()
            ?? string.Empty;
    }

    private static decimal? ReadDecimal(SqlDataReader reader, string columnName)
    {
        int ordinal = reader.GetOrdinal(columnName);
        if (reader.IsDBNull(ordinal))
            return null;

        object value = reader.GetValue(ordinal);
        try
        {
            return Convert.ToDecimal(value, CultureInfo.InvariantCulture);
        }
        catch
        {
            return null;
        }
    }

    private static long ReadInt64(SqlDataReader reader, string columnName)
    {
        int ordinal = reader.GetOrdinal(columnName);
        if (reader.IsDBNull(ordinal))
            return 0;

        object value = reader.GetValue(ordinal);
        try
        {
            return Convert.ToInt64(value, CultureInfo.InvariantCulture);
        }
        catch
        {
            return 0;
        }
    }

    private static int ReadInt32(SqlDataReader reader, string columnName)
    {
        int ordinal = reader.GetOrdinal(columnName);
        if (reader.IsDBNull(ordinal))
            return 0;

        object value = reader.GetValue(ordinal);
        try
        {
            return Convert.ToInt32(value, CultureInfo.InvariantCulture);
        }
        catch
        {
            return 0;
        }
    }

    private static bool ReadBool(SqlDataReader reader, string columnName)
        => ReadInt32(reader, columnName) != 0;

    [GeneratedRegex(@"^(?<year>\d{4})-(?<week>\d{1,2})$")]
    private static partial Regex WeekPeriodRegex();

    [GeneratedRegex(@"^(?<year>\d{4})(?<month>\d{2})$")]
    private static partial Regex MonthPeriodRegex();

    [GeneratedRegex(@"[\p{L}\p{N}]+")]
    private static partial Regex BomModelTokenRegex();

    private const string FCostActualSql =
        """
        SELECT TOP (@MaxRows)
            e.FACCO AS Fact,
            e.WERKS AS Plant,
            e.VERID AS Line,
            e.MATNR AS ProductCode,
            ISNULL(productName.ProductName, N'') AS ProductName,
            e.ASSEM AS MaterialCode,
            ISNULL(materialName.MaterialName, N'') AS MaterialName,
            CAST(ROUND(SUM(CAST(ISNULL(e.INAMT, 0) AS decimal(38, 6))) * 100, 0) AS decimal(38, 0)) AS StandardPriceVnd,
            CAST(ROUND(SUM(CAST(ISNULL(e.FCAMT, 0) AS decimal(38, 6))) * 100, 0) AS decimal(38, 0)) AS ActualInputPriceVnd,
            CASE
                WHEN ROUND(
                    SUM(CAST(ISNULL(e.FCAMT, 0) AS decimal(38, 6)) - CAST(ISNULL(e.INAMT, 0) AS decimal(38, 6))) * 100,
                    0
                ) > 0 THEN CAST(ROUND(
                    SUM(CAST(ISNULL(e.FCAMT, 0) AS decimal(38, 6)) - CAST(ISNULL(e.INAMT, 0) AS decimal(38, 6))) * 100,
                    0
                ) AS decimal(38, 0))
                ELSE NULL
            END AS FCostVnd,
            COUNT_BIG(*) AS SourceRows
        FROM dbo.E008 AS e WITH (NOLOCK)
        OUTER APPLY (
            SELECT TOP (1) CAST(m.MAKTX AS nvarchar(200)) AS ProductName
            FROM dbo.MATE AS m WITH (NOLOCK)
            WHERE m.MATNR = e.MATNR
            ORDER BY m.MAKTX
        ) AS productName
        OUTER APPLY (
            SELECT TOP (1) CAST(m.MAKTX AS nvarchar(200)) AS MaterialName
            FROM dbo.MATE AS m WITH (NOLOCK)
            WHERE m.MATNR = e.ASSEM
            ORDER BY m.MAKTX
        ) AS materialName
        WHERE e.WDATE >= @StartDate
          AND e.WDATE < @EndDateExclusive
          AND (@Fact = N'' OR e.FACCO = @Fact)
          AND (@Plant = N'' OR e.WERKS = @Plant)
          AND (@Line = N'' OR e.VERID = @Line)
          AND (@ProductCode = N'' OR e.MATNR = @ProductCode)
          AND (@MaterialCode = N'' OR e.ASSEM = @MaterialCode)
          AND ISNULL(e.DAUAT, N'') <> N'ZP01'
          AND NOT (ISNULL(e.DAUAT, N'') = N'ZP09' AND ISNULL(e.FCAMT, 0) < 0)
          AND (
                @SearchText = N''
                OR e.MATNR LIKE @SearchLike
                OR e.ASSEM LIKE @SearchLike
                OR EXISTS (
                    SELECT 1
                    FROM dbo.MATE AS searchProduct WITH (NOLOCK)
                    WHERE searchProduct.MATNR = e.MATNR
                      AND CAST(searchProduct.MAKTX AS nvarchar(200)) LIKE @SearchLike
                )
                OR EXISTS (
                    SELECT 1
                    FROM dbo.MATE AS searchMaterial WITH (NOLOCK)
                    WHERE searchMaterial.MATNR = e.ASSEM
                      AND CAST(searchMaterial.MAKTX AS nvarchar(200)) LIKE @SearchLike
                )
          )
        GROUP BY
            e.FACCO,
            e.WERKS,
            e.VERID,
            e.MATNR,
            productName.ProductName,
            e.ASSEM,
            materialName.MaterialName
        HAVING
            @IncludeZeroFcost = 1
            OR ROUND(
                SUM(CAST(ISNULL(e.FCAMT, 0) AS decimal(38, 6)) - CAST(ISNULL(e.INAMT, 0) AS decimal(38, 6))) * 100,
                0
            ) > 0
        ORDER BY
            e.MATNR,
            e.VERID,
            e.ASSEM;
        """;

    private const string BomMaterialsSql =
        """
        WITH ProductRows AS (
            SELECT TOP (20)
                RTRIM(CAST(m.MATNR AS nvarchar(80))) AS ProductCode,
                CAST(m.MAKTX AS nvarchar(200)) AS ProductName
            FROM dbo.MATE AS m WITH (NOLOCK)
            WHERE RTRIM(CAST(m.MATNR AS nvarchar(80))) = @ModelSearch
               OR CAST(m.MAKTX AS nvarchar(200)) = @ModelSearch
               OR CAST(m.MAKTX AS nvarchar(200)) LIKE @ModelLike
            ORDER BY
                CASE
                    WHEN CAST(m.MAKTX AS nvarchar(200)) = @ModelSearch THEN 0
                    WHEN RTRIM(CAST(m.MATNR AS nvarchar(80))) = @ModelSearch THEN 1
                    ELSE 2
                END,
                m.MATNR
        ),
        BomTree AS (
            SELECT
                p.ProductCode,
                p.ProductName,
                p.ProductCode AS ParentMaterialCode,
                RTRIM(CAST(b.CMATE AS nvarchar(80))) AS MaterialCode,
                CAST(b.MENGE AS decimal(28, 6)) AS UsageQty,
                RTRIM(CAST(b.MEINS AS nvarchar(40))) AS UsageUnit,
                1 AS BomLevel,
                CAST(p.ProductCode + N'>' + RTRIM(CAST(b.CMATE AS nvarchar(80))) AS nvarchar(4000)) AS BomPath
            FROM ProductRows AS p
            JOIN dbo.MAST AS mast WITH (NOLOCK)
                ON RTRIM(CAST(mast.MATNR AS nvarchar(80))) = p.ProductCode
               AND (@Plant = N'' OR RTRIM(CAST(mast.WERKS AS nvarchar(20))) = @Plant)
            JOIN dbo.BOMC AS b WITH (NOLOCK)
                ON b.STLNR = mast.STLNR
               AND b.STLAL = mast.STLAL
               AND b.STLAN = mast.STLAN
            WHERE ISNULL(b.CMATE, N'') <> N''
              AND ISNULL(b.USEYN, N'Y') = N'Y'
              AND (b.SDATE IS NULL OR b.SDATE <= @WorkDate)
              AND (b.EDATE IS NULL OR b.EDATE >= @WorkDate)

            UNION ALL

            SELECT
                bt.ProductCode,
                bt.ProductName,
                bt.MaterialCode AS ParentMaterialCode,
                RTRIM(CAST(c.CMATE AS nvarchar(80))) AS MaterialCode,
                CAST(c.MENGE AS decimal(28, 6)) AS UsageQty,
                RTRIM(CAST(c.MEINS AS nvarchar(40))) AS UsageUnit,
                bt.BomLevel + 1 AS BomLevel,
                CAST(bt.BomPath + N'>' + RTRIM(CAST(c.CMATE AS nvarchar(80))) AS nvarchar(4000)) AS BomPath
            FROM BomTree AS bt
            JOIN dbo.MAST AS mast WITH (NOLOCK)
                ON RTRIM(CAST(mast.MATNR AS nvarchar(80))) = bt.MaterialCode
               AND (@Plant = N'' OR RTRIM(CAST(mast.WERKS AS nvarchar(20))) = @Plant)
            JOIN dbo.BOMC AS c WITH (NOLOCK)
                ON c.STLNR = mast.STLNR
               AND c.STLAL = mast.STLAL
               AND c.STLAN = mast.STLAN
            WHERE ISNULL(c.CMATE, N'') <> N''
              AND ISNULL(c.USEYN, N'Y') = N'Y'
              AND (c.SDATE IS NULL OR c.SDATE <= @WorkDate)
              AND (c.EDATE IS NULL OR c.EDATE >= @WorkDate)
              AND bt.BomLevel < @MaxDepth
              AND CHARINDEX(N'>' + RTRIM(CAST(c.CMATE AS nvarchar(80))) + N'>', N'>' + bt.BomPath + N'>') = 0
        )
        SELECT TOP (@MaxRows)
            bt.ProductCode,
            bt.ProductName,
            bt.ParentMaterialCode,
            ISNULL(parentName.MaterialName, N'') AS ParentMaterialName,
            bt.MaterialCode,
            ISNULL(childName.MaterialName, N'') AS MaterialName,
            bt.UsageQty,
            bt.UsageUnit,
            bt.BomLevel,
            bt.BomPath,
            bt.BomPath AS BomPathText,
            CAST(0 AS int) AS HasChildren,
            CAST(1 AS bigint) AS SourceRows
        FROM BomTree AS bt
        OUTER APPLY (
            SELECT TOP (1) CAST(m.MAKTX AS nvarchar(200)) AS MaterialName
            FROM dbo.MATE AS m WITH (NOLOCK)
            WHERE RTRIM(CAST(m.MATNR AS nvarchar(80))) = bt.MaterialCode
            ORDER BY m.MAKTX
        ) AS childName
        OUTER APPLY (
            SELECT TOP (1) CAST(m.MAKTX AS nvarchar(200)) AS MaterialName
            FROM dbo.MATE AS m WITH (NOLOCK)
            WHERE RTRIM(CAST(m.MATNR AS nvarchar(80))) = bt.ParentMaterialCode
            ORDER BY m.MAKTX
        ) AS parentName
        ORDER BY
            bt.ProductName,
            bt.BomPath
        OPTION (MAXRECURSION 12);
        """;

    private const string ActualMaterialCandidatesSql =
        """
        WITH ProductRows AS (
            SELECT TOP (20)
                RTRIM(CAST(m.MATNR AS nvarchar(80))) AS ProductCode,
                CAST(m.MAKTX AS nvarchar(200)) AS ProductName
            FROM dbo.MATE AS m WITH (NOLOCK)
            WHERE RTRIM(CAST(m.MATNR AS nvarchar(80))) = @ModelSearch
               OR CAST(m.MAKTX AS nvarchar(200)) = @ModelSearch
               OR CAST(m.MAKTX AS nvarchar(200)) LIKE @ModelLike
            ORDER BY
                CASE
                    WHEN CAST(m.MAKTX AS nvarchar(200)) = @ModelSearch THEN 0
                    WHEN RTRIM(CAST(m.MATNR AS nvarchar(80))) = @ModelSearch THEN 1
                    ELSE 2
                END,
                m.MATNR
        )
        SELECT TOP (@MaxRows)
            p.ProductCode,
            p.ProductName,
            p.ProductCode AS ParentMaterialCode,
            p.ProductName AS ParentMaterialName,
            RTRIM(CAST(e.ASSEM AS nvarchar(80))) AS MaterialCode,
            ISNULL(materialName.MaterialName, N'') AS MaterialName,
            CAST(NULL AS decimal(38, 10)) AS UsageQty,
            CAST(N'' AS nvarchar(40)) AS UsageUnit,
            CAST(1 AS int) AS BomLevel,
            CAST(p.ProductCode + N'>' + RTRIM(CAST(e.ASSEM AS nvarchar(80))) AS nvarchar(max)) AS BomPath,
            CAST(p.ProductName + N' > ' + ISNULL(materialName.MaterialName, N'') AS nvarchar(max)) AS BomPathText,
            CAST(0 AS int) AS HasChildren,
            COUNT_BIG(*) AS SourceRows
        FROM ProductRows AS p
        JOIN dbo.E008 AS e WITH (NOLOCK)
            ON RTRIM(CAST(e.MATNR AS nvarchar(80))) = p.ProductCode
        OUTER APPLY (
            SELECT TOP (1) CAST(m.MAKTX AS nvarchar(200)) AS MaterialName
            FROM dbo.MATE AS m WITH (NOLOCK)
            WHERE RTRIM(CAST(m.MATNR AS nvarchar(80))) = RTRIM(CAST(e.ASSEM AS nvarchar(80)))
            ORDER BY m.MAKTX
        ) AS materialName
        WHERE ISNULL(e.ASSEM, N'') <> N''
          AND (@Plant = N'' OR RTRIM(CAST(e.WERKS AS nvarchar(20))) = @Plant)
          AND ISNULL(e.DAUAT, N'') <> N'ZP01'
          AND NOT (ISNULL(e.DAUAT, N'') = N'ZP09' AND ISNULL(e.FCAMT, 0) < 0)
        GROUP BY
            p.ProductCode,
            p.ProductName,
            e.ASSEM,
            materialName.MaterialName
        ORDER BY
            COUNT_BIG(*) DESC,
            materialName.MaterialName,
            e.ASSEM;
        """;

    private static string BuildRawBreakdownSql(int periodCount, int lineShiftCount)
    {
        string periodRows = string.Join(
            "," + Environment.NewLine,
            Enumerable.Range(0, periodCount).Select(i =>
                $"(@PeriodOrdinal{i}, @PeriodKey{i}, @PeriodHeader{i}, @PeriodKind{i}, @StartDate{i}, @EndDateExclusive{i})"));
        string lineShiftRows = string.Join(
            "," + Environment.NewLine,
            Enumerable.Range(0, lineShiftCount).Select(i => $"(@LineShift{i}, @LineVerid{i})"));

        return
            $"""
            WITH Periods AS (
                SELECT
                    v.PeriodOrdinal,
                    v.PeriodKey,
                    v.PeriodHeader,
                    v.PeriodKind,
                    v.StartDate,
                    v.EndDateExclusive
                FROM (VALUES
                    {periodRows}
                ) AS v(PeriodOrdinal, PeriodKey, PeriodHeader, PeriodKind, StartDate, EndDateExclusive)
            ),
            LineKeys AS (
                SELECT DISTINCT
                    v.LineShift,
                    v.VeridCandidate
                FROM (VALUES
                    {lineShiftRows}
                ) AS v(LineShift, VeridCandidate)
                WHERE v.LineShift <> N''
                  AND v.VeridCandidate <> N''
            ),
            RawAgg AS (
                SELECT
                    p.PeriodOrdinal AS PeriodOrdinal,
                    p.PeriodKey AS PeriodKey,
                    p.PeriodHeader AS PeriodHeader,
                    p.PeriodKind AS PeriodKind,
                    p.StartDate AS PeriodStartDate,
                    lk.LineShift AS LineShift,
                    e.WERKS AS Plant,
                    e.VERID AS Line,
                    e.MATNR AS ProductCode,
                    e.ASSEM AS MaterialCode,
                    CAST(ROUND(
                        SUM(CAST(ISNULL(e.FCAMT, 0) AS decimal(38, 6)) - CAST(ISNULL(e.INAMT, 0) AS decimal(38, 6))) * 100,
                        0
                    ) AS decimal(38, 0)) AS FCostVnd,
                    COUNT_BIG(*) AS SourceRows
                FROM Periods AS p
                JOIN dbo.E008 AS e WITH (NOLOCK)
                    ON e.WDATE >= p.StartDate
                   AND e.WDATE < p.EndDateExclusive
                JOIN LineKeys AS lk
                    ON e.VERID = lk.VeridCandidate
                WHERE e.FACCO = @Fact
                  AND e.WERKS = @Plant
                  AND ISNULL(e.DAUAT, N'') <> N'ZP01'
                  AND NOT (ISNULL(e.DAUAT, N'') = N'ZP09' AND ISNULL(e.FCAMT, 0) < 0)
                GROUP BY
                    p.PeriodOrdinal,
                    p.PeriodKey,
                    p.PeriodHeader,
                    p.PeriodKind,
                    p.StartDate,
                    lk.LineShift,
                    e.WERKS,
                    e.VERID,
                    e.MATNR,
                    e.ASSEM
                HAVING ROUND(
                    SUM(CAST(ISNULL(e.FCAMT, 0) AS decimal(38, 6)) - CAST(ISNULL(e.INAMT, 0) AS decimal(38, 6))) * 100,
                    0
                ) > 0
            )
            SELECT
                ra.PeriodOrdinal AS PeriodOrdinal,
                ra.PeriodKey AS PeriodKey,
                ra.PeriodHeader AS PeriodHeader,
                ra.PeriodKind AS PeriodKind,
                ra.LineShift AS LineShift,
                ra.Line AS Line,
                ra.ProductCode AS ProductCode,
                ISNULL(productName.ProductName, N'') AS ProductName,
                ra.MaterialCode AS MaterialCode,
                ISNULL(materialName.MaterialName, N'') AS MaterialName,
                ra.FCostVnd AS FCostVnd,
                ra.SourceRows AS SourceRows,
                resolvedPrice.UnitPrice AS UnitPrice,
                resolvedPrice.PriceCurrency AS PriceCurrency,
                resolvedPrice.PriceUnit AS PriceUnit,
                CAST(CASE
                    WHEN resolvedPrice.UnitPrice IS NULL THEN NULL
                    WHEN resolvedPrice.PriceCurrency = N'VND' THEN resolvedPrice.UnitPrice
                    WHEN resolvedPrice.PriceCurrency = N'KRW' AND vndRate.KrwPerVnd > 0
                        THEN resolvedPrice.UnitPrice / vndRate.KrwPerVnd
                    WHEN priceRate.KrwPerCurrency > 0 AND vndRate.KrwPerVnd > 0
                        THEN resolvedPrice.UnitPrice * priceRate.KrwPerCurrency / vndRate.KrwPerVnd
                    ELSE NULL
                END AS decimal(38, 10)) AS UnitPriceVnd
            FROM RawAgg AS ra
            OUTER APPLY (
                SELECT CAST(MIN(m.MAKTX) AS nvarchar(200)) AS ProductName
                FROM dbo.MATE AS m WITH (NOLOCK)
                WHERE m.MATNR = ra.ProductCode
            ) AS productName
            OUTER APPLY (
                SELECT CAST(MIN(m.MAKTX) AS nvarchar(200)) AS MaterialName
                FROM dbo.MATE AS m WITH (NOLOCK)
                WHERE m.MATNR = ra.MaterialCode
            ) AS materialName
            OUTER APPLY (
                SELECT TOP (1)
                    CAST(CAST(i.KBETR AS decimal(38, 10)) / NULLIF(CAST(i.KPEIN AS decimal(38, 10)), 0) AS decimal(38, 10)) AS UnitPrice,
                    UPPER(RTRIM(CAST(i.KONWA AS nvarchar(20)))) AS PriceCurrency,
                    RTRIM(CAST(i.KMEIN AS nvarchar(40))) AS PriceUnit
                FROM dbo.INFR AS i WITH (NOLOCK)
                WHERE i.MATNR = ra.MaterialCode
                  AND i.DATAB <= CONVERT(nvarchar(8), ra.PeriodStartDate, 112)
                  AND i.DATBI >= CONVERT(nvarchar(8), ra.PeriodStartDate, 112)
                ORDER BY
                    CASE
                        WHEN i.EKORG = ra.Plant THEN 0
                        WHEN LEFT(i.EKORG, 2) = LEFT(ra.Plant, 2) THEN 1
                        ELSE 2
                    END,
                    i.DATAB DESC,
                    i.EKORG,
                    i.LIFNR
            ) AS price
            OUTER APPLY (
                SELECT TOP (1)
                    CAST(CAST(s.DMBTR AS decimal(38, 10)) / NULLIF(CAST(s.MENGE AS decimal(38, 10)), 0) AS decimal(38, 10)) AS UnitPrice,
                    UPPER(RTRIM(CAST(s.WAERS AS nvarchar(20)))) AS PriceCurrency,
                    RTRIM(CAST(s.MEINS AS nvarchar(40))) AS PriceUnit
                FROM dbo.STBI AS s WITH (NOLOCK)
                WHERE s.MATNR = ra.MaterialCode
                  AND s.MENGE > 0
                  AND s.DMBTR > 0
                  AND s.ZDATE >= DATEFROMPARTS(YEAR(ra.PeriodStartDate), MONTH(ra.PeriodStartDate), 1)
                  AND s.ZDATE < DATEADD(MONTH, 1, DATEFROMPARTS(YEAR(ra.PeriodStartDate), MONTH(ra.PeriodStartDate), 1))
                ORDER BY
                    CASE WHEN RTRIM(s.ZBUKRS) = @Fact THEN 0 ELSE 1 END,
                    CASE
                        WHEN RTRIM(s.GUBUN) = N'A' THEN 0
                        WHEN RTRIM(s.GUBUN) = N'B' THEN 1
                        WHEN RTRIM(s.GUBUN) = N'G' THEN 2
                        ELSE 3
                    END,
                    s.ZDATE DESC
            ) AS stockPrice
            OUTER APPLY (
                SELECT
                    CASE
                        WHEN price.UnitPrice IS NOT NULL AND NULLIF(price.PriceCurrency, N'') IS NOT NULL THEN price.UnitPrice
                        ELSE stockPrice.UnitPrice
                    END AS UnitPrice,
                    CASE
                        WHEN price.UnitPrice IS NOT NULL AND NULLIF(price.PriceCurrency, N'') IS NOT NULL THEN price.PriceCurrency
                        ELSE stockPrice.PriceCurrency
                    END AS PriceCurrency,
                    CASE
                        WHEN price.UnitPrice IS NOT NULL AND NULLIF(price.PriceCurrency, N'') IS NOT NULL THEN price.PriceUnit
                        ELSE stockPrice.PriceUnit
                    END AS PriceUnit
            ) AS resolvedPrice
            OUTER APPLY (
                SELECT TOP (1)
                    CAST(t.UKURS AS decimal(38, 10)) * CAST(t.TFACT AS decimal(38, 10)) / NULLIF(CAST(t.FFACT AS decimal(38, 10)), 0) AS KrwPerCurrency
                FROM dbo.TCURR AS t WITH (NOLOCK)
                WHERE RTRIM(t.KURST) = N'BWCU'
                  AND RTRIM(t.TCURR) = N'KRW'
                  AND RTRIM(t.FCURR) = resolvedPrice.PriceCurrency
                  AND CONVERT(date, t.GDATU) = DATEFROMPARTS(YEAR(ra.PeriodStartDate), MONTH(ra.PeriodStartDate), 1)
                ORDER BY t.GDATU DESC
            ) AS priceRate
            OUTER APPLY (
                SELECT TOP (1)
                    CAST(t.UKURS AS decimal(38, 10)) * CAST(t.TFACT AS decimal(38, 10)) / NULLIF(CAST(t.FFACT AS decimal(38, 10)), 0) AS KrwPerVnd
                FROM dbo.TCURR AS t WITH (NOLOCK)
                WHERE RTRIM(t.KURST) = N'BWCU'
                  AND RTRIM(t.TCURR) = N'KRW'
                  AND RTRIM(t.FCURR) = N'VND'
                  AND CONVERT(date, t.GDATU) = DATEFROMPARTS(YEAR(ra.PeriodStartDate), MONTH(ra.PeriodStartDate), 1)
                ORDER BY t.GDATU DESC
            ) AS vndRate
            WHERE CONCAT(ISNULL(productName.ProductName, N''), N'_', ra.Line) = ra.LineShift
            ORDER BY
                ra.PeriodOrdinal,
                productName.ProductName,
                ra.Line,
                ra.MaterialCode;
            """;
    }

    private sealed record BmesFcostResolvedPeriod(
        string DisplayText,
        string Kind,
        DateTime StartDate,
        DateTime EndDateExclusive);
}

public sealed class BmesFcostActualQuery
{
    public BmesFcostDbConnection Connection { get; init; } = new();
    public string WorkPeriod { get; init; } = string.Empty;
    public string Fact { get; init; } = string.Empty;
    public string Plant { get; init; } = string.Empty;
    public string Line { get; init; } = string.Empty;
    public string ProductCode { get; init; } = string.Empty;
    public string MaterialCode { get; init; } = string.Empty;
    public string SearchText { get; init; } = string.Empty;
    public int MaxRows { get; init; } = 5000;
    public bool IncludeZeroFcost { get; init; } = true;
}

public sealed class BmesBomMaterialQuery
{
    public BmesFcostDbConnection Connection { get; init; } = new();
    public string ModelName { get; init; } = string.Empty;
    public string Plant { get; init; } = "3200";
    public int MaxRows { get; init; } = 300;
    public int MaxDepth { get; init; } = 6;

    /// <summary>Date used against the <c>BOMC.SDATE</c>/<c>BOMC.EDATE</c> validity window.</summary>
    public DateTime WorkDate { get; init; } = DateTime.Today;
}

public sealed class BmesBomModelSearchQuery
{
    public BmesFcostDbConnection Connection { get; init; } = new();
    public string SearchText { get; init; } = string.Empty;
    public string Plant { get; init; } = "3200";
    public DateTime WorkDate { get; init; } = DateTime.Today;
    public int MaxRows { get; init; } = 10;
}

public sealed class BmesBomModelCatalogQuery
{
    public BmesFcostDbConnection Connection { get; init; } = new();
    public string Plant { get; init; } = "3200";
    public DateTime WorkDate { get; init; } = DateTime.Today;
    public IReadOnlyList<string> ProductCodePrefixes { get; init; } = ["P-S-", "P-M-", "P-N-", "P-H-"];
    public int MaxRows { get; init; } = 20000;
}

public sealed class BmesFcostDbConnection
{
    public string Server { get; init; } = string.Empty;
    public int Port { get; init; } = 1430;
    public string Database { get; init; } = "BMES_LIV";
    public string UserId { get; init; } = string.Empty;
    public string Password { get; init; } = string.Empty;
    public int TimeoutSeconds { get; init; } = 300;
    public bool Encrypt { get; init; } = true;
    public bool TrustServerCertificate { get; init; } = true;
}

public sealed class BmesBomModelCandidate
{
    public string ProductCode { get; init; } = string.Empty;
    public string ProductName { get; init; } = string.Empty;
    public int MatchScore { get; init; }
}

public sealed class BmesFcostActualResult
{
    public string WorkPeriod { get; init; } = string.Empty;
    public string PeriodKind { get; init; } = string.Empty;
    public DateTime StartDate { get; init; }
    public DateTime EndDateExclusive { get; init; }
    public int MaxRows { get; init; }
    public string SourceTable { get; init; } = "dbo.E008";
    public string NameSource { get; init; } = "dbo.MATE.MAKTX";
    public List<BmesFcostActualRow> Rows { get; init; } = new();
    public bool HitMaxRows => Rows.Count >= MaxRows;
}

public sealed class BmesBomMaterialCandidate
{
    public string ProductCode { get; init; } = string.Empty;
    public string ProductName { get; init; } = string.Empty;
    public string ParentMaterialCode { get; init; } = string.Empty;
    public string ParentMaterialName { get; init; } = string.Empty;
    public string MaterialCode { get; init; } = string.Empty;
    public string MaterialName { get; init; } = string.Empty;
    public decimal? UsageQty { get; init; }
    public string UsageUnit { get; init; } = string.Empty;
    public int BomLevel { get; init; }
    public string BomPath { get; init; } = string.Empty;
    public string BomPathText { get; init; } = string.Empty;
    public bool HasChildren { get; init; }
    public long SourceRows { get; init; }
    public string Source { get; init; } = string.Empty;
}

public sealed class BmesFcostActualRow
{
    public string Fact { get; init; } = string.Empty;
    public string Plant { get; init; } = string.Empty;
    public string Line { get; init; } = string.Empty;
    public string ProductCode { get; init; } = string.Empty;
    public string ProductName { get; init; } = string.Empty;
    public string MaterialCode { get; init; } = string.Empty;
    public string MaterialName { get; init; } = string.Empty;
    public string WorkPeriod { get; init; } = string.Empty;
    public decimal? StandardPriceVnd { get; init; }
    public decimal? ActualInputPriceVnd { get; init; }
    public decimal? FCostVnd { get; init; }
    public long SourceRows { get; init; }
}

public sealed class BmesFcostRawBreakdownQuery
{
    public BmesFcostDbConnection Connection { get; init; } = new();
    public string Fact { get; init; } = "GN";
    public string Plant { get; init; } = "3200";
    public IReadOnlyList<BmesFcostRawBreakdownPeriod> Periods { get; init; } = [];
    public IReadOnlyList<BmesFcostRawBreakdownLineShift> LineShifts { get; init; } = [];
}

public sealed record BmesFcostRawBreakdownPeriod(
    int Ordinal,
    string Key,
    string Header,
    string Kind,
    DateTime StartDate,
    DateTime EndDateExclusive);

public sealed record BmesFcostRawBreakdownLineShift(
    string GroupName,
    string ModelName,
    string LineShift);

public sealed class BmesFcostRawBreakdownResult
{
    public string SourceTable { get; init; } = "dbo.E008";
    public string NameSource { get; init; } = "dbo.MATE.MAKTX";
    public string WarningMessage { get; init; } = string.Empty;
    public IReadOnlyList<BmesFcostRawBreakdownPeriod> Periods { get; init; } = [];
    public IReadOnlyList<BmesFcostExchangeRate> ExchangeRates { get; init; } = [];
    public List<BmesFcostRawMaterialBreakdownRow> Rows { get; init; } = [];
}

public sealed record BmesFcostExchangeRate(
    string PeriodKey,
    DateTime StandardDate,
    decimal? KrwPerUsd,
    decimal? KrwPerVnd);

public sealed class BmesFcostRawMaterialBreakdownRow
{
    public string GroupName { get; init; } = string.Empty;
    public string ModelName { get; init; } = string.Empty;
    public string MaterialCode { get; init; } = string.Empty;
    public string MaterialName { get; init; } = string.Empty;
    public Dictionary<string, decimal> FCostByPeriod { get; init; } = new(StringComparer.Ordinal);
    public Dictionary<string, decimal> EquivalentQtyByPeriod { get; init; } = new(StringComparer.Ordinal);
    public Dictionary<string, BmesFcostRawMaterialPeriodPrice> PriceByPeriod { get; init; } = new(StringComparer.Ordinal);
    public decimal TotalFCostVnd { get; set; }
    public long SourceRows { get; set; }
}

public sealed class BmesFcostRawMaterialPeriodPrice
{
    public decimal? UnitPrice { get; set; }
    public string Currency { get; set; } = string.Empty;
    public string PriceUnit { get; set; } = string.Empty;
    public decimal? UnitPriceVnd { get; set; }
    public bool IsMixed { get; set; }
}
