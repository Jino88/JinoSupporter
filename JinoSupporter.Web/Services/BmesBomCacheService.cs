using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Local, model-level cache for the read-only BMES MAST/BOMC explosion used by Test 5.
/// The network connection and its credentials are deliberately not persisted.
/// </summary>
public sealed class BmesBomCacheService(NgRateSettingsService settings)
{
    private readonly object _writeGate = new();

    public string CacheDbPath =>
        Path.Combine(settings.SettingsDbDirectory, "bmes_bom_cache.db");

    public BmesBomCacheEntry? TryLoad(BmesBomCacheQuery query)
    {
        EnsureDatabase();
        string cacheKey = BuildCacheKey(query);

        using var conn = OpenConnection(readOnly: true);
        using var header = conn.CreateCommand();
        header.CommandText =
            """
            SELECT ModelName, Plant, WorkDate, MaxDepth, MaxRows, FetchedAt, RowCount
            FROM BmesBomCacheHeaders
            WHERE CacheKey = @key;
            """;
        header.Parameters.AddWithValue("@key", cacheKey);

        string modelName;
        string plant;
        string workDate;
        int maxDepth;
        int maxRows;
        string fetchedAt;
        int rowCount;
        using (var reader = header.ExecuteReader())
        {
            if (!reader.Read())
                return null;

            modelName = reader.GetString(0);
            plant = reader.GetString(1);
            workDate = reader.GetString(2);
            maxDepth = reader.GetInt32(3);
            maxRows = reader.GetInt32(4);
            fetchedAt = reader.GetString(5);
            rowCount = reader.GetInt32(6);
        }

        using var rowsCmd = conn.CreateCommand();
        rowsCmd.CommandText =
            """
            SELECT
                ProductCode, ProductName, ParentMaterialCode, ParentMaterialName,
                MaterialCode, MaterialName, UsageQty, UsageUnit, BomLevel,
                BomPath, BomPathText, HasChildren, SourceRows, Source
            FROM BmesBomCacheRows
            WHERE CacheKey = @key
            ORDER BY RowNo;
            """;
        rowsCmd.Parameters.AddWithValue("@key", cacheKey);

        var rows = new List<BmesBomMaterialCandidate>(Math.Max(rowCount, 0));
        using (var reader = rowsCmd.ExecuteReader())
        {
            while (reader.Read())
            {
                rows.Add(new BmesBomMaterialCandidate
                {
                    ProductCode = ReadText(reader, 0),
                    ProductName = ReadText(reader, 1),
                    ParentMaterialCode = ReadText(reader, 2),
                    ParentMaterialName = ReadText(reader, 3),
                    MaterialCode = ReadText(reader, 4),
                    MaterialName = ReadText(reader, 5),
                    UsageQty = ParseDecimal(ReadText(reader, 6)),
                    UsageUnit = ReadText(reader, 7),
                    BomLevel = reader.IsDBNull(8) ? 0 : reader.GetInt32(8),
                    BomPath = ReadText(reader, 9),
                    BomPathText = ReadText(reader, 10),
                    HasChildren = !reader.IsDBNull(11) && reader.GetInt32(11) != 0,
                    SourceRows = reader.IsDBNull(12) ? 0 : reader.GetInt64(12),
                    Source = ReadText(reader, 13),
                });
            }
        }

        return new BmesBomCacheEntry
        {
            Query = new BmesBomCacheQuery(
                modelName,
                plant,
                ParseDate(workDate),
                maxDepth,
                maxRows),
            FetchedAt = fetchedAt,
            Rows = rows,
        };
    }

    public void Save(BmesBomCacheQuery query, IReadOnlyList<BmesBomMaterialCandidate> rows)
    {
        lock (_writeGate)
        {
            EnsureDatabaseCore();
            string cacheKey = BuildCacheKey(query);
            string fetchedAt = DateTime.Now.ToString("O", CultureInfo.InvariantCulture);

            using var conn = OpenConnection(readOnly: false);
            using var tx = conn.BeginTransaction();

            using (var delete = conn.CreateCommand())
            {
                delete.Transaction = tx;
                delete.CommandText = "DELETE FROM BmesBomCacheRows WHERE CacheKey = @key;";
                delete.Parameters.AddWithValue("@key", cacheKey);
                delete.ExecuteNonQuery();
            }

            using (var header = conn.CreateCommand())
            {
                header.Transaction = tx;
                header.CommandText =
                    """
                    INSERT INTO BmesBomCacheHeaders
                        (CacheKey, ModelName, Plant, WorkDate, MaxDepth, MaxRows, FetchedAt, RowCount)
                    VALUES
                        (@key, @model, @plant, @workDate, @maxDepth, @maxRows, @fetchedAt, @rowCount)
                    ON CONFLICT(CacheKey) DO UPDATE SET
                        ModelName = excluded.ModelName,
                        Plant = excluded.Plant,
                        WorkDate = excluded.WorkDate,
                        MaxDepth = excluded.MaxDepth,
                        MaxRows = excluded.MaxRows,
                        FetchedAt = excluded.FetchedAt,
                        RowCount = excluded.RowCount;
                    """;
                header.Parameters.AddWithValue("@key", cacheKey);
                header.Parameters.AddWithValue("@model", Normalize(query.ModelName));
                header.Parameters.AddWithValue("@plant", Normalize(query.Plant));
                header.Parameters.AddWithValue("@workDate", query.WorkDate.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
                header.Parameters.AddWithValue("@maxDepth", NormalizeDepth(query.MaxDepth));
                header.Parameters.AddWithValue("@maxRows", NormalizeMaxRows(query.MaxRows));
                header.Parameters.AddWithValue("@fetchedAt", fetchedAt);
                header.Parameters.AddWithValue("@rowCount", rows.Count);
                header.ExecuteNonQuery();
            }

            using var insert = conn.CreateCommand();
            insert.Transaction = tx;
            insert.CommandText =
                """
                INSERT INTO BmesBomCacheRows
                    (CacheKey, RowNo, ProductCode, ProductName, ParentMaterialCode,
                     ParentMaterialName, MaterialCode, MaterialName, UsageQty, UsageUnit,
                     BomLevel, BomPath, BomPathText, HasChildren, SourceRows, Source)
                VALUES
                    (@key, @rowNo, @productCode, @productName, @parentCode,
                     @parentName, @materialCode, @materialName, @usageQty, @usageUnit,
                     @level, @path, @pathText, @hasChildren, @sourceRows, @source);
                """;

            SqliteParameter Key(string name, SqliteType type) => insert.Parameters.Add(name, type);
            var pKey = Key("@key", SqliteType.Text);
            var pRowNo = Key("@rowNo", SqliteType.Integer);
            var pProductCode = Key("@productCode", SqliteType.Text);
            var pProductName = Key("@productName", SqliteType.Text);
            var pParentCode = Key("@parentCode", SqliteType.Text);
            var pParentName = Key("@parentName", SqliteType.Text);
            var pMaterialCode = Key("@materialCode", SqliteType.Text);
            var pMaterialName = Key("@materialName", SqliteType.Text);
            var pUsageQty = Key("@usageQty", SqliteType.Text);
            var pUsageUnit = Key("@usageUnit", SqliteType.Text);
            var pLevel = Key("@level", SqliteType.Integer);
            var pPath = Key("@path", SqliteType.Text);
            var pPathText = Key("@pathText", SqliteType.Text);
            var pHasChildren = Key("@hasChildren", SqliteType.Integer);
            var pSourceRows = Key("@sourceRows", SqliteType.Integer);
            var pSource = Key("@source", SqliteType.Text);
            insert.Prepare();

            for (int i = 0; i < rows.Count; i++)
            {
                BmesBomMaterialCandidate row = rows[i];
                pKey.Value = cacheKey;
                pRowNo.Value = i + 1;
                pProductCode.Value = row.ProductCode;
                pProductName.Value = row.ProductName;
                pParentCode.Value = row.ParentMaterialCode;
                pParentName.Value = row.ParentMaterialName;
                pMaterialCode.Value = row.MaterialCode;
                pMaterialName.Value = row.MaterialName;
                pUsageQty.Value = row.UsageQty?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
                pUsageUnit.Value = row.UsageUnit;
                pLevel.Value = row.BomLevel;
                pPath.Value = row.BomPath;
                pPathText.Value = row.BomPathText;
                pHasChildren.Value = row.HasChildren ? 1 : 0;
                pSourceRows.Value = row.SourceRows;
                pSource.Value = row.Source;
                insert.ExecuteNonQuery();
            }

            tx.Commit();
        }
    }

    /// <summary>Models with at least one BOM row in any successful local cache entry.</summary>
    public List<BmesBomCachedModel> ListSuccessfulModels()
    {
        EnsureDatabase();
        using var conn = OpenConnection(readOnly: true);
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            """
            SELECT h.ModelName, h.RowCount, h.FetchedAt
            FROM BmesBomCacheHeaders AS h
            JOIN (
                SELECT ModelName, MAX(FetchedAt) AS LatestFetchedAt
                FROM BmesBomCacheHeaders
                WHERE RowCount > 0
                GROUP BY ModelName
            ) AS latest
              ON latest.ModelName = h.ModelName
             AND latest.LatestFetchedAt = h.FetchedAt
            WHERE h.RowCount > 0
            ORDER BY h.ModelName;
            """;

        var result = new List<BmesBomCachedModel>();
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new BmesBomCachedModel(
                ReadText(reader, 0),
                reader.IsDBNull(1) ? 0 : reader.GetInt32(1),
                ReadText(reader, 2)));
        }

        return result;
    }

    /// <summary>
    /// Loads the locally synchronized BMES product-name catalog for one plant. This is the
    /// source used by Test 5 while the user types; no network connection is opened here.
    /// </summary>
    public BmesBomModelCatalogEntry LoadModelCatalog(string plant)
    {
        EnsureDatabase();
        plant = Normalize(plant);

        using var conn = OpenConnection(readOnly: true);
        using var statusCmd = conn.CreateCommand();
        statusCmd.CommandText =
            """
            SELECT WorkDate, SyncedAt, RowCount
            FROM BmesBomModelCatalogStatus
            WHERE Plant = @plant;
            """;
        statusCmd.Parameters.AddWithValue("@plant", plant);

        string workDate = string.Empty;
        string syncedAt = string.Empty;
        int rowCount = 0;
        using (var reader = statusCmd.ExecuteReader())
        {
            if (reader.Read())
            {
                workDate = ReadText(reader, 0);
                syncedAt = ReadText(reader, 1);
                rowCount = reader.IsDBNull(2) ? 0 : reader.GetInt32(2);
            }
        }

        using var rowsCmd = conn.CreateCommand();
        rowsCmd.CommandText =
            """
            SELECT ProductCode, ProductName
            FROM BmesBomModelCatalog
            WHERE Plant = @plant
            ORDER BY ProductName, ProductCode;
            """;
        rowsCmd.Parameters.AddWithValue("@plant", plant);

        var models = new List<BmesBomModelCandidate>(Math.Max(rowCount, 0));
        using (var reader = rowsCmd.ExecuteReader())
        {
            while (reader.Read())
            {
                models.Add(new BmesBomModelCandidate
                {
                    ProductCode = ReadText(reader, 0),
                    ProductName = ReadText(reader, 1),
                });
            }
        }

        return new BmesBomModelCatalogEntry
        {
            Plant = plant,
            WorkDate = ParseDate(workDate),
            SyncedAt = syncedAt,
            Models = models,
        };
    }

    public void SaveModelCatalog(
        string plant,
        DateTime workDate,
        IReadOnlyList<BmesBomModelCandidate> models)
    {
        lock (_writeGate)
        {
            EnsureDatabaseCore();
            plant = Normalize(plant);
            string syncedAt = DateTime.Now.ToString("O", CultureInfo.InvariantCulture);

            var uniqueModels = models
                .Where(model =>
                    !string.IsNullOrWhiteSpace(model.ProductCode) &&
                    !string.IsNullOrWhiteSpace(model.ProductName))
                .GroupBy(
                    model => Normalize(model.ProductCode) + "\t" + Normalize(model.ProductName),
                    StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .OrderBy(model => model.ProductName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(model => model.ProductCode, StringComparer.OrdinalIgnoreCase)
                .ToList();

            using var conn = OpenConnection(readOnly: false);
            using var tx = conn.BeginTransaction();

            using (var delete = conn.CreateCommand())
            {
                delete.Transaction = tx;
                delete.CommandText = "DELETE FROM BmesBomModelCatalog WHERE Plant = @plant;";
                delete.Parameters.AddWithValue("@plant", plant);
                delete.ExecuteNonQuery();
            }

            using var insert = conn.CreateCommand();
            insert.Transaction = tx;
            insert.CommandText =
                """
                INSERT INTO BmesBomModelCatalog
                    (Plant, ProductCode, ProductName)
                VALUES
                    (@plant, @productCode, @productName);
                """;
            var pPlant = insert.Parameters.Add("@plant", SqliteType.Text);
            var pProductCode = insert.Parameters.Add("@productCode", SqliteType.Text);
            var pProductName = insert.Parameters.Add("@productName", SqliteType.Text);
            insert.Prepare();

            foreach (BmesBomModelCandidate model in uniqueModels)
            {
                pPlant.Value = plant;
                pProductCode.Value = Normalize(model.ProductCode);
                pProductName.Value = Normalize(model.ProductName);
                insert.ExecuteNonQuery();
            }

            using (var status = conn.CreateCommand())
            {
                status.Transaction = tx;
                status.CommandText =
                    """
                    INSERT INTO BmesBomModelCatalogStatus
                        (Plant, WorkDate, SyncedAt, RowCount)
                    VALUES
                        (@plant, @workDate, @syncedAt, @rowCount)
                    ON CONFLICT(Plant) DO UPDATE SET
                        WorkDate = excluded.WorkDate,
                        SyncedAt = excluded.SyncedAt,
                        RowCount = excluded.RowCount;
                    """;
                status.Parameters.AddWithValue("@plant", plant);
                status.Parameters.AddWithValue(
                    "@workDate",
                    workDate.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
                status.Parameters.AddWithValue("@syncedAt", syncedAt);
                status.Parameters.AddWithValue("@rowCount", uniqueModels.Count);
                status.ExecuteNonQuery();
            }

            tx.Commit();
        }
    }

    private void EnsureDatabase()
    {
        lock (_writeGate)
            EnsureDatabaseCore();
    }

    private void EnsureDatabaseCore()
    {
        Directory.CreateDirectory(settings.SettingsDbDirectory);
        using var conn = OpenConnection(readOnly: false);
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            """
            PRAGMA journal_mode=WAL;
            PRAGMA busy_timeout=5000;

            CREATE TABLE IF NOT EXISTS BmesBomCacheHeaders (
                CacheKey TEXT PRIMARY KEY,
                ModelName TEXT NOT NULL COLLATE NOCASE,
                Plant TEXT NOT NULL,
                WorkDate TEXT NOT NULL,
                MaxDepth INTEGER NOT NULL,
                MaxRows INTEGER NOT NULL,
                FetchedAt TEXT NOT NULL,
                RowCount INTEGER NOT NULL
            );

            CREATE TABLE IF NOT EXISTS BmesBomCacheRows (
                CacheKey TEXT NOT NULL,
                RowNo INTEGER NOT NULL,
                ProductCode TEXT,
                ProductName TEXT,
                ParentMaterialCode TEXT,
                ParentMaterialName TEXT,
                MaterialCode TEXT,
                MaterialName TEXT,
                UsageQty TEXT,
                UsageUnit TEXT,
                BomLevel INTEGER,
                BomPath TEXT,
                BomPathText TEXT,
                HasChildren INTEGER,
                SourceRows INTEGER,
                Source TEXT,
                PRIMARY KEY (CacheKey, RowNo)
            );

            CREATE INDEX IF NOT EXISTS IX_BmesBomCacheHeaders_Model
                ON BmesBomCacheHeaders(ModelName, FetchedAt);
            CREATE INDEX IF NOT EXISTS IX_BmesBomCacheRows_Key
                ON BmesBomCacheRows(CacheKey, RowNo);

            CREATE TABLE IF NOT EXISTS BmesBomModelCatalogStatus (
                Plant TEXT PRIMARY KEY COLLATE NOCASE,
                WorkDate TEXT NOT NULL,
                SyncedAt TEXT NOT NULL,
                RowCount INTEGER NOT NULL
            );

            CREATE TABLE IF NOT EXISTS BmesBomModelCatalog (
                Plant TEXT NOT NULL COLLATE NOCASE,
                ProductCode TEXT NOT NULL COLLATE NOCASE,
                ProductName TEXT NOT NULL COLLATE NOCASE,
                PRIMARY KEY (Plant, ProductCode, ProductName)
            );

            CREATE INDEX IF NOT EXISTS IX_BmesBomModelCatalog_Search
                ON BmesBomModelCatalog(Plant, ProductName, ProductCode);
            """;
        cmd.ExecuteNonQuery();
    }

    private SqliteConnection OpenConnection(bool readOnly)
    {
        var conn = new SqliteConnection(
            readOnly
                ? $"Data Source={CacheDbPath};Mode=ReadOnly"
                : $"Data Source={CacheDbPath}");
        conn.Open();
        using var pragma = conn.CreateCommand();
        pragma.CommandText = "PRAGMA busy_timeout=5000;";
        pragma.ExecuteNonQuery();
        return conn;
    }

    private static string BuildCacheKey(BmesBomCacheQuery query)
    {
        // WorkDate is metadata, not part of the identity: once a model BOM has been loaded,
        // the page should keep using that model cache on later days until the user explicitly
        // checks "Load BOM from server" to refresh it.
        string canonical = string.Join(
            "\u001f",
            Normalize(query.ModelName).ToUpperInvariant(),
            Normalize(query.Plant).ToUpperInvariant(),
            NormalizeDepth(query.MaxDepth).ToString(CultureInfo.InvariantCulture),
            NormalizeMaxRows(query.MaxRows).ToString(CultureInfo.InvariantCulture));
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(canonical)));
    }

    private static int NormalizeDepth(int depth) => Math.Clamp(depth <= 0 ? 6 : depth, 1, 12);

    private static int NormalizeMaxRows(int maxRows) => Math.Clamp(maxRows <= 0 ? 2000 : maxRows, 1, 2000);

    private static string Normalize(string? value) => (value ?? string.Empty).Trim();

    private static string ReadText(SqliteDataReader reader, int ordinal) =>
        reader.IsDBNull(ordinal) ? string.Empty : reader.GetString(ordinal);

    private static decimal? ParseDecimal(string value) =>
        decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal parsed)
            ? parsed
            : null;

    private static DateTime ParseDate(string value) =>
        DateTime.TryParseExact(
            value,
            "yyyy-MM-dd",
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out DateTime parsed)
            ? parsed.Date
            : DateTime.MinValue;
}

public sealed record BmesBomCacheQuery(
    string ModelName,
    string Plant,
    DateTime WorkDate,
    int MaxDepth,
    int MaxRows);

public sealed class BmesBomCacheEntry
{
    public BmesBomCacheQuery Query { get; init; } =
        new(string.Empty, string.Empty, DateTime.MinValue, 1, 1);
    public string FetchedAt { get; init; } = string.Empty;
    public List<BmesBomMaterialCandidate> Rows { get; init; } = [];
}

public sealed record BmesBomCachedModel(string ModelName, int RowCount, string FetchedAt);

public sealed class BmesBomModelCatalogEntry
{
    public string Plant { get; init; } = string.Empty;
    public DateTime WorkDate { get; init; }
    public string SyncedAt { get; init; } = string.Empty;
    public List<BmesBomModelCandidate> Models { get; init; } = [];
}
