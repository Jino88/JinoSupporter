using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Data.Sqlite;
using WorkbenchHost.Infrastructure;

namespace JinoSupporter.App.Modules.DataInference;

// ── Column definition ──────────────────────────────────────────────────────

public sealed class ColumnDef
{
    [JsonPropertyName("field")] public string Field { get; set; } = string.Empty;
    [JsonPropertyName("label")] public string Label { get; set; } = string.Empty;
}

// ── Table metadata ─────────────────────────────────────────────────────────

public sealed class DataTableInfo
{
    public long           Id          { get; init; }
    public string         DatasetName { get; init; } = string.Empty;
    public string         TableName   { get; init; } = string.Empty;
    public List<ColumnDef> Columns    { get; init; } = [];
    public string         CreatedAt   { get; init; } = string.Empty;
    public int            RowCount    { get; init; }

    public string DisplayLabel => $"{TableName}  ({RowCount:N0} rows)";

    public string CreatedAtLocal
    {
        get
        {
            if (DateTime.TryParse(CreatedAt, null,
                    System.Globalization.DateTimeStyles.RoundtripKind, out DateTime dt))
                return dt.ToLocalTime().ToString("yyyy-MM-dd HH:mm");
            return CreatedAt;
        }
    }
}

// ── Image row ──────────────────────────────────────────────────────────────

public sealed class DatasetImageRow
{
    public long   Id          { get; init; }
    public string DatasetName { get; init; } = string.Empty;
    public string FileName    { get; init; } = string.Empty;
    public byte[] ImageData   { get; init; } = [];
    public string CreatedAt   { get; init; } = string.Empty;
}

public sealed class MfgTestRecordSave
{
    public string Title { get; set; } = string.Empty;
    public int RecordNo { get; set; }
    public string TestDate { get; set; } = string.Empty;
    public string Shift { get; set; } = string.Empty;
    public string Model { get; set; } = string.Empty;
    public string TestType { get; set; } = string.Empty;
    public int Input { get; set; }
    public int Ok { get; set; }
    public int TotalNG { get; set; }
    public double NGRate { get; set; }
    public int NgSpl { get; set; }
    public int NgSplRb { get; set; }
    public int NgNoSound { get; set; }
    public int NgNoise { get; set; }
    public int NgTouch { get; set; }
    public int NgHearing { get; set; }
}

public sealed class MfgTestRecordRow
{
    public long Id { get; init; }
    public int RecordNo { get; init; }
    public string TestDate { get; init; } = string.Empty;
    public string Shift { get; init; } = string.Empty;
    public string Model { get; init; } = string.Empty;
    public string TestType { get; init; } = string.Empty;
    public int Input { get; init; }
    public int Ok { get; init; }
    public int TotalNG { get; init; }
    public string NGRateDisplay { get; init; } = string.Empty;
    public int NgSpl { get; init; }
    public int NgSplRb { get; init; }
    public int NgNoSound { get; init; }
    public int NgNoise { get; init; }
    public int NgTouch { get; init; }
    public int NgHearing { get; init; }
    public string BatchId { get; init; } = string.Empty;
    public string ImportedAtLocal { get; init; } = string.Empty;
}

// ── Repository ─────────────────────────────────────────────────────────────

public sealed class DataInferenceRepository
{
    private static readonly JsonSerializerOptions JsonOpts = new() { PropertyNameCaseInsensitive = true };

    public static string DefaultDatabasePath =>
        Path.Combine(WorkbenchSettingsStore.DefaultStorageRootDirectory, "process-review.db");

    public static string DatabasePath
    {
        get
        {
            string custom = WorkbenchSettingsStore.GetDataInferenceDatabasePath();
            return string.IsNullOrWhiteSpace(custom) ? DefaultDatabasePath : custom;
        }
    }

    private static readonly HashSet<string> _initializedPaths = new(StringComparer.OrdinalIgnoreCase);
    private static readonly object           _schemaLock       = new();

    public DataInferenceRepository()
    {
        EnsureSchemaForCurrentPath();
    }

    // ── Schema ─────────────────────────────────────────────────────────────

    private static void EnsureDatabase(SqliteConnection conn)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            CREATE TABLE IF NOT EXISTS DataTables (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                DatasetName TEXT    NOT NULL,
                TableName   TEXT    NOT NULL DEFAULT '',
                Columns     TEXT    NOT NULL DEFAULT '[]',
                CreatedAt   TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_dt_dataset ON DataTables(DatasetName);

            CREATE TABLE IF NOT EXISTS DataTableRows (
                Id      INTEGER PRIMARY KEY AUTOINCREMENT,
                TableId INTEGER NOT NULL REFERENCES DataTables(Id),
                RowData TEXT    NOT NULL DEFAULT '{}'
            );
            CREATE INDEX IF NOT EXISTS idx_dtr_table ON DataTableRows(TableId);

            CREATE TABLE IF NOT EXISTS DatasetImages (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                DatasetName TEXT    NOT NULL,
                FileName    TEXT    NOT NULL DEFAULT '',
                ImageData   BLOB    NOT NULL,
                CreatedAt   TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_di_dataset ON DatasetImages(DatasetName);

            CREATE TABLE IF NOT EXISTS DatasetTags (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                DatasetName TEXT    NOT NULL UNIQUE,
                Tags        TEXT    NOT NULL DEFAULT '[]',
                CreatedAt   TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_dtag_dataset ON DatasetTags(DatasetName);

            CREATE TABLE IF NOT EXISTS Glossary (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                Term        TEXT    NOT NULL UNIQUE,
                Description TEXT    NOT NULL DEFAULT '',
                UpdatedAt   TEXT    NOT NULL
            );

            CREATE TABLE IF NOT EXISTS DatasetMemo (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                DatasetName TEXT    NOT NULL UNIQUE,
                Memo        TEXT    NOT NULL DEFAULT '',
                UpdatedAt   TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_dmemo_dataset ON DatasetMemo(DatasetName);

            CREATE TABLE IF NOT EXISTS Reports (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                Title       TEXT    NOT NULL DEFAULT '',
                DatasetNames TEXT   NOT NULL DEFAULT '',
                HtmlContent TEXT    NOT NULL DEFAULT '',
                CreatedAt   TEXT    NOT NULL
            );
            """;
        cmd.ExecuteNonQuery();
        MigrateSchema(conn);
    }

    private static void MigrateSchema(SqliteConnection conn)
    {
        // Add DatasetNames column to Reports if it was created before this column existed
        using SqliteCommand check = conn.CreateCommand();
        check.CommandText = "PRAGMA table_info(Reports);";
        bool hasDatasetNames = false;
        using (SqliteDataReader r = check.ExecuteReader())
        {
            while (r.Read())
            {
                if (r.GetString(1).Equals("DatasetNames", StringComparison.OrdinalIgnoreCase))
                {
                    hasDatasetNames = true;
                    break;
                }
            }
        }
        if (!hasDatasetNames)
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE Reports ADD COLUMN DatasetNames TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }
    }

    // ── Write ──────────────────────────────────────────────────────────────

    public long SaveTable(string datasetName, string tableName,
                          List<ColumnDef> columns, List<Dictionary<string, string>> rows)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();

        string colJson = JsonSerializer.Serialize(columns, JsonOpts);

        using SqliteCommand ins = conn.CreateCommand();
        ins.Transaction = tx;
        ins.CommandText = """
            INSERT INTO DataTables (DatasetName, TableName, Columns, CreatedAt)
            VALUES (@d, @n, @c, @at);
            SELECT last_insert_rowid();
            """;
        ins.Parameters.AddWithValue("@d", datasetName);
        ins.Parameters.AddWithValue("@n", tableName);
        ins.Parameters.AddWithValue("@c", colJson);
        ins.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture));
        long tableId = Convert.ToInt64(ins.ExecuteScalar() ?? 0);

        foreach (Dictionary<string, string> row in rows)
        {
            using SqliteCommand insRow = conn.CreateCommand();
            insRow.Transaction = tx;
            insRow.CommandText = "INSERT INTO DataTableRows (TableId, RowData) VALUES (@t, @r);";
            insRow.Parameters.AddWithValue("@t", tableId);
            insRow.Parameters.AddWithValue("@r", JsonSerializer.Serialize(row, JsonOpts));
            insRow.ExecuteNonQuery();
        }

        tx.Commit();
        return tableId;
    }

    public void SaveMemo(string datasetName, string memo)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO DatasetMemo (DatasetName, Memo, UpdatedAt)
            VALUES (@d, @m, @at)
            ON CONFLICT(DatasetName) DO UPDATE SET Memo=@m, UpdatedAt=@at;
            """;
        cmd.Parameters.AddWithValue("@d", datasetName);
        cmd.Parameters.AddWithValue("@m", memo);
        cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture));
        cmd.ExecuteNonQuery();
    }

    public string GetMemo(string datasetName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Memo FROM DatasetMemo WHERE DatasetName=@d;";
        cmd.Parameters.AddWithValue("@d", datasetName);
        using SqliteDataReader r = cmd.ExecuteReader();
        return r.Read() ? r.GetString(0) : string.Empty;
    }

    public void SaveTags(string datasetName, List<string> tags)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO DatasetTags (DatasetName, Tags, CreatedAt)
            VALUES (@d, @t, @at)
            ON CONFLICT(DatasetName) DO UPDATE SET Tags=@t, CreatedAt=@at;
            """;
        cmd.Parameters.AddWithValue("@d", datasetName);
        cmd.Parameters.AddWithValue("@t", JsonSerializer.Serialize(tags));
        cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture));
        cmd.ExecuteNonQuery();
    }

    public List<string> GetTags(string datasetName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Tags FROM DatasetTags WHERE DatasetName=@d;";
        cmd.Parameters.AddWithValue("@d", datasetName);
        using SqliteDataReader r = cmd.ExecuteReader();
        if (r.Read())
        {
            try { return JsonSerializer.Deserialize<List<string>>(r.GetString(0)) ?? []; } catch { }
        }
        return [];
    }

    // ── Glossary ───────────────────────────────────────────────────────────

    /// <summary>Replace entire Glossary table with the given entries (used by GlossaryWindow save).</summary>
    public void ReplaceGlossary(IEnumerable<(string Term, string Description)> entries)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();
        string now = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);

        using SqliteCommand del = conn.CreateCommand();
        del.Transaction = tx;
        del.CommandText = "DELETE FROM Glossary;";
        del.ExecuteNonQuery();

        foreach ((string term, string desc) in entries)
        {
            if (string.IsNullOrWhiteSpace(term)) continue;
            using SqliteCommand ins = conn.CreateCommand();
            ins.Transaction = tx;
            ins.CommandText = "INSERT INTO Glossary (Term, Description, UpdatedAt) VALUES (@t, @d, @at);";
            ins.Parameters.AddWithValue("@t", term.Trim());
            ins.Parameters.AddWithValue("@d", desc.Trim());
            ins.Parameters.AddWithValue("@at", now);
            ins.ExecuteNonQuery();
        }

        tx.Commit();
    }

    /// <summary>Merge lines of "term: description" into the global Glossary table (upsert).</summary>
    public void MergeGlossary(string glossaryText)
    {
        if (string.IsNullOrWhiteSpace(glossaryText)) return;

        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();
        string now = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);

        foreach (string line in glossaryText.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            int colon = line.IndexOf(':', StringComparison.Ordinal);
            if (colon <= 0) continue;

            string term = line[..colon].Trim();
            string desc = line[(colon + 1)..].Trim();
            if (string.IsNullOrWhiteSpace(term)) continue;

            using SqliteCommand cmd = conn.CreateCommand();
            cmd.Transaction = tx;
            cmd.CommandText = """
                INSERT INTO Glossary (Term, Description, UpdatedAt) VALUES (@t, @d, @at)
                ON CONFLICT(Term) DO UPDATE SET Description=@d, UpdatedAt=@at;
                """;
            cmd.Parameters.AddWithValue("@t", term);
            cmd.Parameters.AddWithValue("@d", desc);
            cmd.Parameters.AddWithValue("@at", now);
            cmd.ExecuteNonQuery();
        }

        tx.Commit();
    }

    /// <summary>Returns all glossary entries as "term: description" lines (for Claude prompts).</summary>
    public string GetGlossaryText()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Term, Description FROM Glossary ORDER BY Term;";

        var lines = new List<string>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
            lines.Add($"{r.GetString(0)}: {r.GetString(1)}");

        return string.Join("\n", lines);
    }

    /// <summary>Returns all glossary entries sorted by most recently updated first.</summary>
    public List<(string Term, string Description)> GetGlossaryEntriesNewestFirst()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Term, Description FROM Glossary ORDER BY UpdatedAt DESC;";

        var list = new List<(string, string)>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
            list.Add((r.GetString(0), r.GetString(1)));

        return list;
    }

    public List<string> GetAllDistinctTags()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Tags FROM DatasetTags;";

        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            try
            {
                var tags = JsonSerializer.Deserialize<List<string>>(r.GetString(0)) ?? [];
                foreach (string t in tags)
                    if (!string.IsNullOrWhiteSpace(t)) set.Add(t);
            }
            catch { }
        }

        return [.. set.OrderBy(t => t, StringComparer.OrdinalIgnoreCase)];
    }

    /// <summary>Bulk-replaces oldTag with newTag across all DatasetTags rows.</summary>
    public int RenameTag(string oldTag, string newTag)
    {
        if (string.IsNullOrWhiteSpace(oldTag) || string.IsNullOrWhiteSpace(newTag)) return 0;
        oldTag = oldTag.Trim();
        newTag = newTag.Trim();

        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id, Tags FROM DatasetTags;";

        // Affected DatasetName → new Tags JSON map
        var updates = new List<(long Id, string NewTagsJson)>();
        using (SqliteDataReader r = cmd.ExecuteReader())
        {
            while (r.Read())
            {
                long id = r.GetInt64(0);
                try
                {
                    List<string> tags = JsonSerializer.Deserialize<List<string>>(r.GetString(1)) ?? [];
                    bool changed = false;
                    for (int i = 0; i < tags.Count; i++)
                    {
                        if (string.Equals(tags[i], oldTag, StringComparison.OrdinalIgnoreCase))
                        {
                            tags[i] = newTag;
                            changed = true;
                        }
                    }
                    if (changed)
                        updates.Add((id, JsonSerializer.Serialize(tags)));
                }
                catch { }
            }
        }

        using SqliteTransaction tx = conn.BeginTransaction();
        foreach ((long id, string json) in updates)
        {
            using SqliteCommand upd = conn.CreateCommand();
            upd.Transaction = tx;
            upd.CommandText = "UPDATE DatasetTags SET Tags=@t WHERE Id=@id;";
            upd.Parameters.AddWithValue("@t", json);
            upd.Parameters.AddWithValue("@id", id);
            upd.ExecuteNonQuery();
        }
        tx.Commit();
        return updates.Count;
    }

    /// <summary>Returns a list of DatasetNames that include the specified tags (returns all if the list is empty).</summary>
    public List<string> GetDatasetsByTags(IReadOnlyList<string> tags)
    {
        if (tags.Count == 0)
        {
            // All distinct DatasetNames
            using SqliteConnection conn2 = OpenConnection();
            using SqliteCommand cmd2 = conn2.CreateCommand();
            cmd2.CommandText = "SELECT DISTINCT DatasetName FROM DataTables ORDER BY DatasetName;";
            var all = new List<string>();
            using SqliteDataReader r2 = cmd2.ExecuteReader();
            while (r2.Read()) all.Add(r2.GetString(0));
            return all;
        }

        // Collect DatasetNames from DatasetTags that contain at least one of the specified tags
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT DatasetName, Tags FROM DatasetTags;";

        var matchSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var tagSet   = new HashSet<string>(tags, StringComparer.OrdinalIgnoreCase);

        using (SqliteDataReader r = cmd.ExecuteReader())
        {
            while (r.Read())
            {
                string dsName = r.GetString(0);
                try
                {
                    List<string> dsTags = JsonSerializer.Deserialize<List<string>>(r.GetString(1)) ?? [];
                    if (dsTags.Any(tagSet.Contains))
                        matchSet.Add(dsName);
                }
                catch { }
            }
        }

        return [.. matchSet.OrderBy(n => n, StringComparer.OrdinalIgnoreCase)];
    }

    public void UpdateTableRows(long tableId, List<ColumnDef> columns, System.Data.DataTable dt)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();

        using SqliteCommand del = conn.CreateCommand();
        del.Transaction = tx;
        del.CommandText = "DELETE FROM DataTableRows WHERE TableId=@id;";
        del.Parameters.AddWithValue("@id", tableId);
        del.ExecuteNonQuery();

        foreach (System.Data.DataRow dr in dt.Rows)
        {
            if (dr.RowState == System.Data.DataRowState.Deleted) continue;
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (ColumnDef col in columns)
                dict[col.Field] = dr[col.Field]?.ToString() ?? string.Empty;

            using SqliteCommand ins = conn.CreateCommand();
            ins.Transaction = tx;
            ins.CommandText = "INSERT INTO DataTableRows (TableId, RowData) VALUES (@t, @r);";
            ins.Parameters.AddWithValue("@t", tableId);
            ins.Parameters.AddWithValue("@r", JsonSerializer.Serialize(dict));
            ins.ExecuteNonQuery();
        }

        tx.Commit();
    }

    public void RenameTable(long tableId, string newName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE DataTables SET TableName=@n WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@n",   newName);
        cmd.Parameters.AddWithValue("@id",  tableId);
        cmd.ExecuteNonQuery();
    }

    public void DeleteColumn(long tableId, string fieldName)
    {
        using SqliteConnection conn  = OpenConnection();
        using SqliteTransaction tx   = conn.BeginTransaction();

        // 1. Remove from columns JSON
        using SqliteCommand readCmd = conn.CreateCommand();
        readCmd.Transaction = tx;
        readCmd.CommandText = "SELECT Columns FROM DataTables WHERE Id=@id;";
        readCmd.Parameters.AddWithValue("@id", tableId);
        string colJson = (string?)readCmd.ExecuteScalar() ?? "[]";

        List<ColumnDef> cols = JsonSerializer.Deserialize<List<ColumnDef>>(colJson, JsonOpts) ?? [];
        cols.RemoveAll(c => string.Equals(c.Field, fieldName, StringComparison.OrdinalIgnoreCase));

        using SqliteCommand updCols = conn.CreateCommand();
        updCols.Transaction = tx;
        updCols.CommandText = "UPDATE DataTables SET Columns=@c WHERE Id=@id;";
        updCols.Parameters.AddWithValue("@c",  JsonSerializer.Serialize(cols, JsonOpts));
        updCols.Parameters.AddWithValue("@id", tableId);
        updCols.ExecuteNonQuery();

        // 2. Remove field from every row JSON
        using SqliteCommand rowsCmd = conn.CreateCommand();
        rowsCmd.Transaction = tx;
        rowsCmd.CommandText = "SELECT Id, RowData FROM DataTableRows WHERE TableId=@t;";
        rowsCmd.Parameters.AddWithValue("@t", tableId);

        var rowUpdates = new List<(long Id, string Json)>();
        using (SqliteDataReader r = rowsCmd.ExecuteReader())
        {
            while (r.Read())
            {
                long   rowId   = r.GetInt64(0);
                string rowJson = r.GetString(1);
                var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(rowJson, JsonOpts)
                           ?? [];
                dict.Remove(fieldName);
                rowUpdates.Add((rowId, JsonSerializer.Serialize(dict, JsonOpts)));
            }
        }

        using SqliteCommand updRow = conn.CreateCommand();
        updRow.Transaction = tx;
        updRow.CommandText = "UPDATE DataTableRows SET RowData=@d WHERE Id=@id;";
        SqliteParameter pData = updRow.Parameters.Add("@d",  SqliteType.Text);
        SqliteParameter pId   = updRow.Parameters.Add("@id", SqliteType.Integer);
        foreach ((long rowId, string json) in rowUpdates)
        {
            pData.Value = json;
            pId.Value   = rowId;
            updRow.ExecuteNonQuery();
        }

        tx.Commit();
    }

    /// <summary>Saves a report. datasetNames = comma-separated list of dataset names</summary>
    public void SaveReport(string title, string datasetNames, string htmlContent)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO Reports (Title, DatasetNames, HtmlContent, CreatedAt)
            VALUES (@title, @dsNames, @html, @createdAt);
            """;
        cmd.Parameters.AddWithValue("@title",    title);
        cmd.Parameters.AddWithValue("@dsNames",  datasetNames);
        cmd.Parameters.AddWithValue("@html",     htmlContent);
        cmd.Parameters.AddWithValue("@createdAt", DateTime.UtcNow.ToString("o"));
        cmd.ExecuteNonQuery();
    }

    /// <summary>List of reports containing the specified dataset (newest first). Returns all if datasetName is empty.</summary>
    public List<(long Id, string Title, string DatasetNames, string CreatedAt)> GetReportList(string datasetName = "")
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        if (string.IsNullOrWhiteSpace(datasetName))
            cmd.CommandText = "SELECT Id, Title, DatasetNames, CreatedAt FROM Reports ORDER BY Id DESC;";
        else
        {
            cmd.CommandText = "SELECT Id, Title, DatasetNames, CreatedAt FROM Reports WHERE DatasetNames LIKE @pat ORDER BY Id DESC;";
            cmd.Parameters.AddWithValue("@pat", $"%{datasetName}%");
        }

        var list = new List<(long, string, string, string)>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            string utc = r.GetString(3);
            string local = DateTime.TryParse(utc, null,
                System.Globalization.DateTimeStyles.RoundtripKind, out DateTime dt)
                ? dt.ToLocalTime().ToString("yyyy-MM-dd HH:mm")
                : utc;
            list.Add((r.GetInt64(0), r.GetString(1), r.GetString(2), local));
        }
        return list;
    }

    public string GetReportHtml(long reportId)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT HtmlContent FROM Reports WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", reportId);
        return cmd.ExecuteScalar() as string ?? string.Empty;
    }

    public void DeleteReport(long reportId)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM Reports WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", reportId);
        cmd.ExecuteNonQuery();
    }

    public void DeleteTable(long tableId)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();

        using SqliteCommand d1 = conn.CreateCommand();
        d1.Transaction = tx;
        d1.CommandText = "DELETE FROM DataTableRows WHERE TableId=@id;";
        d1.Parameters.AddWithValue("@id", tableId);
        d1.ExecuteNonQuery();

        using SqliteCommand d2 = conn.CreateCommand();
        d2.Transaction = tx;
        d2.CommandText = "DELETE FROM DataTables WHERE Id=@id;";
        d2.Parameters.AddWithValue("@id", tableId);
        d2.ExecuteNonQuery();

        tx.Commit();
    }

    // ── Read ───────────────────────────────────────────────────────────────

    public List<string> GetDistinctDatasets()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT DISTINCT DatasetName FROM DataTables ORDER BY DatasetName;";

        var list = new List<string> { "All" };
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) { string s = r.GetString(0); if (!string.IsNullOrWhiteSpace(s)) list.Add(s); }
        return list;
    }

    public List<(string Name, int TableCount, int ImageCount)> GetDatasetSummary()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT
                d.DatasetName,
                (SELECT COUNT(*) FROM DataTables   WHERE DatasetName = d.DatasetName) AS TableCount,
                (SELECT COUNT(*) FROM DatasetImages WHERE DatasetName = d.DatasetName) AS ImageCount
            FROM (SELECT DISTINCT DatasetName FROM DataTables) d
            ORDER BY d.DatasetName;
            """;
        var list = new List<(string, int, int)>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
            list.Add((r.GetString(0), r.GetInt32(1), r.GetInt32(2)));
        return list;
    }

    public List<DataTableInfo> GetTables(string? datasetName = null)
    {
        bool all = string.IsNullOrWhiteSpace(datasetName) || datasetName == "All";
        string where = all ? string.Empty : "WHERE dt.DatasetName=@d";

        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = $"""
            SELECT dt.Id, dt.DatasetName, dt.TableName, dt.Columns, dt.CreatedAt,
                   COUNT(r.Id) AS RowCount
            FROM DataTables dt
            LEFT JOIN DataTableRows r ON r.TableId = dt.Id
            {where}
            GROUP BY dt.Id
            ORDER BY dt.Id;
            """;
        if (!all) cmd.Parameters.AddWithValue("@d", datasetName!);

        var list = new List<DataTableInfo>();
        using SqliteDataReader r2 = cmd.ExecuteReader();
        while (r2.Read())
        {
            List<ColumnDef> cols = [];
            try { cols = JsonSerializer.Deserialize<List<ColumnDef>>(r2.GetString(3), JsonOpts) ?? []; } catch { }
            list.Add(new DataTableInfo
            {
                Id          = r2.GetInt64(0),
                DatasetName = r2.GetString(1),
                TableName   = r2.GetString(2),
                Columns     = cols,
                CreatedAt   = r2.GetString(4),
                RowCount    = r2.GetInt32(5)
            });
        }
        return list;
    }

    public List<(long Id, Dictionary<string, string> Data)> GetTableRows(long tableId)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id, RowData FROM DataTableRows WHERE TableId=@id ORDER BY Id;";
        cmd.Parameters.AddWithValue("@id", tableId);

        var list = new List<(long, Dictionary<string, string>)>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            long id = r.GetInt64(0);
            Dictionary<string, string> data = [];
            try
            {
                var raw = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(r.GetString(1)) ?? [];
                foreach ((string key, JsonElement val) in raw)
                    data[key] = val.ValueKind == JsonValueKind.String
                        ? (val.GetString() ?? string.Empty)
                        : val.ToString();
            }
            catch { }
            list.Add((id, data));
        }
        return list;
    }

    public int GetTotalRowCount()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COUNT(*) FROM DataTableRows;";
        return Convert.ToInt32(cmd.ExecuteScalar() ?? 0);
    }

    // Compatibility layer for legacy DataInference views

    public int GetTotalCount() => GetTotalRowCount();

    public List<string> GetDistinctModels(string? datasetName = null)
    {
        var models = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (MfgTestRecordRow row in GetRecords(datasetName, null))
        {
            if (!string.IsNullOrWhiteSpace(row.Model))
            {
                models.Add(row.Model);
            }
        }

        List<string> result = ["All"];
        result.AddRange(models.OrderBy(model => model, StringComparer.OrdinalIgnoreCase));
        return result;
    }

    public void SaveBatch(string datasetName, List<MfgTestRecordSave> records)
    {
        if (records.Count == 0)
        {
            return;
        }

        List<ColumnDef> columns =
        [
            new() { Field = "RecordNo", Label = "No" },
            new() { Field = "TestDate", Label = "날짜" },
            new() { Field = "Shift", Label = "근무조" },
            new() { Field = "Model", Label = "모델" },
            new() { Field = "TestType", Label = "검사유형" },
            new() { Field = "Input", Label = "투입" },
            new() { Field = "Ok", Label = "OK" },
            new() { Field = "TotalNG", Label = "NG합계" },
            new() { Field = "NGRate", Label = "NG%" },
            new() { Field = "NgSpl", Label = "SPL" },
            new() { Field = "NgSplRb", Label = "SPL+RB" },
            new() { Field = "NgNoSound", Label = "NoSound" },
            new() { Field = "NgNoise", Label = "Noise" },
            new() { Field = "NgTouch", Label = "Touch" },
            new() { Field = "NgHearing", Label = "Hearing" }
        ];

        List<Dictionary<string, string>> rows = records.Select(record => new Dictionary<string, string>
        {
            ["RecordNo"] = record.RecordNo.ToString(CultureInfo.InvariantCulture),
            ["TestDate"] = record.TestDate,
            ["Shift"] = record.Shift,
            ["Model"] = record.Model,
            ["TestType"] = record.TestType,
            ["Input"] = record.Input.ToString(CultureInfo.InvariantCulture),
            ["Ok"] = record.Ok.ToString(CultureInfo.InvariantCulture),
            ["TotalNG"] = record.TotalNG.ToString(CultureInfo.InvariantCulture),
            ["NGRate"] = record.NGRate.ToString(CultureInfo.InvariantCulture),
            ["NgSpl"] = record.NgSpl.ToString(CultureInfo.InvariantCulture),
            ["NgSplRb"] = record.NgSplRb.ToString(CultureInfo.InvariantCulture),
            ["NgNoSound"] = record.NgNoSound.ToString(CultureInfo.InvariantCulture),
            ["NgNoise"] = record.NgNoise.ToString(CultureInfo.InvariantCulture),
            ["NgTouch"] = record.NgTouch.ToString(CultureInfo.InvariantCulture),
            ["NgHearing"] = record.NgHearing.ToString(CultureInfo.InvariantCulture)
        }).ToList();

        string tableName = string.IsNullOrWhiteSpace(records[0].Title) ? $"Table {DateTime.Now:HHmmss}" : records[0].Title;
        SaveTable(datasetName, tableName, columns, rows);
    }

    public void DeleteById(long rowId)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM DataTableRows WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", rowId);
        cmd.ExecuteNonQuery();
    }

    public List<MfgTestRecordRow> GetRecords(string? datasetName, string? model)
    {
        bool matchAllDatasets = string.IsNullOrWhiteSpace(datasetName) || datasetName == "All";
        bool matchAllModels = string.IsNullOrWhiteSpace(model) || model == "All";

        List<MfgTestRecordRow> rows = [];
        foreach (DataTableInfo table in GetTables(matchAllDatasets ? null : datasetName))
        {
            foreach ((long id, Dictionary<string, string> data) in GetTableRows(table.Id))
            {
                MfgTestRecordRow row = new()
                {
                    Id = id,
                    RecordNo = ParseInt(data, "RecordNo"),
                    TestDate = GetValue(data, "TestDate"),
                    Shift = GetValue(data, "Shift"),
                    Model = GetValue(data, "Model"),
                    TestType = GetValue(data, "TestType"),
                    Input = ParseInt(data, "Input"),
                    Ok = ParseInt(data, "Ok"),
                    TotalNG = ParseInt(data, "TotalNG"),
                    NGRateDisplay = FormatPercent(ParseDouble(data, "NGRate")),
                    NgSpl = ParseInt(data, "NgSpl"),
                    NgSplRb = ParseInt(data, "NgSplRb"),
                    NgNoSound = ParseInt(data, "NgNoSound"),
                    NgNoise = ParseInt(data, "NgNoise"),
                    NgTouch = ParseInt(data, "NgTouch"),
                    NgHearing = ParseInt(data, "NgHearing"),
                    BatchId = table.DatasetName,
                    ImportedAtLocal = table.CreatedAt
                };

                if (!matchAllModels && !string.Equals(row.Model, model, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                rows.Add(row);
            }
        }

        return rows
            .OrderByDescending(row => row.Id)
            .ToList();
    }

    // ── Images ─────────────────────────────────────────────────────────────

    public void SaveImage(string datasetName, string fileName, byte[] data)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "INSERT INTO DatasetImages (DatasetName, FileName, ImageData, CreatedAt) VALUES (@d, @f, @img, @at);";
        cmd.Parameters.AddWithValue("@d", datasetName);
        cmd.Parameters.AddWithValue("@f", fileName);
        cmd.Parameters.AddWithValue("@img", data);
        cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture));
        cmd.ExecuteNonQuery();
    }

    public void DeleteImage(long id)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM DatasetImages WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    public List<DatasetImageRow> GetImages(string datasetName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id, DatasetName, FileName, ImageData, CreatedAt FROM DatasetImages WHERE DatasetName=@name ORDER BY Id;";
        cmd.Parameters.AddWithValue("@name", datasetName);

        var rows = new List<DatasetImageRow>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            long byteLen = r.GetBytes(3, 0, null, 0, 0);
            byte[] buf = new byte[byteLen];
            r.GetBytes(3, 0, buf, 0, (int)byteLen);
            rows.Add(new DatasetImageRow
            {
                Id          = r.GetInt64(0),
                DatasetName = r.GetString(1),
                FileName    = r.IsDBNull(2) ? string.Empty : r.GetString(2),
                ImageData   = buf,
                CreatedAt   = r.IsDBNull(4) ? string.Empty : r.GetString(4)
            });
        }
        return rows;
    }

    // ── Helpers ────────────────────────────────────────────────────────────

    private SqliteConnection OpenConnection()
    {
        EnsureSchemaForCurrentPath();
        var conn = new SqliteConnection($"Data Source={DatabasePath}");
        conn.Open();
        return conn;
    }

    private static void EnsureSchemaForCurrentPath()
    {
        string path = DatabasePath;
        lock (_schemaLock)
        {
            if (!_initializedPaths.Add(path)) return;
        }

        string? dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);

        // Re-use the existing EnsureDatabase logic against this path
        using var conn = new SqliteConnection($"Data Source={path}");
        conn.Open();
        EnsureDatabase(conn);
    }

    private static string GetValue(Dictionary<string, string> data, string key) =>
        data.TryGetValue(key, out string? value) ? value : string.Empty;

    private static int ParseInt(Dictionary<string, string> data, string key)
    {
        string value = GetValue(data, key);
        return int.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out int parsed) ? parsed : 0;
    }

    private static double ParseDouble(Dictionary<string, string> data, string key)
    {
        string value = GetValue(data, key).Replace("%", string.Empty, StringComparison.Ordinal);
        return double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed) ? parsed : 0d;
    }

    private static string FormatPercent(double value) => $"{value:0.0}%";
}
