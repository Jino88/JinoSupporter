using System.Text.Json;
using Microsoft.Data.Sqlite;
using Microsoft.Extensions.Configuration;

namespace JinoSupporter.Web.Services;

public sealed class UserRecord
{
    public long   Id           { get; init; }
    public string Username     { get; init; } = string.Empty;
    public string PasswordHash { get; init; } = string.Empty;
    public string Role         { get; init; } = string.Empty;
    public string CreatedAt    { get; init; } = string.Empty;
}

public sealed record ScheduleItem(
    long         Id,
    string       Title,
    string       Description,
    DateOnly     StartDate,
    DateOnly     EndDate,
    TimeOnly?    StartTime,
    TimeOnly?    EndTime,
    string       Color,
    List<string> Tags,
    string       CreatedAt)
{
    public bool   IsAllDay   => StartTime is null;
    public bool   IsMultiDay => EndDate > StartDate;
    public string TimeDisplay => IsAllDay
        ? "All day"
        : EndTime.HasValue
            ? $"{StartTime!.Value:HH\\:mm}–{EndTime.Value:HH\\:mm}"
            : StartTime!.Value.ToString("HH:mm");
}

public sealed class WebRepository
{
    private static readonly JsonSerializerOptions JsonOpts = new() { PropertyNameCaseInsensitive = true };

    private readonly string _dbPath;
    private readonly string _scheduleDbPath;

    public WebRepository(IConfiguration config)
    {
        // Priority 1: appsettings.json → Database:Path (explicit override)
        // Priority 2: WPF app settings file (DataInference.DatabasePath)
        // Priority 3: default path (%LOCALAPPDATA%\JinoWorkHost\process-review.db)
        string? configured = config["Database:Path"];
        _dbPath = !string.IsNullOrWhiteSpace(configured)
            ? configured
            : WpfSettingsReader.TryGetDatabasePath()
              ?? Path.Combine(
                  Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                  "JinoWorkHost", "process-review.db");

        // Schedule DB: shared with WPF app (schedule.db)
        // Priority 1: appsettings.json → Schedule:Path
        // Priority 2: WPF app settings file (Schedule.DatabasePath)
        // Priority 3: default %LOCALAPPDATA%\JinoWorkHost\schedule.db
        string? schedulePath = config["Schedule:Path"];
        _scheduleDbPath = !string.IsNullOrWhiteSpace(schedulePath)
            ? schedulePath
            : WpfSettingsReader.TryGetScheduleDatabasePath()
              ?? Path.Combine(
                  Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                  "JinoWorkHost", "schedule.db");

        EnsureDatabase();
        EnsureScheduleDatabase();
    }

    public string GetDbPath() => _dbPath;

    // ── Connection ────────────────────────────────────────────────────────────

    private SqliteConnection OpenConnection()
    {
        var conn = new SqliteConnection($"Data Source={_dbPath}");
        conn.Open();
        return conn;
    }

    private SqliteConnection OpenScheduleConnection()
    {
        string? dir = Path.GetDirectoryName(_scheduleDbPath);
        if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);
        var conn = new SqliteConnection($"Data Source={_scheduleDbPath}");
        conn.Open();
        return conn;
    }

    private void EnsureScheduleDatabase()
    {
        using SqliteConnection conn = OpenScheduleConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            CREATE TABLE IF NOT EXISTS Schedules (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                Title       TEXT    NOT NULL DEFAULT '',
                Description TEXT    NOT NULL DEFAULT '',
                StartDate   TEXT    NOT NULL,
                EndDate     TEXT    NOT NULL,
                Color       TEXT    NOT NULL DEFAULT '#4A90D9',
                Tags        TEXT    NOT NULL DEFAULT '[]',
                CreatedAt   TEXT    NOT NULL,
                StartTime   TEXT    NULL,
                EndTime     TEXT    NULL
            );
            CREATE INDEX IF NOT EXISTS idx_sched_start ON Schedules(StartDate);
            CREATE INDEX IF NOT EXISTS idx_sched_end   ON Schedules(EndDate);
            """;
        cmd.ExecuteNonQuery();

        // Migrate: ensure StartTime/EndTime columns exist (legacy DBs may not have them)
        using SqliteCommand check = conn.CreateCommand();
        check.CommandText = "PRAGMA table_info(Schedules);";
        bool hasStartTime = false, hasEndTime = false;
        using (SqliteDataReader r = check.ExecuteReader())
            while (r.Read())
            {
                string col = r.GetString(1);
                if (col.Equals("StartTime", StringComparison.OrdinalIgnoreCase)) hasStartTime = true;
                if (col.Equals("EndTime",   StringComparison.OrdinalIgnoreCase)) hasEndTime   = true;
            }
        if (!hasStartTime) { using SqliteCommand a = conn.CreateCommand(); a.CommandText = "ALTER TABLE Schedules ADD COLUMN StartTime TEXT NULL;"; a.ExecuteNonQuery(); }
        if (!hasEndTime)   { using SqliteCommand a = conn.CreateCommand(); a.CommandText = "ALTER TABLE Schedules ADD COLUMN EndTime TEXT NULL;";   a.ExecuteNonQuery(); }
    }

    private void EnsureDatabase()
    {
        string? dir = Path.GetDirectoryName(_dbPath);
        if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);

        using SqliteConnection conn = OpenConnection();
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

            CREATE TABLE IF NOT EXISTS DatasetTags (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                DatasetName TEXT    NOT NULL UNIQUE,
                Tags        TEXT    NOT NULL DEFAULT '[]',
                CreatedAt   TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_dtag_dataset ON DatasetTags(DatasetName);

            CREATE TABLE IF NOT EXISTS DatasetMemo (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                DatasetName TEXT    NOT NULL UNIQUE,
                Memo        TEXT    NOT NULL DEFAULT '',
                UpdatedAt   TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_dmemo_dataset ON DatasetMemo(DatasetName);

            CREATE TABLE IF NOT EXISTS Reports (
                Id           INTEGER PRIMARY KEY AUTOINCREMENT,
                Title        TEXT    NOT NULL DEFAULT '',
                DatasetNames TEXT    NOT NULL DEFAULT '',
                HtmlContent  TEXT    NOT NULL DEFAULT '',
                CreatedAt    TEXT    NOT NULL
            );

            CREATE TABLE IF NOT EXISTS Users (
                Id           INTEGER PRIMARY KEY AUTOINCREMENT,
                Username     TEXT    NOT NULL UNIQUE COLLATE NOCASE,
                PasswordHash TEXT    NOT NULL,
                Role         TEXT    NOT NULL DEFAULT 'Viewer',
                CreatedAt    TEXT    NOT NULL
            );
            """;
        cmd.ExecuteNonQuery();
        MigrateSchema(conn);
    }

    private static void MigrateSchema(SqliteConnection conn)
    {
        bool hasDatasetNames = false;
        using SqliteCommand check = conn.CreateCommand();
        check.CommandText = "PRAGMA table_info(Reports);";
        using (SqliteDataReader r = check.ExecuteReader())
            while (r.Read())
                if (r.GetString(1).Equals("DatasetNames", StringComparison.OrdinalIgnoreCase))
                    hasDatasetNames = true;

        if (!hasDatasetNames)
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE Reports ADD COLUMN DatasetNames TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }
    }

    // ── Datasets ──────────────────────────────────────────────────────────────

    public List<string> GetAllDatasets()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT DISTINCT DatasetName FROM DataTables ORDER BY DatasetName;";
        var list = new List<string>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
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
        return [.. set.Order(StringComparer.OrdinalIgnoreCase)];
    }

    public List<string> GetDatasetsByTags(IReadOnlyList<string> tags)
    {
        if (tags.Count == 0) return GetAllDatasets();

        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT DatasetName, Tags FROM DatasetTags;";

        var tagSet   = new HashSet<string>(tags, StringComparer.OrdinalIgnoreCase);
        var matchSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            try
            {
                var dsTags = JsonSerializer.Deserialize<List<string>>(r.GetString(1)) ?? [];
                if (dsTags.Any(tagSet.Contains))
                    matchSet.Add(r.GetString(0));
            }
            catch { }
        }
        return [.. matchSet.Order(StringComparer.OrdinalIgnoreCase)];
    }

    // ── Tables ────────────────────────────────────────────────────────────────

    public List<DataTableInfo> GetTablesForDataset(string datasetName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT dt.Id, dt.DatasetName, dt.TableName, dt.Columns, dt.CreatedAt,
                   COUNT(dtr.Id) AS RowCount
            FROM DataTables dt
            LEFT JOIN DataTableRows dtr ON dtr.TableId = dt.Id
            WHERE dt.DatasetName = @d
            GROUP BY dt.Id
            ORDER BY dt.Id;
            """;
        cmd.Parameters.AddWithValue("@d", datasetName);

        var list = new List<DataTableInfo>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            List<ColumnDef> cols = [];
            try { cols = JsonSerializer.Deserialize<List<ColumnDef>>(r.GetString(3), JsonOpts) ?? []; } catch { }

            list.Add(new DataTableInfo
            {
                Id          = r.GetInt64(0),
                DatasetName = r.GetString(1),
                TableName   = r.GetString(2),
                Columns     = cols,
                CreatedAt   = r.GetString(4),
                RowCount    = r.GetInt32(5)
            });
        }
        return list;
    }

    public List<Dictionary<string, string>> GetTableRows(long tableId)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT RowData FROM DataTableRows WHERE TableId = @id ORDER BY Id;";
        cmd.Parameters.AddWithValue("@id", tableId);

        var list = new List<Dictionary<string, string>>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            try
            {
                var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(r.GetString(0), JsonOpts) ?? [];
                list.Add(dict);
            }
            catch { }
        }
        return list;
    }

    public long SaveTable(string datasetName, string tableName,
                          List<ColumnDef> columns, List<Dictionary<string, string>> rows)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();

        using SqliteCommand ins = conn.CreateCommand();
        ins.Transaction = tx;
        ins.CommandText = """
            INSERT INTO DataTables (DatasetName, TableName, Columns, CreatedAt)
            VALUES (@d, @n, @c, @at);
            SELECT last_insert_rowid();
            """;
        ins.Parameters.AddWithValue("@d",  datasetName);
        ins.Parameters.AddWithValue("@n",  tableName);
        ins.Parameters.AddWithValue("@c",  JsonSerializer.Serialize(columns, JsonOpts));
        ins.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
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

    // ── Tags & Memo ───────────────────────────────────────────────────────────

    public void SaveTags(string datasetName, List<string> tags)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO DatasetTags (DatasetName, Tags, CreatedAt)
            VALUES (@d, @t, @at)
            ON CONFLICT(DatasetName) DO UPDATE SET Tags=@t, CreatedAt=@at;
            """;
        cmd.Parameters.AddWithValue("@d",  datasetName);
        cmd.Parameters.AddWithValue("@t",  JsonSerializer.Serialize(tags));
        cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
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
            try { return JsonSerializer.Deserialize<List<string>>(r.GetString(0)) ?? []; } catch { }
        return [];
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
        cmd.Parameters.AddWithValue("@d",  datasetName);
        cmd.Parameters.AddWithValue("@m",  memo);
        cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
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

    public int RenameTag(string oldTag, string newTag)
    {
        if (string.IsNullOrWhiteSpace(oldTag) || string.IsNullOrWhiteSpace(newTag)) return 0;

        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id, Tags FROM DatasetTags;";

        var updates = new List<(long Id, string Json)>();
        using (SqliteDataReader r = cmd.ExecuteReader())
        {
            while (r.Read())
            {
                try
                {
                    var tags = JsonSerializer.Deserialize<List<string>>(r.GetString(1)) ?? [];
                    bool changed = false;
                    for (int i = 0; i < tags.Count; i++)
                    {
                        if (string.Equals(tags[i], oldTag, StringComparison.OrdinalIgnoreCase))
                        { tags[i] = newTag; changed = true; }
                    }
                    if (changed) updates.Add((r.GetInt64(0), JsonSerializer.Serialize(tags)));
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
            upd.Parameters.AddWithValue("@t",  json);
            upd.Parameters.AddWithValue("@id", id);
            upd.ExecuteNonQuery();
        }
        tx.Commit();
        return updates.Count;
    }

    // ── Reports ───────────────────────────────────────────────────────────────

    public List<(long Id, string Title, string DatasetNames, string CreatedAt)> GetReports()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id, Title, DatasetNames, CreatedAt FROM Reports ORDER BY CreatedAt DESC;";
        var list = new List<(long, string, string, string)>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
            list.Add((r.GetInt64(0), r.GetString(1), r.GetString(2), r.GetString(3)));
        return list;
    }

    public string GetReportHtml(long id)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT HtmlContent FROM Reports WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", id);
        using SqliteDataReader r = cmd.ExecuteReader();
        return r.Read() ? r.GetString(0) : string.Empty;
    }

    public long SaveReport(string title, string datasetNames, string htmlContent)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO Reports (Title, DatasetNames, HtmlContent, CreatedAt)
            VALUES (@t, @d, @h, @at);
            SELECT last_insert_rowid();
            """;
        cmd.Parameters.AddWithValue("@t",  title);
        cmd.Parameters.AddWithValue("@d",  datasetNames);
        cmd.Parameters.AddWithValue("@h",  htmlContent);
        cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
        return Convert.ToInt64(cmd.ExecuteScalar() ?? 0);
    }

    public void DeleteReport(long id)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM Reports WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    // ── Dataset management ────────────────────────────────────────────────────

    public List<(string Name, int TableCount)> GetDatasetSummary()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT DatasetName, COUNT(*) AS TableCount
            FROM DataTables
            GROUP BY DatasetName
            ORDER BY DatasetName;
            """;
        var list = new List<(string, int)>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) list.Add((r.GetString(0), r.GetInt32(1)));
        return list;
    }

    public void DeleteTable(long tableId)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();
        using SqliteCommand del1 = conn.CreateCommand();
        del1.Transaction = tx;
        del1.CommandText = "DELETE FROM DataTableRows WHERE TableId=@id;";
        del1.Parameters.AddWithValue("@id", tableId);
        del1.ExecuteNonQuery();

        using SqliteCommand del2 = conn.CreateCommand();
        del2.Transaction = tx;
        del2.CommandText = "DELETE FROM DataTables WHERE Id=@id;";
        del2.Parameters.AddWithValue("@id", tableId);
        del2.ExecuteNonQuery();
        tx.Commit();
    }

    public void RenameTable(long tableId, string newName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE DataTables SET TableName=@n WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@n",  newName);
        cmd.Parameters.AddWithValue("@id", tableId);
        cmd.ExecuteNonQuery();
    }

    public void DeleteTableRow(long tableId, int rowIndex)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id FROM DataTableRows WHERE TableId=@tid ORDER BY Id;";
        cmd.Parameters.AddWithValue("@tid", tableId);
        var ids = new List<long>();
        using (SqliteDataReader r = cmd.ExecuteReader()) while (r.Read()) ids.Add(r.GetInt64(0));
        if (rowIndex < 0 || rowIndex >= ids.Count) return;
        using SqliteCommand del = conn.CreateCommand();
        del.CommandText = "DELETE FROM DataTableRows WHERE Id=@id;";
        del.Parameters.AddWithValue("@id", ids[rowIndex]);
        del.ExecuteNonQuery();
    }

    public void DeleteTableColumn(long tableId, string colField)
    {
        using SqliteConnection conn = OpenConnection();

        // 1. Remove column from schema
        using SqliteCommand getSchema = conn.CreateCommand();
        getSchema.CommandText = "SELECT Columns FROM DataTables WHERE Id=@id;";
        getSchema.Parameters.AddWithValue("@id", tableId);
        string colsJson = (getSchema.ExecuteScalar() as string) ?? "[]";
        var cols = JsonSerializer.Deserialize<List<ColumnDef>>(colsJson, JsonOpts) ?? [];
        cols.RemoveAll(c => c.Field.Equals(colField, StringComparison.OrdinalIgnoreCase));
        using SqliteCommand updSchema = conn.CreateCommand();
        updSchema.CommandText = "UPDATE DataTables SET Columns=@c WHERE Id=@id;";
        updSchema.Parameters.AddWithValue("@c", JsonSerializer.Serialize(cols, JsonOpts));
        updSchema.Parameters.AddWithValue("@id", tableId);
        updSchema.ExecuteNonQuery();

        // 2. Remove field from every row
        using SqliteCommand getRows = conn.CreateCommand();
        getRows.CommandText = "SELECT Id, RowData FROM DataTableRows WHERE TableId=@tid;";
        getRows.Parameters.AddWithValue("@tid", tableId);
        var updates = new List<(long Id, string Json)>();
        using (SqliteDataReader r = getRows.ExecuteReader())
            while (r.Read())
            {
                try
                {
                    var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(r.GetString(1), JsonOpts) ?? [];
                    dict.Remove(colField);
                    updates.Add((r.GetInt64(0), JsonSerializer.Serialize(dict)));
                }
                catch { }
            }
        using SqliteTransaction tx = conn.BeginTransaction();
        foreach ((long id, string json) in updates)
        {
            using SqliteCommand upd = conn.CreateCommand();
            upd.Transaction = tx;
            upd.CommandText = "UPDATE DataTableRows SET RowData=@d WHERE Id=@id;";
            upd.Parameters.AddWithValue("@d", json);
            upd.Parameters.AddWithValue("@id", id);
            upd.ExecuteNonQuery();
        }
        tx.Commit();
    }

    public void UpdateTableRowData(long tableId, int rowIndex, Dictionary<string, string> data)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id FROM DataTableRows WHERE TableId=@tid ORDER BY Id;";
        cmd.Parameters.AddWithValue("@tid", tableId);
        var ids = new List<long>();
        using (SqliteDataReader r = cmd.ExecuteReader()) while (r.Read()) ids.Add(r.GetInt64(0));
        if (rowIndex < 0 || rowIndex >= ids.Count) return;
        using SqliteCommand upd = conn.CreateCommand();
        upd.CommandText = "UPDATE DataTableRows SET RowData=@d WHERE Id=@id;";
        upd.Parameters.AddWithValue("@d", JsonSerializer.Serialize(data));
        upd.Parameters.AddWithValue("@id", ids[rowIndex]);
        upd.ExecuteNonQuery();
    }

    public void DeleteDataset(string datasetName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();

        // Get table IDs first
        using SqliteCommand getIds = conn.CreateCommand();
        getIds.Transaction = tx;
        getIds.CommandText = "SELECT Id FROM DataTables WHERE DatasetName=@d;";
        getIds.Parameters.AddWithValue("@d", datasetName);
        var ids = new List<long>();
        using (SqliteDataReader r = getIds.ExecuteReader())
            while (r.Read()) ids.Add(r.GetInt64(0));

        foreach (long id in ids)
        {
            using SqliteCommand dr = conn.CreateCommand();
            dr.Transaction = tx;
            dr.CommandText = "DELETE FROM DataTableRows WHERE TableId=@id;";
            dr.Parameters.AddWithValue("@id", id);
            dr.ExecuteNonQuery();
        }

        using SqliteCommand dt = conn.CreateCommand();
        dt.Transaction = tx;
        dt.CommandText = "DELETE FROM DataTables WHERE DatasetName=@d;";
        dt.Parameters.AddWithValue("@d", datasetName);
        dt.ExecuteNonQuery();

        using SqliteCommand dtag = conn.CreateCommand();
        dtag.Transaction = tx;
        dtag.CommandText = "DELETE FROM DatasetTags WHERE DatasetName=@d;";
        dtag.Parameters.AddWithValue("@d", datasetName);
        dtag.ExecuteNonQuery();

        using SqliteCommand dm = conn.CreateCommand();
        dm.Transaction = tx;
        dm.CommandText = "DELETE FROM DatasetMemo WHERE DatasetName=@d;";
        dm.Parameters.AddWithValue("@d", datasetName);
        dm.ExecuteNonQuery();

        tx.Commit();
    }

    public int GetTotalRowCount()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COUNT(*) FROM DataTableRows;";
        return Convert.ToInt32(cmd.ExecuteScalar() ?? 0);
    }

    // ── Users ─────────────────────────────────────────────────────────────────

    public UserRecord? GetUser(string username)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id, Username, PasswordHash, Role, CreatedAt FROM Users WHERE Username=@u COLLATE NOCASE LIMIT 1;";
        cmd.Parameters.AddWithValue("@u", username);
        using SqliteDataReader r = cmd.ExecuteReader();
        if (!r.Read()) return null;
        return new UserRecord
        {
            Id           = r.GetInt64(0),
            Username     = r.GetString(1),
            PasswordHash = r.GetString(2),
            Role         = r.GetString(3),
            CreatedAt    = r.GetString(4)
        };
    }

    public List<UserRecord> GetAllUsers()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id, Username, PasswordHash, Role, CreatedAt FROM Users ORDER BY Id;";
        var list = new List<UserRecord>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new UserRecord
            {
                Id           = r.GetInt64(0),
                Username     = r.GetString(1),
                PasswordHash = r.GetString(2),
                Role         = r.GetString(3),
                CreatedAt    = r.GetString(4)
            });
        return list;
    }

    public void AddUser(string username, string passwordHash, string role)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "INSERT INTO Users (Username, PasswordHash, Role, CreatedAt) VALUES (@u, @p, @r, @t);";
        cmd.Parameters.AddWithValue("@u", username);
        cmd.Parameters.AddWithValue("@p", passwordHash);
        cmd.Parameters.AddWithValue("@r", role);
        cmd.Parameters.AddWithValue("@t", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        cmd.ExecuteNonQuery();
    }

    public void DeleteUser(long id)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM Users WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    public void UpdateUserRole(long id, string role)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE Users SET Role=@r WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@r", role);
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    public void UpdateUserPassword(long id, string passwordHash)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE Users SET PasswordHash=@p WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@p", passwordHash);
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    // Returns (FileName, MediaType, Base64) for all images attached to a dataset.
    public List<(string FileName, string MediaType, string Base64)> GetImages(string datasetName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT FileName, ImageData FROM DatasetImages WHERE DatasetName=@n ORDER BY Id;";
        cmd.Parameters.AddWithValue("@n", datasetName);

        var result = new List<(string, string, string)>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            string fileName = r.IsDBNull(0) ? "image" : r.GetString(0);
            long   byteLen  = r.GetBytes(1, 0, null, 0, 0);
            byte[] buf      = new byte[byteLen];
            r.GetBytes(1, 0, buf, 0, (int)byteLen);

            // Infer media type from file extension; default to image/png
            string ext = Path.GetExtension(fileName).TrimStart('.').ToLowerInvariant();
            string mediaType = ext switch
            {
                "jpg" or "jpeg" => "image/jpeg",
                "gif"           => "image/gif",
                "webp"          => "image/webp",
                _               => "image/png"
            };
            result.Add((fileName, mediaType, Convert.ToBase64String(buf)));
        }
        return result;
    }

    // ── Schedule ──────────────────────────────────────────────────────────────

    private static ScheduleItem ReadScheduleItem(SqliteDataReader r)
    {
        string? stStr = r.IsDBNull(5) ? null : r.GetString(5);
        string? etStr = r.IsDBNull(6) ? null : r.GetString(6);
        TimeOnly? st = stStr is not null && TimeOnly.TryParse(stStr, out var tp) ? tp : null;
        TimeOnly? et = etStr is not null && TimeOnly.TryParse(etStr, out var tp2) ? tp2 : null;
        List<string> tags;
        try { tags = JsonSerializer.Deserialize<List<string>>(r.GetString(8), JsonOpts) ?? []; }
        catch { tags = []; }
        return new ScheduleItem(
            r.GetInt64(0),
            r.GetString(1),
            r.GetString(2),
            DateOnly.Parse(r.GetString(3)),
            DateOnly.Parse(r.GetString(4)),
            st, et,
            r.GetString(7),
            tags,
            r.GetString(9));
    }

    public List<ScheduleItem> GetSchedulesInRange(DateOnly from, DateOnly to, List<string>? filterTags = null)
    {
        using SqliteConnection conn = OpenScheduleConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Title, Description, StartDate, EndDate, StartTime, EndTime, Color, Tags, CreatedAt
            FROM Schedules
            WHERE EndDate >= @from AND StartDate <= @to
            ORDER BY StartDate, StartTime, EndDate;
            """;
        cmd.Parameters.AddWithValue("@from", from.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@to",   to.ToString("yyyy-MM-dd"));

        var items = new List<ScheduleItem>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) items.Add(ReadScheduleItem(r));

        if (filterTags is { Count: > 0 })
            items = items.Where(s => s.Tags.Any(t =>
                filterTags.Contains(t, StringComparer.OrdinalIgnoreCase))).ToList();
        return items;
    }

    public List<string> GetAllScheduleTags()
    {
        using SqliteConnection conn = OpenScheduleConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Tags FROM Schedules;";
        var set = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            try
            {
                var tags = JsonSerializer.Deserialize<List<string>>(r.GetString(0), JsonOpts) ?? [];
                foreach (string t in tags) set.Add(t.Trim());
            }
            catch { }
        }
        return [.. set];
    }

    public long AddSchedule(string title, string description,
        DateOnly start, DateOnly end, string color, List<string> tags,
        TimeOnly? startTime = null, TimeOnly? endTime = null)
    {
        using SqliteConnection conn = OpenScheduleConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO Schedules (Title, Description, StartDate, EndDate, StartTime, EndTime, Color, Tags, CreatedAt)
            VALUES (@t, @d, @sd, @ed, @st, @et, @c, @tg, @at);
            SELECT last_insert_rowid();
            """;
        cmd.Parameters.AddWithValue("@t",  title);
        cmd.Parameters.AddWithValue("@d",  description);
        cmd.Parameters.AddWithValue("@sd", start.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@ed", end.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@st", startTime.HasValue ? (object)startTime.Value.ToString("HH:mm") : DBNull.Value);
        cmd.Parameters.AddWithValue("@et", endTime.HasValue   ? (object)endTime.Value.ToString("HH:mm")   : DBNull.Value);
        cmd.Parameters.AddWithValue("@c",  color);
        cmd.Parameters.AddWithValue("@tg", JsonSerializer.Serialize(tags));
        cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("o"));
        return Convert.ToInt64(cmd.ExecuteScalar());
    }

    public void UpdateSchedule(long id, string title, string description,
        DateOnly start, DateOnly end, string color, List<string> tags,
        TimeOnly? startTime = null, TimeOnly? endTime = null)
    {
        using SqliteConnection conn = OpenScheduleConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE Schedules SET Title=@t, Description=@d, StartDate=@sd, EndDate=@ed,
                StartTime=@st, EndTime=@et, Color=@c, Tags=@tg
            WHERE Id=@id;
            """;
        cmd.Parameters.AddWithValue("@t",  title);
        cmd.Parameters.AddWithValue("@d",  description);
        cmd.Parameters.AddWithValue("@sd", start.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@ed", end.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@st", startTime.HasValue ? (object)startTime.Value.ToString("HH:mm") : DBNull.Value);
        cmd.Parameters.AddWithValue("@et", endTime.HasValue   ? (object)endTime.Value.ToString("HH:mm")   : DBNull.Value);
        cmd.Parameters.AddWithValue("@c",  color);
        cmd.Parameters.AddWithValue("@tg", JsonSerializer.Serialize(tags));
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    public void DeleteSchedule(long id)
    {
        using SqliteConnection conn = OpenScheduleConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM Schedules WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }
}
