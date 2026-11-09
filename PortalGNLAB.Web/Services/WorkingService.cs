using System.Globalization;
using System.Text.Json;
using Microsoft.Data.Sqlite;

namespace PortalGNLAB.Web.Services;

public sealed class WorkingConflictException(string message) : Exception(message);

public sealed record WorkingGroup(string Id, string Name, string? ParentId, List<string> Models, int Sort, int Position);

public sealed record WorkingGroups(List<WorkingGroup> Groups, string Source, string? Error);

public sealed record WorkingAddedModel(string Name, string GroupId, string CreatedBy, string CreatedAt);

public sealed record WorkingTree(List<string> Hidden, List<WorkingAddedModel> Added);

public sealed record WorkingModel(string Name, JsonElement Info, int Revision, string UpdatedAt, string UpdatedBy);

public sealed record WorkingModelSave(JsonElement Info, int Revision);

/// <summary>
/// An entry. Issue fields (Seq … FinalResult) are "" / 0 for other kinds; Logs are the dated progress notes, oldest first;
/// Images are the entry's own pictures (each note carries its own). SetupRowId is the Setup Plan row an issue was raised
/// against (it is one of that row's Issue (Problem) list), or "".
/// </summary>
public sealed record WorkingEntry(
    string Id, string ModelName, string Kind, string Date, string Title, string Body, string Status, string Severity,
    string Result, string Owner, string Due, int Revision, string CreatedBy, string CreatedAt, string UpdatedBy, string UpdatedAt,
    int Seq, string Phase, string Category, string Process, string Cause, string DoneDate, string FinalResult, List<WorkingEntryLog> Logs,
    List<WorkingImage> Images, string SetupRowId = "", string ParentId = "");

public sealed record WorkingEntrySave(
    string? Kind, string? Date, string? Title, string? Body, string? Status, string? Severity,
    string? Result, string? Owner, string? Due, int Revision,
    string? Phase = null, string? Category = null, string? Process = null, string? Cause = null, string? DoneDate = null, string? FinalResult = null,
    string? SetupRowId = null, string? ParentId = null);

/// <summary>
/// A process row of a model's Setup Plan: machine line plan / state / note and setup plan / finish / status.
/// States are "", "O" (OK), "X" (not yet set up) or "Δ" (setting up now); dates are "" or yyyy-MM-dd.
/// Its problems are the issues whose SetupRowId is this row.
/// </summary>
public sealed record WorkingSetupRow(
    string Id, string ModelName, int No, string LayoutNo, string Line, string ProcessName,
    string MachinePlan, string MachineState, string MachineNote, string SetupPlan, string SetupFinish, string SetupStatus,
    int Revision, string CreatedBy, string CreatedAt, string UpdatedBy, string UpdatedAt);

/// <summary>No 0 / null on create = after the last row; on update = keep.</summary>
public sealed record WorkingSetupRowSave(
    int? No, string? LayoutNo, string? Line, string? ProcessName, string? MachinePlan, string? MachineState, string? MachineNote,
    string? SetupPlan, string? SetupFinish, string? SetupStatus, int Revision = 0);

public sealed record WorkingSetupImport(List<WorkingSetupRowSave>? Rows);

/// <summary>A dated progress note; Due / DoneDate are "" or yyyy-MM-dd (a note with a done date is done).</summary>
public sealed record WorkingEntryLog(
    string Id, string EntryId, string Date, string Text, int Revision, string CreatedBy, string CreatedAt, string UpdatedBy, string UpdatedAt,
    string Owner, string Due, string DoneDate, List<WorkingImage> Images);

public sealed record WorkingLogSave(string? Date, string? Text, int Revision = 0, string? Owner = null, string? Due = null, string? DoneDate = null);

/// <summary>
/// A picture or video pasted into an entry (LogId "") or into one of its notes. The original bytes are stored in the DB
/// unchanged and served by OpenImage.
/// </summary>
public sealed record WorkingImage(string Id, string EntryId, string LogId, string FileName, string ContentType, long Size, string CreatedBy, string CreatedAt);

/// <summary>An uploaded file: its name, length and a way to read it (opened once to tell the format and once to store it).</summary>
public sealed record WorkingUpload(string Name, long Length, Func<Stream> Open)
{
    public static WorkingUpload Of(string name, byte[] data) => new(name, data.LongLength, () => new MemoryStream(data, false));
}

/// <summary>A stored file opened for reading; disposing it closes its DB connection.</summary>
public sealed class WorkingImageFile(string fileName, string contentType, long size, Stream data) : IDisposable
{
    public string FileName { get; } = fileName;
    public string ContentType { get; } = contentType;
    public long Size { get; } = size;
    public Stream Data { get; } = data;
    public void Dispose() => Data.Dispose();
}

/// <summary>
/// Per-model work history. Model groups are read (never written) from the CompanyGlossary DB; model info
/// and dev / issue / review entries live in PortalGNLAB.db.
/// </summary>
public sealed class WorkingService(string? dbPath = null, string? glossaryDbPath = null)
{
    public const string SchemaSql = """
        CREATE TABLE IF NOT EXISTS WorkingModels (
            ModelName TEXT PRIMARY KEY COLLATE NOCASE,
            InfoJson TEXT NOT NULL DEFAULT '{}',
            Revision INTEGER NOT NULL,
            UpdatedAt TEXT NOT NULL,
            UpdatedBy TEXT NOT NULL DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS WorkingEntries (
            Id TEXT PRIMARY KEY,
            ModelName TEXT NOT NULL COLLATE NOCASE,
            Kind TEXT NOT NULL,
            Date TEXT NOT NULL,
            Title TEXT NOT NULL,
            Body TEXT NOT NULL DEFAULT '',
            Status TEXT NOT NULL DEFAULT '',
            Severity TEXT NOT NULL DEFAULT '',
            Result TEXT NOT NULL DEFAULT '',
            Owner TEXT NOT NULL DEFAULT '',
            Due TEXT NOT NULL DEFAULT '',
            Revision INTEGER NOT NULL,
            CreatedBy TEXT NOT NULL DEFAULT '',
            CreatedAt TEXT NOT NULL,
            UpdatedBy TEXT NOT NULL DEFAULT '',
            UpdatedAt TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS IX_WorkingEntries_ModelName ON WorkingEntries (ModelName);
        CREATE TABLE IF NOT EXISTS WorkingEntryLogs (
            Id TEXT PRIMARY KEY,
            EntryId TEXT NOT NULL,
            Date TEXT NOT NULL,
            Text TEXT NOT NULL,
            Revision INTEGER NOT NULL,
            CreatedBy TEXT NOT NULL DEFAULT '',
            CreatedAt TEXT NOT NULL,
            UpdatedBy TEXT NOT NULL DEFAULT '',
            UpdatedAt TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS IX_WorkingEntryLogs_EntryId ON WorkingEntryLogs (EntryId);
        CREATE TABLE IF NOT EXISTS WorkingImages (
            Id TEXT PRIMARY KEY,
            EntryId TEXT NOT NULL,
            LogId TEXT NOT NULL DEFAULT '',
            FileName TEXT NOT NULL,
            ContentType TEXT NOT NULL,
            Size INTEGER NOT NULL,
            Data BLOB NOT NULL,
            CreatedBy TEXT NOT NULL DEFAULT '',
            CreatedAt TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS IX_WorkingImages_EntryId ON WorkingImages (EntryId);
        CREATE TABLE IF NOT EXISTS WorkingSetupRows (
            Id TEXT PRIMARY KEY,
            ModelName TEXT NOT NULL COLLATE NOCASE,
            No INTEGER NOT NULL,
            LayoutNo TEXT NOT NULL DEFAULT '',
            Line TEXT NOT NULL DEFAULT '',
            ProcessName TEXT NOT NULL,
            MachinePlan TEXT NOT NULL DEFAULT '',
            MachineState TEXT NOT NULL DEFAULT '',
            MachineNote TEXT NOT NULL DEFAULT '',
            SetupPlan TEXT NOT NULL DEFAULT '',
            SetupFinish TEXT NOT NULL DEFAULT '',
            SetupStatus TEXT NOT NULL DEFAULT '',
            Revision INTEGER NOT NULL,
            CreatedBy TEXT NOT NULL DEFAULT '',
            CreatedAt TEXT NOT NULL,
            UpdatedBy TEXT NOT NULL DEFAULT '',
            UpdatedAt TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS IX_WorkingSetupRows_ModelName ON WorkingSetupRows (ModelName);
        CREATE TABLE IF NOT EXISTS WorkingEntrySeq (
            ModelName TEXT NOT NULL COLLATE NOCASE,
            Kind TEXT NOT NULL,
            Last INTEGER NOT NULL,
            PRIMARY KEY (ModelName, Kind)
        );
        CREATE TABLE IF NOT EXISTS WorkingTreeModels (
            Name TEXT PRIMARY KEY COLLATE NOCASE,
            GroupId TEXT NOT NULL DEFAULT '',
            Added INTEGER NOT NULL DEFAULT 0,
            Hidden INTEGER NOT NULL DEFAULT 0,
            CreatedBy TEXT NOT NULL DEFAULT '',
            CreatedAt TEXT NOT NULL DEFAULT '',
            UpdatedBy TEXT NOT NULL DEFAULT '',
            UpdatedAt TEXT NOT NULL DEFAULT ''
        );
        """;

    /// <summary>Issue columns added 2026-11-09; added in place to tables created before.</summary>
    private static readonly (string Name, string Type)[] AddedEntryColumns =
    [
        ("Seq", "INTEGER NOT NULL DEFAULT 0"),
        ("Phase", "TEXT NOT NULL DEFAULT ''"),
        ("Category", "TEXT NOT NULL DEFAULT ''"),
        ("Process", "TEXT NOT NULL DEFAULT ''"),
        ("Cause", "TEXT NOT NULL DEFAULT ''"),
        ("DoneDate", "TEXT NOT NULL DEFAULT ''"),
        ("FinalResult", "TEXT NOT NULL DEFAULT ''"),
        ("SetupRowId", "TEXT NOT NULL DEFAULT ''"),
        ("ParentId", "TEXT NOT NULL DEFAULT ''"),
    ];

    /// <summary>Note columns added 2026-11-09 (schedule and done date of a note).</summary>
    private static readonly (string Name, string Type)[] AddedLogColumns =
    [
        ("Owner", "TEXT NOT NULL DEFAULT ''"),
        ("Due", "TEXT NOT NULL DEFAULT ''"),
        ("DoneDate", "TEXT NOT NULL DEFAULT ''"),
    ];

    /// <summary>Picture and video formats kept, told apart by their first bytes (the sent content type is not trusted).</summary>
    public static readonly string[] ImageTypes = ["image/png", "image/jpeg", "image/gif", "image/webp", "image/bmp"];
    public static readonly string[] VideoTypes = ["video/mp4", "video/quicktime", "video/webm", "video/x-matroska", "video/x-msvideo"];
    public const long MaxImageBytes = 10L * 1024 * 1024;
    public const long MaxVideoBytes = 500L * 1024 * 1024;
    /// <summary>All files of one upload together.</summary>
    public const long MaxUploadBytes = 1024L * 1024 * 1024;
    public const int MaxImagesPerUpload = 20;
    /// <summary>Bytes read from the start of a file to tell its format.</summary>
    public const int SniffBytes = 64;

    public const string DefaultGlossaryDb = @"E:\db\CompanyGlossary\company-glossary.db";
    public const string GlossaryEnvVar = "PORTALGNLAB_GLOSSARY_DB";
    public static readonly string[] Kinds = ["dev", "issue", "review"];

    private const string EntryColumns =
        "Id, ModelName, Kind, Date, Title, Body, Status, Severity, Result, Owner, Due, Revision, CreatedBy, CreatedAt, UpdatedBy, UpdatedAt, " +
        "Seq, Phase, Category, Process, Cause, DoneDate, FinalResult, SetupRowId, ParentId";

    private const string LogColumns = "Id, EntryId, Date, Text, Revision, CreatedBy, CreatedAt, UpdatedBy, UpdatedAt, Owner, Due, DoneDate";

    private const string ImageColumns = "i.Id, i.EntryId, i.LogId, i.FileName, i.ContentType, i.Size, i.CreatedBy, i.CreatedAt";

    private static readonly object SchemaGate = new();
    private static readonly HashSet<string> SchemaReady = [];

    private SqliteConnection Open()
    {
        SqliteConnection conn = dbPath is null ? PortalDb.Open() : PortalDb.Open(dbPath);
        string key = conn.DataSource;
        lock (SchemaGate)
        {
            if (!SchemaReady.Contains(key))
            {
                using SqliteCommand cmd = conn.CreateCommand();
                cmd.CommandText = SchemaSql;
                cmd.ExecuteNonQuery();
                MigrateEntries(conn);
                SchemaReady.Add(key);
            }
        }
        return conn;
    }

    /// <summary>
    /// Adds the columns missing from older WorkingEntries / WorkingEntryLogs tables and numbers entries that have no Seq yet.
    /// </summary>
    private static void MigrateEntries(SqliteConnection conn)
    {
        using SqliteTransaction tx = conn.BeginTransaction();
        AddMissingColumns(conn, tx, "WorkingEntries", AddedEntryColumns);
        AddMissingColumns(conn, tx, "WorkingEntryLogs", AddedLogColumns);
        var unnumbered = new List<(string Id, string Model, string Kind)>();
        using (SqliteCommand find = Cmd(conn, "SELECT Id, ModelName, Kind FROM WorkingEntries WHERE Seq<=0 ORDER BY CreatedAt, Id", tx))
        using (SqliteDataReader r = find.ExecuteReader())
            while (r.Read()) unnumbered.Add((r.GetString(0), r.GetString(1), r.GetString(2)));
        foreach ((string id, string model, string kind) in unnumbered)
            using (SqliteCommand set = Cmd(conn, "UPDATE WorkingEntries SET Seq=@s WHERE Id=@id", tx,
                ("@s", NextSeq(conn, tx, model, kind)), ("@id", id)))
                set.ExecuteNonQuery();
        tx.Commit();
    }

    private static void AddMissingColumns(SqliteConnection conn, SqliteTransaction tx, string table, (string Name, string Type)[] columns)
    {
        var have = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using (SqliteCommand info = Cmd(conn, $"PRAGMA table_info({table})", tx))
        using (SqliteDataReader r = info.ExecuteReader())
            while (r.Read()) have.Add(r.GetString(1));
        foreach ((string name, string type) in columns.Where(c => !have.Contains(c.Name)))
            using (SqliteCommand alter = Cmd(conn, $"ALTER TABLE {table} ADD COLUMN {name} {type}", tx))
                alter.ExecuteNonQuery();
    }

    /// <summary>
    /// Takes the next number of a model + kind. The last number given is kept in WorkingEntrySeq, so a number is
    /// never reused even when the newest entry is deleted.
    /// </summary>
    private static int NextSeq(SqliteConnection conn, SqliteTransaction tx, string model, string kind)
    {
        using SqliteCommand cmd = Cmd(conn,
            "SELECT MAX(COALESCE((SELECT Last FROM WorkingEntrySeq WHERE ModelName=@m AND Kind=@k), 0), " +
            "COALESCE((SELECT MAX(Seq) FROM WorkingEntries WHERE ModelName=@m AND Kind=@k), 0)) + 1", tx, ("@m", model), ("@k", kind));
        int next = Convert.ToInt32(cmd.ExecuteScalar(), CultureInfo.InvariantCulture);
        using SqliteCommand keep = Cmd(conn,
            "INSERT INTO WorkingEntrySeq (ModelName, Kind, Last) VALUES (@m,@k,@n) ON CONFLICT(ModelName, Kind) DO UPDATE SET Last=@n", tx,
            ("@m", model), ("@k", kind), ("@n", next));
        keep.ExecuteNonQuery();
        return next;
    }

    private static string Now() => DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);

    private static SqliteCommand Cmd(SqliteConnection conn, string sql, SqliteTransaction? tx = null, params (string, object?)[] args)
    {
        SqliteCommand cmd = conn.CreateCommand();
        cmd.Transaction = tx;
        cmd.CommandText = sql;
        foreach ((string name, object? value) in args) cmd.Parameters.AddWithValue(name, value ?? DBNull.Value);
        return cmd;
    }

    private static string CleanModelName(string? name)
    {
        string n = (name ?? "").Trim();
        if (n.Length == 0) throw new ArgumentException("Model name is required.");
        if (n.Length > 80) throw new ArgumentException("Model name must be 80 characters or fewer.");
        return n;
    }

    private static string Limit(string? value, string label, int max)
    {
        string v = (value ?? "").Trim();
        if (v.Length > max) throw new ArgumentException($"{label} must be {max} characters or fewer.");
        return v;
    }

    // ── Model groups (read-only) ────────────────────────────────────────────

    public string GlossaryPath => glossaryDbPath
        ?? (Environment.GetEnvironmentVariable(GlossaryEnvVar) is { Length: > 0 } env ? env : DefaultGlossaryDb);

    public WorkingGroups ListGroups()
    {
        string path = GlossaryPath;
        try
        {
            if (!File.Exists(path)) return new([], path, "Model group database was not found.");
            using SqliteConnection conn = PortalDb.Open(path, readOnly: true);
            using SqliteCommand cmd = Cmd(conn,
                "SELECT id, name, parent_id, models, sort, position FROM model_groups ORDER BY sort, position, name");
            var groups = new List<WorkingGroup>();
            using SqliteDataReader r = cmd.ExecuteReader();
            while (r.Read())
            {
                groups.Add(new(r.GetString(0), r.IsDBNull(1) ? "" : r.GetString(1),
                    r.IsDBNull(2) || r.GetString(2).Length == 0 ? null : r.GetString(2),
                    ParseModels(r.IsDBNull(3) ? null : r.GetString(3)),
                    r.IsDBNull(4) ? 0 : r.GetInt32(4), r.IsDBNull(5) ? 0 : r.GetInt32(5)));
            }
            return new(groups, path, null);
        }
        catch (Exception ex)
        {
            return new([], path, $"Could not read model groups: {ex.Message}");
        }
    }

    private static List<string> ParseModels(string? json)
    {
        var list = new List<string>();
        if (string.IsNullOrWhiteSpace(json)) return list;
        try
        {
            using JsonDocument doc = JsonDocument.Parse(json);
            if (doc.RootElement.ValueKind == JsonValueKind.Array)
                foreach (JsonElement e in doc.RootElement.EnumerateArray())
                    if (e.ValueKind == JsonValueKind.String && e.GetString() is { Length: > 0 } s) list.Add(s);
        }
        catch (JsonException) { }
        return list;
    }

    // ── Model tree settings (hidden models, added models) ───────────────────
    // The Glossary DB is never written; what the tree hides or adds lives in PortalGNLAB.db.

    public const int MaxHiddenModels = 5000;
    public const int MaxTreeModelName = 100;

    private static string TreeName(string? name, string label)
    {
        string n = (name ?? "").Trim();
        if (n.Length == 0) throw new ArgumentException($"{label} is required.");
        if (n.Length > MaxTreeModelName) throw new ArgumentException($"{label} must be {MaxTreeModelName} characters or fewer.");
        return n;
    }

    public WorkingTree GetTree()
    {
        using SqliteConnection conn = Open();
        return ReadTree(conn, null);
    }

    private static WorkingTree ReadTree(SqliteConnection conn, SqliteTransaction? tx)
    {
        var hidden = new List<string>();
        using (SqliteCommand cmd = Cmd(conn, "SELECT Name FROM WorkingTreeModels WHERE Hidden=1 ORDER BY CreatedAt, Name", tx))
        using (SqliteDataReader r = cmd.ExecuteReader())
            while (r.Read()) hidden.Add(r.GetString(0));
        var added = new List<WorkingAddedModel>();
        using (SqliteCommand cmd = Cmd(conn, "SELECT Name, GroupId, CreatedBy, CreatedAt FROM WorkingTreeModels WHERE Added=1 ORDER BY CreatedAt, Name", tx))
        using (SqliteDataReader r = cmd.ExecuteReader())
            while (r.Read()) added.Add(new(r.GetString(0), r.GetString(1), r.GetString(2), r.GetString(3)));
        return new(hidden, added);
    }

    /// <summary>Replaces the whole hidden list; names are trimmed, blanks dropped, duplicates (any case) merged.</summary>
    public WorkingTree SetHidden(IEnumerable<string?>? names, string user)
    {
        var clean = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string? raw in names ?? [])
        {
            string n = (raw ?? "").Trim();
            if (n.Length == 0) continue;
            if (n.Length > MaxTreeModelName) throw new ArgumentException($"A model name must be {MaxTreeModelName} characters or fewer.");
            if (seen.Add(n)) clean.Add(n);
            if (clean.Count > MaxHiddenModels) throw new ArgumentException($"At most {MaxHiddenModels} models can be hidden.");
        }
        string now = Now();
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        using (SqliteCommand clear = Cmd(conn, "UPDATE WorkingTreeModels SET Hidden=0, UpdatedBy=@u, UpdatedAt=@t WHERE Hidden=1", tx, ("@u", user), ("@t", now)))
            clear.ExecuteNonQuery();
        foreach (string n in clean)
            using (SqliteCommand set = Cmd(conn,
                "INSERT INTO WorkingTreeModels (Name, Hidden, CreatedBy, CreatedAt, UpdatedBy, UpdatedAt) VALUES (@n,1,@u,@t,@u,@t) " +
                "ON CONFLICT(Name) DO UPDATE SET Hidden=1, UpdatedBy=@u, UpdatedAt=@t", tx, ("@n", n), ("@u", user), ("@t", now)))
                set.ExecuteNonQuery();
        using (SqliteCommand drop = Cmd(conn, "DELETE FROM WorkingTreeModels WHERE Added=0 AND Hidden=0", tx))
            drop.ExecuteNonQuery();
        WorkingTree tree = ReadTree(conn, tx);
        tx.Commit();
        return tree;
    }

    /// <summary>Adds a model to the tree: under a Glossary group, or in the "Added" bundle when groupId is empty.</summary>
    public WorkingAddedModel AddTreeModel(string? name, string? groupId, string user)
    {
        string n = TreeName(name, "Model name");
        string gid = (groupId ?? "").Trim();
        WorkingGroups groups = ListGroups();
        if (gid.Length > 0)
        {
            if (groups.Error is not null) throw new ArgumentException("Model groups could not be read, so the model can only be added without a group.");
            if (!groups.Groups.Any(g => g.Id == gid)) throw new ArgumentException("The model group was not found.");
        }
        if (groups.Groups.Any(g => g.Models.Any(m => string.Equals(m.Trim(), n, StringComparison.OrdinalIgnoreCase))))
            throw new WorkingConflictException("That model is already in a model group.");
        string now = Now();
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        using (SqliteCommand exists = Cmd(conn, "SELECT 1 FROM WorkingTreeModels WHERE Name=@n AND Added=1", tx, ("@n", n)))
            if (exists.ExecuteScalar() is not null) throw new WorkingConflictException("That model was already added.");
        // a name that was only hidden has a row already; it becomes an added model and keeps its hidden flag
        using (SqliteCommand ins = Cmd(conn,
            "INSERT INTO WorkingTreeModels (Name, GroupId, Added, CreatedBy, CreatedAt, UpdatedBy, UpdatedAt) VALUES (@n,@g,1,@u,@t,@u,@t) " +
            "ON CONFLICT(Name) DO UPDATE SET GroupId=@g, Added=1, CreatedBy=@u, CreatedAt=@t, UpdatedBy=@u, UpdatedAt=@t", tx,
            ("@n", n), ("@g", gid), ("@u", user), ("@t", now)))
            ins.ExecuteNonQuery();
        WorkingAddedModel added = ReadTree(conn, tx).Added.First(a => string.Equals(a.Name, n, StringComparison.OrdinalIgnoreCase));
        tx.Commit();
        return added;
    }

    /// <summary>Removes an added model (and its hidden flag); false when no such added model exists.</summary>
    public bool DeleteTreeModel(string? name)
    {
        string n = (name ?? "").Trim();
        using SqliteConnection conn = Open();
        using SqliteCommand del = Cmd(conn, "DELETE FROM WorkingTreeModels WHERE Name=@n AND Added=1", null, ("@n", n));
        return del.ExecuteNonQuery() > 0;
    }

    // ── Model info ──────────────────────────────────────────────────────────

    private static JsonElement ParseInfo(string text)
    {
        using JsonDocument doc = JsonDocument.Parse(string.IsNullOrWhiteSpace(text) ? "{}" : text);
        return doc.RootElement.Clone();
    }

    private static WorkingModel? ReadModel(SqliteConnection conn, string name, SqliteTransaction? tx = null)
    {
        using SqliteCommand cmd = Cmd(conn,
            "SELECT ModelName, InfoJson, Revision, UpdatedAt, UpdatedBy FROM WorkingModels WHERE ModelName=@n", tx, ("@n", name));
        using SqliteDataReader r = cmd.ExecuteReader();
        return r.Read() ? new(r.GetString(0), ParseInfo(r.GetString(1)), r.GetInt32(2), r.GetString(3), r.GetString(4)) : null;
    }

    /// <summary>A model never saved yet comes back with empty info and revision 0.</summary>
    public WorkingModel GetModel(string name)
    {
        name = CleanModelName(name);
        using SqliteConnection conn = Open();
        return ReadModel(conn, name) ?? new(name, ParseInfo("{}"), 0, "", "");
    }

    public WorkingModel SaveModel(string name, WorkingModelSave body, string user)
    {
        name = CleanModelName(name);
        if (body.Info.ValueKind != JsonValueKind.Object) throw new ArgumentException("Info must be a JSON object.");
        string info = body.Info.GetRawText();
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        WorkingModel? current = ReadModel(conn, name, tx);
        if ((current?.Revision ?? 0) != body.Revision)
            throw new WorkingConflictException("Someone else saved this model first. Reload and try again.");
        if (current is null)
        {
            using SqliteCommand ins = Cmd(conn,
                "INSERT INTO WorkingModels (ModelName, InfoJson, Revision, UpdatedAt, UpdatedBy) VALUES (@n,@i,1,@t,@u)", tx,
                ("@n", name), ("@i", info), ("@t", Now()), ("@u", user));
            ins.ExecuteNonQuery();
        }
        else
        {
            using SqliteCommand upd = Cmd(conn,
                "UPDATE WorkingModels SET InfoJson=@i, Revision=Revision+1, UpdatedAt=@t, UpdatedBy=@u WHERE ModelName=@n", tx,
                ("@n", name), ("@i", info), ("@t", Now()), ("@u", user));
            upd.ExecuteNonQuery();
        }
        WorkingModel saved = ReadModel(conn, name, tx)!;
        tx.Commit();
        return saved;
    }

    // ── Entries ─────────────────────────────────────────────────────────────

    private static WorkingEntry ReadEntry(SqliteDataReader r, List<WorkingEntryLog> logs, List<WorkingImage> images) =>
        new(r.GetString(0), r.GetString(1), r.GetString(2), r.GetString(3), r.GetString(4), r.GetString(5), r.GetString(6),
            r.GetString(7), r.GetString(8), r.GetString(9), r.GetString(10), r.GetInt32(11), r.GetString(12), r.GetString(13),
            r.GetString(14), r.GetString(15), r.GetInt32(16), r.GetString(17), r.GetString(18), r.GetString(19), r.GetString(20),
            r.GetString(21), r.GetString(22), logs, images, r.GetString(23), r.GetString(24));

    /// <summary>A note row; its pictures are picked from <paramref name="images"/> (pictures of its entry or wider).</summary>
    private static WorkingEntryLog ReadLog(SqliteDataReader r, List<WorkingImage> images) =>
        new(r.GetString(0), r.GetString(1), r.GetString(2), r.GetString(3), r.GetInt32(4), r.GetString(5), r.GetString(6),
            r.GetString(7), r.GetString(8), r.GetString(9), r.GetString(10), r.GetString(11), ImagesOf(images, r.GetString(0)));

    private static List<WorkingImage> ImagesOf(List<WorkingImage> images, string logId) => images.Where(i => i.LogId == logId).ToList();

    /// <summary>Picture rows (no bytes) in paste order; <paramref name="where"/> may use i. (WorkingImages) and e. (WorkingEntries).</summary>
    private static List<WorkingImage> QueryImages(SqliteConnection conn, SqliteTransaction? tx, string where, params (string, object?)[] args)
    {
        using SqliteCommand cmd = Cmd(conn,
            $"SELECT {ImageColumns} FROM WorkingImages i JOIN WorkingEntries e ON e.Id = i.EntryId WHERE {where} ORDER BY i.CreatedAt, i.Id", tx, args);
        var list = new List<WorkingImage>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new(r.GetString(0), r.GetString(1), r.GetString(2), r.GetString(3), r.GetString(4), r.GetInt64(5), r.GetString(6), r.GetString(7)));
        return list;
    }

    private static List<WorkingEntryLog> ReadLogs(SqliteConnection conn, string entryId, List<WorkingImage> images, SqliteTransaction? tx = null)
    {
        using SqliteCommand cmd = Cmd(conn, $"SELECT {LogColumns} FROM WorkingEntryLogs WHERE EntryId=@e ORDER BY Date, CreatedAt, Id", tx,
            ("@e", entryId));
        var list = new List<WorkingEntryLog>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) list.Add(ReadLog(r, images));
        return list;
    }

    private static WorkingEntry? ReadEntryRow(SqliteConnection conn, string model, string id, SqliteTransaction? tx = null)
    {
        List<WorkingImage> images = QueryImages(conn, tx, "i.EntryId=@e AND e.ModelName=@m", ("@e", id), ("@m", model));
        List<WorkingEntryLog> logs = ReadLogs(conn, id, images, tx);
        using SqliteCommand cmd = Cmd(conn, $"SELECT {EntryColumns} FROM WorkingEntries WHERE Id=@id AND ModelName=@m", tx,
            ("@id", id), ("@m", model));
        using SqliteDataReader r = cmd.ExecuteReader();
        return r.Read() ? ReadEntry(r, logs, ImagesOf(images, "")) : null;
    }

    private static bool IsDate(string value) =>
        DateOnly.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _);

    /// <summary>"" or a yyyy-MM-dd date.</summary>
    private static string OptionalDate(string? value, string label)
    {
        string v = (value ?? "").Trim();
        if (v.Length > 0 && !IsDate(v)) throw new ArgumentException($"{label} must be yyyy-MM-dd.");
        return v;
    }

    private sealed record CleanedEntry(
        string Kind, string Date, string Title, string Body, string Status, string Severity, string Result, string Owner, string Due,
        string Phase, string Category, string Process, string Cause, string DoneDate, string FinalResult);

    /// <summary>Parameters of every editable column, shared by INSERT and UPDATE.</summary>
    private static (string, object?)[] EntryArgs(CleanedEntry e) =>
    [
        ("@k", e.Kind), ("@d", e.Date), ("@t", e.Title), ("@b", e.Body), ("@s", e.Status), ("@sv", e.Severity), ("@r", e.Result),
        ("@o", e.Owner), ("@due", e.Due), ("@ph", e.Phase), ("@cat", e.Category), ("@proc", e.Process), ("@cause", e.Cause),
        ("@done", e.DoneDate), ("@fin", e.FinalResult),
    ];

    private static CleanedEntry CleanEntry(WorkingEntrySave b)
    {
        string kind = (b.Kind ?? "").Trim().ToLowerInvariant();
        if (!Kinds.Contains(kind)) throw new ArgumentException("Kind must be dev, issue or review.");
        string date = (b.Date ?? "").Trim();
        if (!IsDate(date)) throw new ArgumentException("Date must be yyyy-MM-dd.");
        string title = (b.Title ?? "").Trim();
        if (title.Length == 0) throw new ArgumentException("Title is required.");
        if (title.Length > 200) throw new ArgumentException("Title must be 200 characters or fewer.");
        string body = b.Body ?? "";
        if (body.Length > 20000) throw new ArgumentException("Body must be 20000 characters or fewer.");
        return new(kind, date, title, body, Limit(b.Status, "Status", 100), Limit(b.Severity, "Severity", 100),
            Limit(b.Result, "Result", 2000), Limit(b.Owner, "Owner", 100), OptionalDate(b.Due, "Due"),
            Limit(b.Phase, "Phase", 20), Limit(b.Category, "Issue type", 60), Limit(b.Process, "Process", 60),
            Limit(b.Cause, "Cause group", 60), OptionalDate(b.DoneDate, "Done date"), Limit(b.FinalResult, "Final result", 2000));
    }

    public List<WorkingEntry> ListEntries(string model)
    {
        model = CleanModelName(model);
        using SqliteConnection conn = Open();
        // every picture and note of the model in one read each, grouped by entry
        var images = QueryImages(conn, null, "e.ModelName=@m", ("@m", model)).GroupBy(i => i.EntryId).ToDictionary(g => g.Key, g => g.ToList());
        var logs = new Dictionary<string, List<WorkingEntryLog>>();
        using (SqliteCommand lc = Cmd(conn,
            "SELECT " + string.Join(", ", LogColumns.Split(", ").Select(c => "l." + c)) +
            " FROM WorkingEntryLogs l JOIN WorkingEntries e ON e.Id = l.EntryId WHERE e.ModelName=@m ORDER BY l.Date, l.CreatedAt, l.Id",
            null, ("@m", model)))
        using (SqliteDataReader lr = lc.ExecuteReader())
            while (lr.Read())
            {
                WorkingEntryLog log = ReadLog(lr, images.GetValueOrDefault(lr.GetString(1)) ?? []);
                if (!logs.TryGetValue(log.EntryId, out List<WorkingEntryLog>? forEntry)) logs[log.EntryId] = forEntry = [];
                forEntry.Add(log);
            }
        using SqliteCommand cmd = Cmd(conn,
            $"SELECT {EntryColumns} FROM WorkingEntries WHERE ModelName=@m ORDER BY Date DESC, CreatedAt DESC", null, ("@m", model));
        var list = new List<WorkingEntry>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(ReadEntry(r, logs.GetValueOrDefault(r.GetString(0)) ?? [], ImagesOf(images.GetValueOrDefault(r.GetString(0)) ?? [], "")));
        return list;
    }

    public WorkingEntry CreateEntry(string model, WorkingEntrySave body, string user)
    {
        model = CleanModelName(model);
        var e = CleanEntry(body);
        string id = Guid.NewGuid().ToString("N");
        string now = Now();
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        string parentId = (body.ParentId ?? "").Trim();
        // a sub-issue follows its parent's setup row
        string setupRowId = parentId.Length > 0
            ? ParentOf(conn, tx, model, e.Kind, parentId, null).SetupRowId
            : SetupRowOf(conn, tx, model, e.Kind, (body.SetupRowId ?? "").Trim());
        using (SqliteCommand ins = Cmd(conn,
            $"INSERT INTO WorkingEntries ({EntryColumns}) VALUES " +
            "(@id,@m,@k,@d,@t,@b,@s,@sv,@r,@o,@due,1,@u,@now,@u,@now,@seq,@ph,@cat,@proc,@cause,@done,@fin,@setup,@parent)", tx,
            [("@id", id), ("@m", model), ("@u", user), ("@now", now), ("@seq", NextSeq(conn, tx, model, e.Kind)), ("@setup", setupRowId), ("@parent", parentId),
             .. EntryArgs(e)]))
            ins.ExecuteNonQuery();
        WorkingEntry saved = ReadEntryRow(conn, model, id, tx)!;
        tx.Commit();
        return saved;
    }

    /// <summary>Null when the entry does not exist.</summary>
    public WorkingEntry? UpdateEntry(string model, string id, WorkingEntrySave body, string user)
    {
        model = CleanModelName(model);
        var e = CleanEntry(body);
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        WorkingEntry? current = ReadEntryRow(conn, model, id, tx);
        if (current is null) return null;
        if (current.Revision != body.Revision)
            throw new WorkingConflictException("Someone else saved this entry first. Reload and try again.");
        // a kind change gives the entry the next number of its new kind
        int seq = current.Kind == e.Kind ? current.Seq : NextSeq(conn, tx, model, e.Kind);
        // an issue saved without a SetupRowId stays on its setup row
        string parentId = body.ParentId is null ? current.ParentId : body.ParentId.Trim();
        if (parentId.Length > 0)
        {
            if (parentId == id) throw new ArgumentException("An issue cannot be its own sub-issue's parent.");
            if (HasSubIssues(conn, tx, model, id)) throw new ArgumentException("An issue that has sub-issues cannot become a sub-issue.");
        }
        else if (e.Kind != "issue" && HasSubIssues(conn, tx, model, id))
            throw new ArgumentException("An issue that has sub-issues must stay an issue.");
        string setupRowId = parentId.Length > 0 ? ParentOf(conn, tx, model, e.Kind, parentId, id).SetupRowId
            : body.SetupRowId is null && e.Kind == "issue" ? current.SetupRowId
            : SetupRowOf(conn, tx, model, e.Kind, (body.SetupRowId ?? "").Trim());
        using (SqliteCommand upd = Cmd(conn,
            "UPDATE WorkingEntries SET Kind=@k, Date=@d, Title=@t, Body=@b, Status=@s, Severity=@sv, Result=@r, Owner=@o, Due=@due, " +
            "Seq=@seq, Phase=@ph, Category=@cat, Process=@proc, Cause=@cause, DoneDate=@done, FinalResult=@fin, SetupRowId=@setup, ParentId=@parent, " +
            "Revision=Revision+1, UpdatedBy=@u, UpdatedAt=@now WHERE Id=@id AND ModelName=@m", tx,
            [("@id", id), ("@m", model), ("@u", user), ("@now", Now()), ("@seq", seq), ("@setup", setupRowId), ("@parent", parentId),
             .. EntryArgs(e)]))
            upd.ExecuteNonQuery();
        // sub-issues follow the setup row of their parent
        if (setupRowId != current.SetupRowId)
            using (SqliteCommand kids = Cmd(conn,
                "UPDATE WorkingEntries SET SetupRowId=@setup, Revision=Revision+1, UpdatedAt=@now WHERE ParentId=@id AND ModelName=@m", tx,
                ("@setup", setupRowId), ("@now", Now()), ("@id", id), ("@m", model)))
                kids.ExecuteNonQuery();
        WorkingEntry saved = ReadEntryRow(conn, model, id, tx)!;
        tx.Commit();
        return saved;
    }

    /// <summary>Deletes the entry together with its progress notes and every picture of both.</summary>
    public bool DeleteEntry(string model, string id)
    {
        model = CleanModelName(model);
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        using (SqliteCommand del = Cmd(conn, "DELETE FROM WorkingEntries WHERE Id=@id AND ModelName=@m", tx, ("@id", id), ("@m", model)))
            if (del.ExecuteNonQuery() == 0) return false;
        // its sub-issues are kept as top-level issues
        using (SqliteCommand unlink = Cmd(conn,
            "UPDATE WorkingEntries SET ParentId='', Revision=Revision+1, UpdatedAt=@now WHERE ParentId=@id AND ModelName=@m", tx,
            ("@id", id), ("@m", model), ("@now", Now())))
            unlink.ExecuteNonQuery();
        using (SqliteCommand logs = Cmd(conn, "DELETE FROM WorkingEntryLogs WHERE EntryId=@id", tx, ("@id", id)))
            logs.ExecuteNonQuery();
        using (SqliteCommand pics = Cmd(conn, "DELETE FROM WorkingImages WHERE EntryId=@id", tx, ("@id", id)))
            pics.ExecuteNonQuery();
        tx.Commit();
        return true;
    }

    private static bool HasSubIssues(SqliteConnection conn, SqliteTransaction tx, string model, string id)
    {
        using SqliteCommand cmd = Cmd(conn, "SELECT 1 FROM WorkingEntries WHERE ParentId=@id AND ModelName=@m LIMIT 1", tx, ("@id", id), ("@m", model));
        return cmd.ExecuteScalar() is not null;
    }

    /// <summary>The parent of a sub-issue: an issue of the same model that is itself top-level (one level only).</summary>
    private static WorkingEntry ParentOf(SqliteConnection conn, SqliteTransaction tx, string model, string kind, string parentId, string? selfId)
    {
        if (kind != "issue") throw new ArgumentException("Only an issue can be a sub-issue.");
        if (parentId == selfId) throw new ArgumentException("An issue cannot be its own parent.");
        WorkingEntry? parent = ReadEntryRow(conn, model, parentId, tx);
        if (parent is null || parent.Kind != "issue") throw new ArgumentException("The parent must be an issue of this model.");
        if (parent.ParentId.Length > 0) throw new ArgumentException("A sub-issue cannot have its own sub-issues.");
        return parent;
    }

    /// <summary>The setup row an entry may be linked to: "" for none; otherwise an issue on a row of the same model.</summary>
    private static string SetupRowOf(SqliteConnection conn, SqliteTransaction tx, string model, string kind, string setupRowId)
    {
        if (setupRowId.Length == 0) return "";
        if (kind != "issue") throw new ArgumentException("Only an issue can be a problem of a setup row.");
        if (ReadSetupRow(conn, tx, model, setupRowId) is null) throw new ArgumentException("That setup row is not in this model.");
        return setupRowId;
    }

    // ── Setup Plan: process rows with machine line and setup progress ───────

    private const string SetupColumns =
        "Id, ModelName, No, LayoutNo, Line, ProcessName, MachinePlan, MachineState, MachineNote, SetupPlan, SetupFinish, SetupStatus, " +
        "Revision, CreatedBy, CreatedAt, UpdatedBy, UpdatedAt";

    public const int MaxSetupRowsPerImport = 2000;

    private static WorkingSetupRow ReadSetup(SqliteDataReader r) =>
        new(r.GetString(0), r.GetString(1), r.GetInt32(2), r.GetString(3), r.GetString(4), r.GetString(5), r.GetString(6), r.GetString(7),
            r.GetString(8), r.GetString(9), r.GetString(10), r.GetString(11), r.GetInt32(12), r.GetString(13), r.GetString(14),
            r.GetString(15), r.GetString(16));

    private static WorkingSetupRow? ReadSetupRow(SqliteConnection conn, SqliteTransaction? tx, string model, string id)
    {
        using SqliteCommand cmd = Cmd(conn, $"SELECT {SetupColumns} FROM WorkingSetupRows WHERE Id=@id AND ModelName=@m", tx, ("@id", id), ("@m", model));
        using SqliteDataReader r = cmd.ExecuteReader();
        return r.Read() ? ReadSetup(r) : null;
    }

    /// <summary>"", "O", "X" or "Δ"; the look-alikes typed in Excel (o, ○, x, ×, △, ▲ …) are taken as those.</summary>
    public static string SetupState(string? value, string label) => (value ?? "").Trim() switch
    {
        "" => "",
        "O" or "o" or "0" or "○" or "◯" or "●" => "O",
        "X" or "x" or "×" or "✕" => "X",
        "Δ" or "△" or "▲" or "δ" => "Δ",
        _ => throw new ArgumentException($"{label} must be O (OK), X (not yet set up) or Δ (setting up now)."),
    };

    private sealed record CleanedSetup(int No, string LayoutNo, string Line, string ProcessName, string MachinePlan, string MachineState,
        string MachineNote, string SetupPlan, string SetupFinish, string SetupStatus);

    private static CleanedSetup CleanSetup(WorkingSetupRowSave b)
    {
        string process = (b.ProcessName ?? "").Trim();
        if (process.Length == 0) throw new ArgumentException("Process name is required.");
        if (process.Length > 200) throw new ArgumentException("Process name must be 200 characters or fewer.");
        int no = b.No ?? 0;
        if (no is < 0 or > 100000) throw new ArgumentException("No. must be between 1 and 100000.");
        return new(no, Limit(b.LayoutNo, "No. layout", 20), Limit(b.Line, "Line", 50), process,
            OptionalDate(b.MachinePlan, "Machine plan"), SetupState(b.MachineState, "Machine state"), Limit(b.MachineNote, "Machine note", 100),
            OptionalDate(b.SetupPlan, "Setup plan"), OptionalDate(b.SetupFinish, "Setup finish"), SetupState(b.SetupStatus, "Setup status"));
    }

    private static (string, object?)[] SetupArgs(CleanedSetup c) =>
    [
        ("@no", c.No), ("@lay", c.LayoutNo), ("@line", c.Line), ("@proc", c.ProcessName), ("@mp", c.MachinePlan), ("@ms", c.MachineState),
        ("@mn", c.MachineNote), ("@sp", c.SetupPlan), ("@sf", c.SetupFinish), ("@ss", c.SetupStatus),
    ];

    /// <summary>The model's setup rows by No. (then in the order they were added).</summary>
    public List<WorkingSetupRow> ListSetupRows(string model)
    {
        model = CleanModelName(model);
        using SqliteConnection conn = Open();
        using SqliteCommand cmd = Cmd(conn, $"SELECT {SetupColumns} FROM WorkingSetupRows WHERE ModelName=@m ORDER BY No, CreatedAt, Id", null, ("@m", model));
        var list = new List<WorkingSetupRow>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) list.Add(ReadSetup(r));
        return list;
    }

    public WorkingSetupRow CreateSetupRow(string model, WorkingSetupRowSave body, string user) => ImportSetupRows(model, [body], user)[0];

    /// <summary>Adds rows in one go (all or nothing); a row without a No. takes the one after the last.</summary>
    public List<WorkingSetupRow> ImportSetupRows(string model, IReadOnlyList<WorkingSetupRowSave> rows, string user)
    {
        model = CleanModelName(model);
        if (rows.Count == 0) throw new ArgumentException("Add at least one row.");
        if (rows.Count > MaxSetupRowsPerImport) throw new ArgumentException($"Add at most {MaxSetupRowsPerImport} rows at a time.");
        var clean = new List<CleanedSetup>();
        for (int i = 0; i < rows.Count; i++)
        {
            try { clean.Add(CleanSetup(rows[i])); }
            catch (ArgumentException ex) when (rows.Count > 1) { throw new ArgumentException($"Row {i + 1}: {ex.Message}"); }
        }
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        int last;
        using (SqliteCommand max = Cmd(conn, "SELECT COALESCE(MAX(No), 0) FROM WorkingSetupRows WHERE ModelName=@m", tx, ("@m", model)))
            last = Convert.ToInt32(max.ExecuteScalar(), CultureInfo.InvariantCulture);
        var ids = new List<string>();
        DateTime at = DateTime.UtcNow;
        foreach (CleanedSetup c in clean)
        {
            string id = Guid.NewGuid().ToString("N");
            int no = c.No > 0 ? c.No : last + 1;
            last = Math.Max(last, no);
            // a tick apart so rows with the same No. keep the order they came in
            string now = at.AddTicks(ids.Count).ToString("o", CultureInfo.InvariantCulture);
            using SqliteCommand ins = Cmd(conn,
                $"INSERT INTO WorkingSetupRows ({SetupColumns}) VALUES (@id,@m,@no,@lay,@line,@proc,@mp,@ms,@mn,@sp,@sf,@ss,1,@u,@now,@u,@now)", tx,
                [("@id", id), ("@m", model), ("@u", user), ("@now", now), .. SetupArgs(c with { No = no })]);
            ins.ExecuteNonQuery();
            ids.Add(id);
        }
        List<WorkingSetupRow> saved = ids.Select(id => ReadSetupRow(conn, tx, model, id)!).ToList();
        tx.Commit();
        return saved;
    }

    /// <summary>Null when the row is not in this model; a stale revision is a conflict.</summary>
    public WorkingSetupRow? UpdateSetupRow(string model, string id, WorkingSetupRowSave body, string user)
    {
        model = CleanModelName(model);
        CleanedSetup c = CleanSetup(body);
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        WorkingSetupRow? current = ReadSetupRow(conn, tx, model, id);
        if (current is null) return null;
        if (current.Revision != body.Revision)
            throw new WorkingConflictException("Someone else saved this setup row first. Reload and try again.");
        using (SqliteCommand upd = Cmd(conn,
            "UPDATE WorkingSetupRows SET No=@no, LayoutNo=@lay, Line=@line, ProcessName=@proc, MachinePlan=@mp, MachineState=@ms, " +
            "MachineNote=@mn, SetupPlan=@sp, SetupFinish=@sf, SetupStatus=@ss, Revision=Revision+1, UpdatedBy=@u, UpdatedAt=@now " +
            "WHERE Id=@id AND ModelName=@m", tx,
            [("@id", id), ("@m", model), ("@u", user), ("@now", Now()), .. SetupArgs(c with { No = c.No > 0 ? c.No : current.No })]))
            upd.ExecuteNonQuery();
        WorkingSetupRow saved = ReadSetupRow(conn, tx, model, id)!;
        tx.Commit();
        return saved;
    }

    /// <summary>Deletes the row. Its problems are kept as ordinary issues, no longer linked to a row.</summary>
    public bool DeleteSetupRow(string model, string id)
    {
        model = CleanModelName(model);
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        using (SqliteCommand del = Cmd(conn, "DELETE FROM WorkingSetupRows WHERE Id=@id AND ModelName=@m", tx, ("@id", id), ("@m", model)))
            if (del.ExecuteNonQuery() == 0) return false;
        using (SqliteCommand unlink = Cmd(conn,
            "UPDATE WorkingEntries SET SetupRowId='', Revision=Revision+1, UpdatedAt=@now WHERE SetupRowId=@id AND ModelName=@m", tx,
            ("@id", id), ("@m", model), ("@now", Now())))
            unlink.ExecuteNonQuery();
        tx.Commit();
        return true;
    }

    // ── Progress notes: dated notes under an entry; they do not change the entry's revision ──

    private sealed record CleanedLog(string Date, string Text, string Owner, string Due, string DoneDate);

    private static CleanedLog CleanLog(WorkingLogSave b)
    {
        string date = (b.Date ?? "").Trim();
        if (!IsDate(date)) throw new ArgumentException("Date must be yyyy-MM-dd.");
        string text = (b.Text ?? "").Trim();
        if (text.Length == 0) throw new ArgumentException("Note text is required.");
        if (text.Length > 5000) throw new ArgumentException("Note text must be 5000 characters or fewer.");
        return new(date, text, Limit(b.Owner, "Owner", 100), OptionalDate(b.Due, "Due"), OptionalDate(b.DoneDate, "Done date"));
    }

    private static (string, object?)[] LogArgs(CleanedLog l) =>
        [("@d", l.Date), ("@t", l.Text), ("@o", l.Owner), ("@due", l.Due), ("@done", l.DoneDate)];

    private static bool EntryExists(SqliteConnection conn, SqliteTransaction tx, string model, string entryId)
    {
        using SqliteCommand cmd = Cmd(conn, "SELECT 1 FROM WorkingEntries WHERE Id=@id AND ModelName=@m", tx, ("@id", entryId), ("@m", model));
        return cmd.ExecuteScalar() is not null;
    }

    private static WorkingEntryLog? ReadLogRow(SqliteConnection conn, SqliteTransaction tx, string entryId, string logId)
    {
        List<WorkingImage> images = QueryImages(conn, tx, "i.EntryId=@e AND i.LogId=@l", ("@e", entryId), ("@l", logId));
        using SqliteCommand cmd = Cmd(conn, $"SELECT {LogColumns} FROM WorkingEntryLogs WHERE Id=@id AND EntryId=@e", tx,
            ("@id", logId), ("@e", entryId));
        using SqliteDataReader r = cmd.ExecuteReader();
        return r.Read() ? ReadLog(r, images) : null;
    }

    /// <summary>Null when the entry is not in this model.</summary>
    public WorkingEntryLog? CreateLog(string model, string entryId, WorkingLogSave body, string user)
    {
        model = CleanModelName(model);
        CleanedLog log = CleanLog(body);
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        if (!EntryExists(conn, tx, model, entryId)) return null;
        string id = Guid.NewGuid().ToString("N"), now = Now();
        using (SqliteCommand ins = Cmd(conn,
            $"INSERT INTO WorkingEntryLogs ({LogColumns}) VALUES (@id,@e,@d,@t,1,@u,@now,@u,@now,@o,@due,@done)", tx,
            [("@id", id), ("@e", entryId), ("@u", user), ("@now", now), .. LogArgs(log)]))
            ins.ExecuteNonQuery();
        WorkingEntryLog saved = ReadLogRow(conn, tx, entryId, id)!;
        tx.Commit();
        return saved;
    }

    /// <summary>Null when the entry or the note does not exist.</summary>
    public WorkingEntryLog? UpdateLog(string model, string entryId, string logId, WorkingLogSave body, string user)
    {
        model = CleanModelName(model);
        CleanedLog log = CleanLog(body);
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        if (!EntryExists(conn, tx, model, entryId)) return null;
        WorkingEntryLog? current = ReadLogRow(conn, tx, entryId, logId);
        if (current is null) return null;
        if (current.Revision != body.Revision)
            throw new WorkingConflictException("Someone else saved this note first. Reload and try again.");
        using (SqliteCommand upd = Cmd(conn,
            "UPDATE WorkingEntryLogs SET Date=@d, Text=@t, Owner=@o, Due=@due, DoneDate=@done, Revision=Revision+1, UpdatedBy=@u, " +
            "UpdatedAt=@now WHERE Id=@id AND EntryId=@e", tx,
            [("@id", logId), ("@e", entryId), ("@u", user), ("@now", Now()), .. LogArgs(log)]))
            upd.ExecuteNonQuery();
        WorkingEntryLog saved = ReadLogRow(conn, tx, entryId, logId)!;
        tx.Commit();
        return saved;
    }

    /// <summary>Deletes the note and its pictures.</summary>
    public bool DeleteLog(string model, string entryId, string logId)
    {
        model = CleanModelName(model);
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        if (!EntryExists(conn, tx, model, entryId)) return false;
        bool removed;
        using (SqliteCommand del = Cmd(conn, "DELETE FROM WorkingEntryLogs WHERE Id=@id AND EntryId=@e", tx, ("@id", logId), ("@e", entryId)))
            removed = del.ExecuteNonQuery() > 0;
        if (removed)
            using (SqliteCommand pics = Cmd(conn, "DELETE FROM WorkingImages WHERE EntryId=@e AND LogId=@id", tx, ("@id", logId), ("@e", entryId)))
                pics.ExecuteNonQuery();
        tx.Commit();
        return removed;
    }

    // ── Pictures and videos pasted into an entry or a note ──────────────────

    /// <summary>ISO media brands of a video; a file whose own brand is audio or a still picture (HEIC, AVIF, M4A …) is not one.</summary>
    private static readonly string[] VideoBrands =
        ["isom", "iso2", "iso4", "iso5", "iso6", "mp41", "mp42", "avc1", "M4V ", "M4VH", "M4VP", "dash", "qt  ", "3gp4", "3gp5", "3gp6", "3g2a", "mmp4", "XAVC", "MSNV"];
    private static readonly string[] NotVideoBrands = ["M4A ", "M4B ", "M4P ", "F4A ", "F4B ", "heic", "heix", "mif1", "msf1", "avif", "avis", "crx "];

    /// <summary>
    /// The format from the first bytes (<see cref="SniffBytes"/> are enough), or null when the bytes are not a PNG / JPEG / GIF /
    /// WebP / BMP picture or an MP4 / MOV / WebM / MKV / AVI video. <paramref name="length"/> is the whole file's length.
    /// </summary>
    public static string? FileType(ReadOnlySpan<byte> head, long length)
    {
        if (At(head, 0, [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])) return "image/png";
        if (At(head, 0, [0xFF, 0xD8, 0xFF])) return "image/jpeg";
        if (At(head, 0, "GIF8"u8)) return "image/gif";
        if (At(head, 0, "RIFF"u8) && At(head, 8, "WEBP"u8)) return "image/webp";
        if (At(head, 0, "RIFF"u8) && At(head, 8, "AVI "u8)) return "video/x-msvideo";
        if (At(head, 0, "BM"u8) && length > 26) return "image/bmp";
        if (At(head, 0, [0x1A, 0x45, 0xDF, 0xA3]))
            return head.IndexOf("webm"u8) >= 0 ? "video/webm" : head.IndexOf("matroska"u8) >= 0 ? "video/x-matroska" : null;
        if (At(head, 4, "ftyp"u8) && head.Length >= 12)
        {
            string major = System.Text.Encoding.ASCII.GetString(head.Slice(8, 4));
            if (NotVideoBrands.Contains(major)) return null;
            // the ftyp box: size, "ftyp", major brand, minor version, then the compatible brands
            long boxSize = System.Buffers.Binary.BinaryPrimitives.ReadUInt32BigEndian(head);
            int boxEnd = (int)Math.Min(head.Length, boxSize);
            bool video = VideoBrands.Contains(major);
            for (int i = 16; !video && i + 4 <= boxEnd; i += 4)
                video = VideoBrands.Contains(System.Text.Encoding.ASCII.GetString(head.Slice(i, 4)));
            return video ? (major == "qt  " ? "video/quicktime" : "video/mp4") : null;
        }
        return null;
    }

    private static bool At(ReadOnlySpan<byte> head, int offset, ReadOnlySpan<byte> sig) =>
        head.Length >= offset + sig.Length && head.Slice(offset, sig.Length).SequenceEqual(sig);

    /// <summary>Stores pictures / videos given as bytes; see <see cref="AddFiles"/>.</summary>
    public List<WorkingImage>? AddImages(string model, string entryId, string? logId, IReadOnlyList<(string Name, byte[] Data)> files, string user) =>
        AddFiles(model, entryId, logId, files.Select(f => WorkingUpload.Of(f.Name, f.Data)).ToList(), user);

    /// <summary>
    /// Stores pictures and videos under the entry, or under its note <paramref name="logId"/>, byte for byte as sent. Null when the
    /// entry is not in the model or the note is not under the entry. Every file is checked before anything is stored; the format
    /// comes from the bytes. Each file is streamed into its row, so a video is never held in memory whole.
    /// </summary>
    public List<WorkingImage>? AddFiles(string model, string entryId, string? logId, IReadOnlyList<WorkingUpload> files, string user)
    {
        model = CleanModelName(model);
        logId = (logId ?? "").Trim();
        if (files.Count == 0) throw new ArgumentException("Add at least one picture or video.");
        if (files.Count > MaxImagesPerUpload) throw new ArgumentException($"Add at most {MaxImagesPerUpload} files at a time.");
        if (files.Sum(f => f.Length) > MaxUploadBytes) throw new ArgumentException("One upload must be 1 GB or less in total.");
        var clean = new List<(string Name, string Type, WorkingUpload File)>();
        foreach (WorkingUpload file in files)
        {
            string name = Path.GetFileName((file.Name ?? "").Trim());
            if (name.Length == 0) name = "file";
            if (name.Length > 200) name = name[^200..];
            byte[] head = new byte[SniffBytes];
            int read = 0;
            using (Stream s = file.Open())
                for (int n; read < head.Length && (n = s.Read(head, read, head.Length - read)) > 0;) read += n;
            string type = FileType(head.AsSpan(0, read), file.Length)
                ?? throw new ArgumentException($"{name} is not a PNG, JPEG, GIF, WebP or BMP picture or an MP4, MOV, WebM, MKV or AVI video.");
            if (type.StartsWith("image/", StringComparison.Ordinal) && file.Length > MaxImageBytes)
                throw new ArgumentException($"{name} is larger than 10 MB.");
            if (type.StartsWith("video/", StringComparison.Ordinal) && file.Length > MaxVideoBytes)
                throw new ArgumentException($"{name} is larger than 500 MB.");
            clean.Add((name, type, file));
        }
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        if (!EntryExists(conn, tx, model, entryId)) return null;
        if (logId.Length > 0 && ReadLogRow(conn, tx, entryId, logId) is null) return null;
        var ids = new List<string>();
        DateTime at = DateTime.UtcNow;
        foreach ((string name, string type, WorkingUpload file) in clean)
        {
            string id = Guid.NewGuid().ToString("N");
            // a tick apart so the paste order is kept
            string created = at.AddTicks(ids.Count).ToString("o", CultureInfo.InvariantCulture);
            long rowId;
            using (SqliteCommand ins = Cmd(conn,
                "INSERT INTO WorkingImages (Id, EntryId, LogId, FileName, ContentType, Size, Data, CreatedBy, CreatedAt) " +
                "VALUES (@id,@e,@l,@n,@ct,@sz,zeroblob(@sz),@u,@at); SELECT last_insert_rowid();", tx,
                ("@id", id), ("@e", entryId), ("@l", logId), ("@n", name), ("@ct", type), ("@sz", file.Length), ("@u", user), ("@at", created)))
                rowId = (long)ins.ExecuteScalar()!;
            long copied = 0;
            using (var blob = new SqliteBlob(conn, "WorkingImages", "Data", rowId))
            using (Stream src = file.Open())
            {
                byte[] buf = new byte[81920];
                for (int n; (n = src.Read(buf, 0, buf.Length)) > 0;)
                {
                    if (copied + n > file.Length) throw new ArgumentException($"{name} is longer than its stated size.");
                    blob.Write(buf, 0, n);
                    copied += n;
                }
            }
            if (copied != file.Length) throw new ArgumentException($"{name} was cut short.");
            ids.Add(id);
        }
        List<WorkingImage> saved = QueryImages(conn, tx, "i.EntryId=@e AND i.LogId=@l", ("@e", entryId), ("@l", logId))
            .Where(i => ids.Contains(i.Id)).ToList();
        tx.Commit();
        return saved;
    }

    /// <summary>
    /// A stored picture / video opened for reading (seekable, so byte ranges can be served), or null. Dispose it when done;
    /// it keeps its own DB connection open until then.
    /// </summary>
    public WorkingImageFile? OpenImage(string id)
    {
        SqliteConnection conn = Open();
        try
        {
            long rowId;
            string fileName, contentType;
            using (SqliteCommand cmd = Cmd(conn, "SELECT rowid, FileName, ContentType FROM WorkingImages WHERE Id=@id", null, ("@id", id)))
            using (SqliteDataReader r = cmd.ExecuteReader())
            {
                if (!r.Read()) { conn.Dispose(); return null; }
                (rowId, fileName, contentType) = (r.GetInt64(0), r.GetString(1), r.GetString(2));
            }
            var blob = new SqliteBlob(conn, "WorkingImages", "Data", rowId, readOnly: true);
            return new WorkingImageFile(fileName, contentType, blob.Length, new OwnedStream(blob, conn));
        }
        catch
        {
            conn.Dispose();
            throw;
        }
    }

    /// <summary>The bytes of a picture / video, or null (the whole file in memory; the endpoint streams with OpenImage).</summary>
    public (string FileName, string ContentType, byte[] Data)? GetImage(string id)
    {
        using WorkingImageFile? file = OpenImage(id);
        if (file is null) return null;
        using var copy = new MemoryStream();
        file.Data.CopyTo(copy);
        return (file.FileName, file.ContentType, copy.ToArray());
    }

    /// <summary>False when the picture is not under that entry of that model.</summary>
    public bool DeleteImage(string model, string entryId, string imageId)
    {
        model = CleanModelName(model);
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        if (!EntryExists(conn, tx, model, entryId)) return false;
        bool removed;
        using (SqliteCommand del = Cmd(conn, "DELETE FROM WorkingImages WHERE Id=@id AND EntryId=@e", tx, ("@id", imageId), ("@e", entryId)))
            removed = del.ExecuteNonQuery() > 0;
        tx.Commit();
        return removed;
    }

    /// <summary>A blob stream that also closes the connection it reads through (disposing twice is harmless).</summary>
    private sealed class OwnedStream(SqliteBlob blob, SqliteConnection conn) : Stream
    {
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => blob.Length;
        public override long Position { get => blob.Position; set => blob.Position = value; }
        public override int Read(byte[] buffer, int offset, int count) => blob.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => blob.Seek(offset, origin);
        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing)
        {
            if (disposing) { blob.Dispose(); conn.Dispose(); }
            base.Dispose(disposing);
        }
    }
}
