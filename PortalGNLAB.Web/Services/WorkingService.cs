using System.Globalization;
using System.Text.Json;
using Microsoft.Data.Sqlite;

namespace PortalGNLAB.Web.Services;

public sealed class WorkingConflictException(string message) : Exception(message);

public sealed record WorkingGroup(string Id, string Name, string? ParentId, List<string> Models, int Sort, int Position);

public sealed record WorkingGroups(List<WorkingGroup> Groups, string Source, string? Error);

public sealed record WorkingModel(string Name, JsonElement Info, int Revision, string UpdatedAt, string UpdatedBy);

public sealed record WorkingModelSave(JsonElement Info, int Revision);

/// <summary>An entry. Issue fields (Seq … FinalResult) are "" / 0 for other kinds; Logs are the dated progress notes, oldest first.</summary>
public sealed record WorkingEntry(
    string Id, string ModelName, string Kind, string Date, string Title, string Body, string Status, string Severity,
    string Result, string Owner, string Due, int Revision, string CreatedBy, string CreatedAt, string UpdatedBy, string UpdatedAt,
    int Seq, string Phase, string Category, string Process, string Cause, string DoneDate, string FinalResult, List<WorkingEntryLog> Logs);

public sealed record WorkingEntrySave(
    string? Kind, string? Date, string? Title, string? Body, string? Status, string? Severity,
    string? Result, string? Owner, string? Due, int Revision,
    string? Phase = null, string? Category = null, string? Process = null, string? Cause = null, string? DoneDate = null, string? FinalResult = null);

public sealed record WorkingEntryLog(
    string Id, string EntryId, string Date, string Text, int Revision, string CreatedBy, string CreatedAt, string UpdatedBy, string UpdatedAt);

public sealed record WorkingLogSave(string? Date, string? Text, int Revision = 0);

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
        CREATE TABLE IF NOT EXISTS WorkingEntrySeq (
            ModelName TEXT NOT NULL COLLATE NOCASE,
            Kind TEXT NOT NULL,
            Last INTEGER NOT NULL,
            PRIMARY KEY (ModelName, Kind)
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
    ];

    public const string DefaultGlossaryDb = @"E:\db\CompanyGlossary\company-glossary.db";
    public const string GlossaryEnvVar = "PORTALGNLAB_GLOSSARY_DB";
    public static readonly string[] Kinds = ["dev", "issue", "review"];

    private const string EntryColumns =
        "Id, ModelName, Kind, Date, Title, Body, Status, Severity, Result, Owner, Due, Revision, CreatedBy, CreatedAt, UpdatedBy, UpdatedAt, " +
        "Seq, Phase, Category, Process, Cause, DoneDate, FinalResult";

    private const string LogColumns = "Id, EntryId, Date, Text, Revision, CreatedBy, CreatedAt, UpdatedBy, UpdatedAt";

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

    /// <summary>Adds the issue columns missing from an older WorkingEntries and numbers rows that have no Seq yet.</summary>
    private static void MigrateEntries(SqliteConnection conn)
    {
        var have = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using (SqliteCommand info = Cmd(conn, "PRAGMA table_info(WorkingEntries)"))
        using (SqliteDataReader r = info.ExecuteReader())
            while (r.Read()) have.Add(r.GetString(1));
        using SqliteTransaction tx = conn.BeginTransaction();
        foreach ((string name, string type) in AddedEntryColumns.Where(c => !have.Contains(c.Name)))
            using (SqliteCommand alter = Cmd(conn, $"ALTER TABLE WorkingEntries ADD COLUMN {name} {type}", tx))
                alter.ExecuteNonQuery();
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

    private static WorkingEntry ReadEntry(SqliteDataReader r, List<WorkingEntryLog> logs) =>
        new(r.GetString(0), r.GetString(1), r.GetString(2), r.GetString(3), r.GetString(4), r.GetString(5), r.GetString(6),
            r.GetString(7), r.GetString(8), r.GetString(9), r.GetString(10), r.GetInt32(11), r.GetString(12), r.GetString(13),
            r.GetString(14), r.GetString(15), r.GetInt32(16), r.GetString(17), r.GetString(18), r.GetString(19), r.GetString(20),
            r.GetString(21), r.GetString(22), logs);

    private static WorkingEntryLog ReadLog(SqliteDataReader r) =>
        new(r.GetString(0), r.GetString(1), r.GetString(2), r.GetString(3), r.GetInt32(4), r.GetString(5), r.GetString(6),
            r.GetString(7), r.GetString(8));

    private static List<WorkingEntryLog> ReadLogs(SqliteConnection conn, string entryId, SqliteTransaction? tx = null)
    {
        using SqliteCommand cmd = Cmd(conn, $"SELECT {LogColumns} FROM WorkingEntryLogs WHERE EntryId=@e ORDER BY Date, CreatedAt, Id", tx,
            ("@e", entryId));
        var list = new List<WorkingEntryLog>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) list.Add(ReadLog(r));
        return list;
    }

    private static WorkingEntry? ReadEntryRow(SqliteConnection conn, string model, string id, SqliteTransaction? tx = null)
    {
        List<WorkingEntryLog> logs = ReadLogs(conn, id, tx);
        using SqliteCommand cmd = Cmd(conn, $"SELECT {EntryColumns} FROM WorkingEntries WHERE Id=@id AND ModelName=@m", tx,
            ("@id", id), ("@m", model));
        using SqliteDataReader r = cmd.ExecuteReader();
        return r.Read() ? ReadEntry(r, logs) : null;
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
        // every note of the model in one read, grouped by entry
        var logs = new Dictionary<string, List<WorkingEntryLog>>();
        using (SqliteCommand lc = Cmd(conn,
            "SELECT l.Id, l.EntryId, l.Date, l.Text, l.Revision, l.CreatedBy, l.CreatedAt, l.UpdatedBy, l.UpdatedAt " +
            "FROM WorkingEntryLogs l JOIN WorkingEntries e ON e.Id = l.EntryId WHERE e.ModelName=@m ORDER BY l.Date, l.CreatedAt, l.Id",
            null, ("@m", model)))
        using (SqliteDataReader lr = lc.ExecuteReader())
            while (lr.Read())
            {
                WorkingEntryLog log = ReadLog(lr);
                if (!logs.TryGetValue(log.EntryId, out List<WorkingEntryLog>? forEntry)) logs[log.EntryId] = forEntry = [];
                forEntry.Add(log);
            }
        using SqliteCommand cmd = Cmd(conn,
            $"SELECT {EntryColumns} FROM WorkingEntries WHERE ModelName=@m ORDER BY Date DESC, CreatedAt DESC", null, ("@m", model));
        var list = new List<WorkingEntry>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) list.Add(ReadEntry(r, logs.GetValueOrDefault(r.GetString(0)) ?? []));
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
        using (SqliteCommand ins = Cmd(conn,
            $"INSERT INTO WorkingEntries ({EntryColumns}) VALUES " +
            "(@id,@m,@k,@d,@t,@b,@s,@sv,@r,@o,@due,1,@u,@now,@u,@now,@seq,@ph,@cat,@proc,@cause,@done,@fin)", tx,
            [("@id", id), ("@m", model), ("@u", user), ("@now", now), ("@seq", NextSeq(conn, tx, model, e.Kind)), .. EntryArgs(e)]))
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
        using (SqliteCommand upd = Cmd(conn,
            "UPDATE WorkingEntries SET Kind=@k, Date=@d, Title=@t, Body=@b, Status=@s, Severity=@sv, Result=@r, Owner=@o, Due=@due, " +
            "Seq=@seq, Phase=@ph, Category=@cat, Process=@proc, Cause=@cause, DoneDate=@done, FinalResult=@fin, " +
            "Revision=Revision+1, UpdatedBy=@u, UpdatedAt=@now WHERE Id=@id AND ModelName=@m", tx,
            [("@id", id), ("@m", model), ("@u", user), ("@now", Now()), ("@seq", seq), .. EntryArgs(e)]))
            upd.ExecuteNonQuery();
        WorkingEntry saved = ReadEntryRow(conn, model, id, tx)!;
        tx.Commit();
        return saved;
    }

    /// <summary>Deletes the entry together with its progress notes.</summary>
    public bool DeleteEntry(string model, string id)
    {
        model = CleanModelName(model);
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        using (SqliteCommand del = Cmd(conn, "DELETE FROM WorkingEntries WHERE Id=@id AND ModelName=@m", tx, ("@id", id), ("@m", model)))
            if (del.ExecuteNonQuery() == 0) return false;
        using (SqliteCommand logs = Cmd(conn, "DELETE FROM WorkingEntryLogs WHERE EntryId=@id", tx, ("@id", id)))
            logs.ExecuteNonQuery();
        tx.Commit();
        return true;
    }

    // ── Progress notes: dated notes under an entry; they do not change the entry's revision ──

    private static (string Date, string Text) CleanLog(WorkingLogSave b)
    {
        string date = (b.Date ?? "").Trim();
        if (!IsDate(date)) throw new ArgumentException("Date must be yyyy-MM-dd.");
        string text = (b.Text ?? "").Trim();
        if (text.Length == 0) throw new ArgumentException("Note text is required.");
        if (text.Length > 5000) throw new ArgumentException("Note text must be 5000 characters or fewer.");
        return (date, text);
    }

    private static bool EntryExists(SqliteConnection conn, SqliteTransaction tx, string model, string entryId)
    {
        using SqliteCommand cmd = Cmd(conn, "SELECT 1 FROM WorkingEntries WHERE Id=@id AND ModelName=@m", tx, ("@id", entryId), ("@m", model));
        return cmd.ExecuteScalar() is not null;
    }

    private static WorkingEntryLog? ReadLogRow(SqliteConnection conn, SqliteTransaction tx, string entryId, string logId)
    {
        using SqliteCommand cmd = Cmd(conn, $"SELECT {LogColumns} FROM WorkingEntryLogs WHERE Id=@id AND EntryId=@e", tx,
            ("@id", logId), ("@e", entryId));
        using SqliteDataReader r = cmd.ExecuteReader();
        return r.Read() ? ReadLog(r) : null;
    }

    /// <summary>Null when the entry is not in this model.</summary>
    public WorkingEntryLog? CreateLog(string model, string entryId, WorkingLogSave body, string user)
    {
        model = CleanModelName(model);
        (string date, string text) = CleanLog(body);
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        if (!EntryExists(conn, tx, model, entryId)) return null;
        string id = Guid.NewGuid().ToString("N"), now = Now();
        using (SqliteCommand ins = Cmd(conn,
            $"INSERT INTO WorkingEntryLogs ({LogColumns}) VALUES (@id,@e,@d,@t,1,@u,@now,@u,@now)", tx,
            ("@id", id), ("@e", entryId), ("@d", date), ("@t", text), ("@u", user), ("@now", now)))
            ins.ExecuteNonQuery();
        WorkingEntryLog saved = ReadLogRow(conn, tx, entryId, id)!;
        tx.Commit();
        return saved;
    }

    /// <summary>Null when the entry or the note does not exist.</summary>
    public WorkingEntryLog? UpdateLog(string model, string entryId, string logId, WorkingLogSave body, string user)
    {
        model = CleanModelName(model);
        (string date, string text) = CleanLog(body);
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        if (!EntryExists(conn, tx, model, entryId)) return null;
        WorkingEntryLog? current = ReadLogRow(conn, tx, entryId, logId);
        if (current is null) return null;
        if (current.Revision != body.Revision)
            throw new WorkingConflictException("Someone else saved this note first. Reload and try again.");
        using (SqliteCommand upd = Cmd(conn,
            "UPDATE WorkingEntryLogs SET Date=@d, Text=@t, Revision=Revision+1, UpdatedBy=@u, UpdatedAt=@now WHERE Id=@id AND EntryId=@e", tx,
            ("@id", logId), ("@e", entryId), ("@d", date), ("@t", text), ("@u", user), ("@now", Now())))
            upd.ExecuteNonQuery();
        WorkingEntryLog saved = ReadLogRow(conn, tx, entryId, logId)!;
        tx.Commit();
        return saved;
    }

    public bool DeleteLog(string model, string entryId, string logId)
    {
        model = CleanModelName(model);
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        if (!EntryExists(conn, tx, model, entryId)) return false;
        bool removed;
        using (SqliteCommand del = Cmd(conn, "DELETE FROM WorkingEntryLogs WHERE Id=@id AND EntryId=@e", tx, ("@id", logId), ("@e", entryId)))
            removed = del.ExecuteNonQuery() > 0;
        tx.Commit();
        return removed;
    }
}
