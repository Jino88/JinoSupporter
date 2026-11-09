using System.Globalization;
using System.Text.Json;
using Microsoft.Data.Sqlite;

namespace PortalGNLAB.Web.Services;

public sealed class WorkingConflictException(string message) : Exception(message);

public sealed record WorkingGroup(string Id, string Name, string? ParentId, List<string> Models, int Sort, int Position);

public sealed record WorkingGroups(List<WorkingGroup> Groups, string Source, string? Error);

public sealed record WorkingModel(string Name, JsonElement Info, int Revision, string UpdatedAt, string UpdatedBy);

public sealed record WorkingModelSave(JsonElement Info, int Revision);

public sealed record WorkingEntry(
    string Id, string ModelName, string Kind, string Date, string Title, string Body, string Status, string Severity,
    string Result, string Owner, string Due, int Revision, string CreatedBy, string CreatedAt, string UpdatedBy, string UpdatedAt);

public sealed record WorkingEntrySave(
    string? Kind, string? Date, string? Title, string? Body, string? Status, string? Severity,
    string? Result, string? Owner, string? Due, int Revision);

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
        """;

    public const string DefaultGlossaryDb = @"E:\db\CompanyGlossary\company-glossary.db";
    public const string GlossaryEnvVar = "PORTALGNLAB_GLOSSARY_DB";
    public static readonly string[] Kinds = ["dev", "issue", "review"];

    private const string EntryColumns =
        "Id, ModelName, Kind, Date, Title, Body, Status, Severity, Result, Owner, Due, Revision, CreatedBy, CreatedAt, UpdatedBy, UpdatedAt";

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
                SchemaReady.Add(key);
            }
        }
        return conn;
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

    private static WorkingEntry ReadEntry(SqliteDataReader r) =>
        new(r.GetString(0), r.GetString(1), r.GetString(2), r.GetString(3), r.GetString(4), r.GetString(5), r.GetString(6),
            r.GetString(7), r.GetString(8), r.GetString(9), r.GetString(10), r.GetInt32(11), r.GetString(12), r.GetString(13),
            r.GetString(14), r.GetString(15));

    private static WorkingEntry? ReadEntryRow(SqliteConnection conn, string model, string id, SqliteTransaction? tx = null)
    {
        using SqliteCommand cmd = Cmd(conn, $"SELECT {EntryColumns} FROM WorkingEntries WHERE Id=@id AND ModelName=@m", tx,
            ("@id", id), ("@m", model));
        using SqliteDataReader r = cmd.ExecuteReader();
        return r.Read() ? ReadEntry(r) : null;
    }

    private static (string Kind, string Date, string Title, string Body, string Status, string Severity, string Result, string Owner, string Due)
        CleanEntry(WorkingEntrySave b)
    {
        string kind = (b.Kind ?? "").Trim().ToLowerInvariant();
        if (!Kinds.Contains(kind)) throw new ArgumentException("Kind must be dev, issue or review.");
        string date = (b.Date ?? "").Trim();
        if (!DateOnly.TryParseExact(date, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
            throw new ArgumentException("Date must be yyyy-MM-dd.");
        string title = (b.Title ?? "").Trim();
        if (title.Length == 0) throw new ArgumentException("Title is required.");
        if (title.Length > 200) throw new ArgumentException("Title must be 200 characters or fewer.");
        string body = b.Body ?? "";
        if (body.Length > 20000) throw new ArgumentException("Body must be 20000 characters or fewer.");
        return (kind, date, title, body, Limit(b.Status, "Status", 100), Limit(b.Severity, "Severity", 100),
            Limit(b.Result, "Result", 2000), Limit(b.Owner, "Owner", 100), Limit(b.Due, "Due", 40));
    }

    public List<WorkingEntry> ListEntries(string model)
    {
        model = CleanModelName(model);
        using SqliteConnection conn = Open();
        using SqliteCommand cmd = Cmd(conn,
            $"SELECT {EntryColumns} FROM WorkingEntries WHERE ModelName=@m ORDER BY Date DESC, CreatedAt DESC", null, ("@m", model));
        var list = new List<WorkingEntry>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) list.Add(ReadEntry(r));
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
            $"INSERT INTO WorkingEntries ({EntryColumns}) VALUES (@id,@m,@k,@d,@t,@b,@s,@sv,@r,@o,@due,1,@u,@now,@u,@now)", tx,
            ("@id", id), ("@m", model), ("@k", e.Kind), ("@d", e.Date), ("@t", e.Title), ("@b", e.Body), ("@s", e.Status),
            ("@sv", e.Severity), ("@r", e.Result), ("@o", e.Owner), ("@due", e.Due), ("@u", user), ("@now", now)))
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
        using (SqliteCommand upd = Cmd(conn,
            "UPDATE WorkingEntries SET Kind=@k, Date=@d, Title=@t, Body=@b, Status=@s, Severity=@sv, Result=@r, Owner=@o, Due=@due, " +
            "Revision=Revision+1, UpdatedBy=@u, UpdatedAt=@now WHERE Id=@id AND ModelName=@m", tx,
            ("@id", id), ("@m", model), ("@k", e.Kind), ("@d", e.Date), ("@t", e.Title), ("@b", e.Body), ("@s", e.Status),
            ("@sv", e.Severity), ("@r", e.Result), ("@o", e.Owner), ("@due", e.Due), ("@u", user), ("@now", Now())))
            upd.ExecuteNonQuery();
        WorkingEntry saved = ReadEntryRow(conn, model, id, tx)!;
        tx.Commit();
        return saved;
    }

    public bool DeleteEntry(string model, string id)
    {
        model = CleanModelName(model);
        using SqliteConnection conn = Open();
        using SqliteCommand del = Cmd(conn, "DELETE FROM WorkingEntries WHERE Id=@id AND ModelName=@m", null, ("@id", id), ("@m", model));
        return del.ExecuteNonQuery() > 0;
    }
}
