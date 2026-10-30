using System.Globalization;
using System.Text.Json;
using ClosedXML.Excel;
using Microsoft.Data.Sqlite;

namespace PortalGNLAB.Web.Services;

public sealed class ProductionConflictException(string message) : Exception(message);

public sealed record ProductionModelSummary(
    string Id, string Name, string Build, int ProcessCount, int MaterialCount, int Revision, string UpdatedAt, string UpdatedBy);

public sealed record ProductionModel(
    string Id, string Name, string Build, JsonElement Setup, int Revision, string UpdatedAt, string UpdatedBy);

public sealed record ProductionDay(
    string ModelId, string Date, string Status, JsonElement Data, JsonElement Summary, int Revision, string EnteredBy, string UpdatedAt);

public sealed record ProductionPlan(JsonElement? Data, int Revision, string UpdatedAt, string UpdatedBy);

public sealed record ProductionDayPair(ProductionDay? Day, ProductionDay? Previous);

public sealed record ProductionModelSave(string? Name, string? Build, JsonElement Setup, int Revision);

public sealed record ProductionDaySave(string? Status, JsonElement Data, JsonElement Summary, int Revision);

public sealed record ProductionPlanSave(JsonElement Data, int Revision);

public sealed record ExportMerge(int Row, int Col, int RowSpan, int ColSpan);

public sealed record ExportSheet(string? Name, List<List<JsonElement>>? Rows, List<ExportMerge>? Merges, int HeaderRows, List<double>? ColWidths);

public sealed record ExportRequest(string? FileName, List<ExportSheet>? Sheets);

/// <summary>
/// Daily production input per model. Setup, day data/summary and plan data are opaque JSON written by the
/// React screen and stored verbatim; the server only counts setup.processes / setup.materials.
/// Tables live in PortalGNLAB.db.
/// </summary>
public sealed class ProductionInputService(string? dbPath = null)
{
    public const string SchemaSql = """
        CREATE TABLE IF NOT EXISTS ProductionModels (
            Id TEXT PRIMARY KEY,
            Name TEXT NOT NULL,
            Build TEXT NOT NULL DEFAULT '',
            SetupJson TEXT NOT NULL DEFAULT 'null',
            Revision INTEGER NOT NULL,
            UpdatedAt TEXT NOT NULL,
            UpdatedBy TEXT NOT NULL DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS ProductionDays (
            ModelId TEXT NOT NULL,
            Date TEXT NOT NULL,
            Status TEXT NOT NULL,
            DataJson TEXT NOT NULL DEFAULT 'null',
            SummaryJson TEXT NOT NULL DEFAULT 'null',
            Revision INTEGER NOT NULL,
            EnteredBy TEXT NOT NULL DEFAULT '',
            UpdatedAt TEXT NOT NULL,
            PRIMARY KEY (ModelId, Date)
        );
        CREATE TABLE IF NOT EXISTS ProductionPlans (
            ModelId TEXT PRIMARY KEY,
            DataJson TEXT NOT NULL DEFAULT 'null',
            Revision INTEGER NOT NULL,
            UpdatedAt TEXT NOT NULL,
            UpdatedBy TEXT NOT NULL DEFAULT ''
        );
        """;

    public static readonly string[] Statuses = ["Draft", "Submitted"];
    public const int MaxRangeDays = 400;
    public const int MaxSheets = 20;
    public const int MaxCells = 200_000;

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

    private static string RawJson(JsonElement e) => e.ValueKind is JsonValueKind.Undefined ? "null" : e.GetRawText();

    private static JsonElement ParseJson(string text)
    {
        using JsonDocument doc = JsonDocument.Parse(string.IsNullOrWhiteSpace(text) ? "null" : text);
        return doc.RootElement.Clone();
    }

    private static int ArrayLength(JsonElement setup, string name) =>
        setup.ValueKind == JsonValueKind.Object && setup.TryGetProperty(name, out JsonElement a) && a.ValueKind == JsonValueKind.Array
            ? a.GetArrayLength() : 0;

    public static DateOnly ParseDate(string? date)
    {
        if (!DateOnly.TryParseExact(date, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOnly d))
            throw new ArgumentException("Date must be yyyy-MM-dd.");
        return d;
    }

    private static (string Name, string Build) CleanNames(ProductionModelSave body)
    {
        string name = (body.Name ?? "").Trim();
        string build = (body.Build ?? "").Trim();
        if (name.Length == 0) throw new ArgumentException("Model name is required.");
        if (name.Length > 80) throw new ArgumentException("Model name must be 80 characters or fewer.");
        if (build.Length > 40) throw new ArgumentException("Build must be 40 characters or fewer.");
        return (name, build);
    }

    // ── Models ──────────────────────────────────────────────────────────────

    public List<ProductionModelSummary> ListModels()
    {
        using SqliteConnection conn = Open();
        using SqliteCommand cmd = Cmd(conn,
            "SELECT Id, Name, Build, SetupJson, Revision, UpdatedAt, UpdatedBy FROM ProductionModels ORDER BY Name COLLATE NOCASE, Build COLLATE NOCASE");
        var list = new List<ProductionModelSummary>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            JsonElement setup = ParseJson(r.GetString(3));
            list.Add(new(r.GetString(0), r.GetString(1), r.GetString(2), ArrayLength(setup, "processes"),
                ArrayLength(setup, "materials"), r.GetInt32(4), r.GetString(5), r.GetString(6)));
        }
        return list;
    }

    private static ProductionModel? ReadModel(SqliteConnection conn, string id, SqliteTransaction? tx = null)
    {
        using SqliteCommand cmd = Cmd(conn,
            "SELECT Id, Name, Build, SetupJson, Revision, UpdatedAt, UpdatedBy FROM ProductionModels WHERE Id=@id", tx, ("@id", id));
        using SqliteDataReader r = cmd.ExecuteReader();
        return r.Read()
            ? new(r.GetString(0), r.GetString(1), r.GetString(2), ParseJson(r.GetString(3)), r.GetInt32(4), r.GetString(5), r.GetString(6))
            : null;
    }

    public ProductionModel? GetModel(string id)
    {
        using SqliteConnection conn = Open();
        return ReadModel(conn, id);
    }

    private static bool Duplicate(SqliteConnection conn, SqliteTransaction tx, string name, string build, string? exceptId)
    {
        using SqliteCommand cmd = Cmd(conn,
            "SELECT 1 FROM ProductionModels WHERE lower(Name)=lower(@n) AND lower(Build)=lower(@b) AND Id<>@id LIMIT 1", tx,
            ("@n", name), ("@b", build), ("@id", exceptId ?? ""));
        return cmd.ExecuteScalar() is not null;
    }

    public ProductionModel CreateModel(ProductionModelSave body, string user)
    {
        (string name, string build) = CleanNames(body);
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        if (Duplicate(conn, tx, name, build, null)) throw new ProductionConflictException("A model with the same name and build already exists.");
        string id = Guid.NewGuid().ToString("N");
        using (SqliteCommand ins = Cmd(conn,
            "INSERT INTO ProductionModels (Id, Name, Build, SetupJson, Revision, UpdatedAt, UpdatedBy) VALUES (@id,@n,@b,@s,1,@t,@u)", tx,
            ("@id", id), ("@n", name), ("@b", build), ("@s", RawJson(body.Setup)), ("@t", Now()), ("@u", user)))
            ins.ExecuteNonQuery();
        ProductionModel saved = ReadModel(conn, id, tx)!;
        tx.Commit();
        return saved;
    }

    /// <summary>Null when the model does not exist.</summary>
    public ProductionModel? UpdateModel(string id, ProductionModelSave body, string user)
    {
        (string name, string build) = CleanNames(body);
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        ProductionModel? current = ReadModel(conn, id, tx);
        if (current is null) return null;
        if (current.Revision != body.Revision)
            throw new ProductionConflictException("The model was changed by someone else. Reload it and apply your change again.");
        if (Duplicate(conn, tx, name, build, id)) throw new ProductionConflictException("A model with the same name and build already exists.");
        using (SqliteCommand upd = Cmd(conn,
            "UPDATE ProductionModels SET Name=@n, Build=@b, SetupJson=@s, Revision=Revision+1, UpdatedAt=@t, UpdatedBy=@u WHERE Id=@id", tx,
            ("@id", id), ("@n", name), ("@b", build), ("@s", RawJson(body.Setup)), ("@t", Now()), ("@u", user)))
            upd.ExecuteNonQuery();
        ProductionModel saved = ReadModel(conn, id, tx)!;
        tx.Commit();
        return saved;
    }

    public bool DeleteModel(string id)
    {
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        foreach (string table in new[] { "ProductionDays", "ProductionPlans" })
            using (SqliteCommand del = Cmd(conn, $"DELETE FROM {table} WHERE ModelId=@id", tx, ("@id", id)))
                del.ExecuteNonQuery();
        int n;
        using (SqliteCommand del = Cmd(conn, "DELETE FROM ProductionModels WHERE Id=@id", tx, ("@id", id)))
            n = del.ExecuteNonQuery();
        tx.Commit();
        return n > 0;
    }

    // ── Days ────────────────────────────────────────────────────────────────

    private const string DayColumns = "ModelId, Date, Status, DataJson, SummaryJson, Revision, EnteredBy, UpdatedAt";

    private static ProductionDay ReadDay(SqliteDataReader r) =>
        new(r.GetString(0), r.GetString(1), r.GetString(2), ParseJson(r.GetString(3)), ParseJson(r.GetString(4)),
            r.GetInt32(5), r.GetString(6), r.GetString(7));

    private static ProductionDay? ReadDayRow(SqliteConnection conn, string modelId, string date, SqliteTransaction? tx = null)
    {
        using SqliteCommand cmd = Cmd(conn, $"SELECT {DayColumns} FROM ProductionDays WHERE ModelId=@m AND Date=@d", tx, ("@m", modelId), ("@d", date));
        using SqliteDataReader r = cmd.ExecuteReader();
        return r.Read() ? ReadDay(r) : null;
    }

    public ProductionDayPair GetDay(string modelId, string date)
    {
        ParseDate(date);
        using SqliteConnection conn = Open();
        ProductionDay? day = ReadDayRow(conn, modelId, date);
        ProductionDay? previous;
        using (SqliteCommand cmd = Cmd(conn,
            $"SELECT {DayColumns} FROM ProductionDays WHERE ModelId=@m AND Date<@d ORDER BY Date DESC LIMIT 1", null, ("@m", modelId), ("@d", date)))
        using (SqliteDataReader r = cmd.ExecuteReader())
            previous = r.Read() ? ReadDay(r) : null;
        return new(day, previous);
    }

    /// <summary>Null when the model does not exist.</summary>
    public ProductionDay? SaveDay(string modelId, string date, ProductionDaySave body, string user)
    {
        ParseDate(date);
        if (body.Status is null || !Statuses.Contains(body.Status, StringComparer.Ordinal))
            throw new ArgumentException("Status must be Draft or Submitted.");
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        if (ReadModel(conn, modelId, tx) is null) return null;
        ProductionDay? current = ReadDayRow(conn, modelId, date, tx);
        string now = Now();
        if (body.Revision == 0)
        {
            if (current is not null) throw new ProductionConflictException("This day already has an entry. Reload it and apply your change again.");
            using SqliteCommand ins = Cmd(conn,
                $"INSERT INTO ProductionDays ({DayColumns}) VALUES (@m,@d,@s,@data,@sum,1,@u,@t)", tx,
                ("@m", modelId), ("@d", date), ("@s", body.Status), ("@data", RawJson(body.Data)), ("@sum", RawJson(body.Summary)), ("@u", user), ("@t", now));
            ins.ExecuteNonQuery();
        }
        else
        {
            if (current is null || current.Revision != body.Revision)
                throw new ProductionConflictException("This day was changed by someone else. Reload it and apply your change again.");
            using SqliteCommand upd = Cmd(conn,
                "UPDATE ProductionDays SET Status=@s, DataJson=@data, SummaryJson=@sum, Revision=Revision+1, EnteredBy=@u, UpdatedAt=@t WHERE ModelId=@m AND Date=@d", tx,
                ("@m", modelId), ("@d", date), ("@s", body.Status), ("@data", RawJson(body.Data)), ("@sum", RawJson(body.Summary)), ("@u", user), ("@t", now));
            upd.ExecuteNonQuery();
        }
        ProductionDay saved = ReadDayRow(conn, modelId, date, tx)!;
        tx.Commit();
        return saved;
    }

    public void DeleteDay(string modelId, string date)
    {
        ParseDate(date);
        using SqliteConnection conn = Open();
        using SqliteCommand del = Cmd(conn, "DELETE FROM ProductionDays WHERE ModelId=@m AND Date=@d", null, ("@m", modelId), ("@d", date));
        del.ExecuteNonQuery();
    }

    public List<ProductionDay> ListDays(string modelId, string? from, string? to)
    {
        if (string.IsNullOrEmpty(from) || string.IsNullOrEmpty(to)) throw new ArgumentException("Both from and to are required.");
        DateOnly f = ParseDate(from), t = ParseDate(to);
        if (t < f) throw new ArgumentException("to must not be before from.");
        if (t.DayNumber - f.DayNumber > MaxRangeDays) throw new ArgumentException($"The range must not exceed {MaxRangeDays} days.");
        using SqliteConnection conn = Open();
        using SqliteCommand cmd = Cmd(conn,
            $"SELECT {DayColumns} FROM ProductionDays WHERE ModelId=@m AND Date>=@f AND Date<=@t ORDER BY Date", null,
            ("@m", modelId), ("@f", from), ("@t", to));
        var list = new List<ProductionDay>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) list.Add(ReadDay(r));
        return list;
    }

    // ── Plan ────────────────────────────────────────────────────────────────

    private static ProductionPlan? ReadPlan(SqliteConnection conn, string modelId, SqliteTransaction? tx = null)
    {
        using SqliteCommand cmd = Cmd(conn, "SELECT DataJson, Revision, UpdatedAt, UpdatedBy FROM ProductionPlans WHERE ModelId=@m", tx, ("@m", modelId));
        using SqliteDataReader r = cmd.ExecuteReader();
        return r.Read() ? new(ParseJson(r.GetString(0)), r.GetInt32(1), r.GetString(2), r.GetString(3)) : null;
    }

    public ProductionPlan GetPlan(string modelId)
    {
        using SqliteConnection conn = Open();
        return ReadPlan(conn, modelId) ?? new(null, 0, "", "");
    }

    /// <summary>Null when the model does not exist.</summary>
    public ProductionPlan? SavePlan(string modelId, ProductionPlanSave body, string user)
    {
        using SqliteConnection conn = Open();
        using SqliteTransaction tx = conn.BeginTransaction();
        if (ReadModel(conn, modelId, tx) is null) return null;
        ProductionPlan? current = ReadPlan(conn, modelId, tx);
        if ((current?.Revision ?? 0) != body.Revision)
            throw new ProductionConflictException("The plan was changed by someone else. Reload it and apply your change again.");
        string sql = current is null
            ? "INSERT INTO ProductionPlans (ModelId, DataJson, Revision, UpdatedAt, UpdatedBy) VALUES (@m,@data,1,@t,@u)"
            : "UPDATE ProductionPlans SET DataJson=@data, Revision=Revision+1, UpdatedAt=@t, UpdatedBy=@u WHERE ModelId=@m";
        using (SqliteCommand cmd = Cmd(conn, sql, tx, ("@m", modelId), ("@data", RawJson(body.Data)), ("@t", Now()), ("@u", user)))
            cmd.ExecuteNonQuery();
        ProductionPlan saved = ReadPlan(conn, modelId, tx)!;
        tx.Commit();
        return saved;
    }

    // ── Excel export ────────────────────────────────────────────────────────

    public static string SafeFileName(string? fileName)
    {
        string name = Path.GetFileName((fileName ?? "").Trim());
        var bad = Path.GetInvalidFileNameChars().Concat(['"', ';', ',']).ToHashSet();
        name = new string(name.Select(c => bad.Contains(c) || char.IsControl(c) ? '_' : c).ToArray()).Trim(' ', '.');
        if (name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) name = name[..^5].TrimEnd(' ', '.');
        if (name.Length == 0) name = "production-input";
        if (name.Length > 120) name = name[..120];
        return name + ".xlsx";
    }

    private static string SheetName(string? raw, int index, HashSet<string> used)
    {
        string name = new((raw ?? "").Where(c => !"[]:*?/\\".Contains(c) && !char.IsControl(c)).ToArray());
        name = name.Trim().Trim('\'');
        if (name.Length == 0) name = $"Sheet{index + 1}";
        if (name.Length > 31) name = name[..31];
        string candidate = name;
        for (int n = 2; !used.Add(candidate.ToLowerInvariant()); n++)
        {
            string suffix = $" ({n})";
            candidate = name[..Math.Min(name.Length, 31 - suffix.Length)] + suffix;
        }
        return candidate;
    }

    public static byte[] ExportXlsx(ExportRequest request)
    {
        List<ExportSheet> sheets = request.Sheets ?? [];
        if (sheets.Count == 0) throw new ArgumentException("At least one sheet is required.");
        if (sheets.Count > MaxSheets) throw new ArgumentException($"At most {MaxSheets} sheets can be exported.");
        long cells = sheets.Sum(s => (s.Rows ?? []).Sum(r => (long)r.Count));
        if (cells > MaxCells) throw new ArgumentException($"At most {MaxCells} cells can be exported.");

        using var book = new XLWorkbook();
        var used = new HashSet<string>();
        for (int si = 0; si < sheets.Count; si++)
        {
            ExportSheet sheet = sheets[si];
            IXLWorksheet ws = book.Worksheets.Add(SheetName(sheet.Name, si, used));
            List<List<JsonElement>> rows = sheet.Rows ?? [];
            int maxCol = 0;
            for (int r = 0; r < rows.Count; r++)
            {
                maxCol = Math.Max(maxCol, rows[r].Count);
                for (int c = 0; c < rows[r].Count; c++)
                {
                    JsonElement v = rows[r][c];
                    IXLCell cell = ws.Cell(r + 1, c + 1);
                    if (v.ValueKind == JsonValueKind.Number && v.TryGetDouble(out double d)) cell.Value = d;
                    else if (v.ValueKind == JsonValueKind.String) cell.Value = v.GetString() ?? "";
                    else if (v.ValueKind is JsonValueKind.True or JsonValueKind.False) cell.Value = v.GetBoolean() ? "TRUE" : "FALSE";
                }
            }
            if (rows.Count > 0 && maxCol > 0)
            {
                IXLRange usedRange = ws.Range(1, 1, rows.Count, maxCol);
                usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                int header = Math.Clamp(sheet.HeaderRows, 0, rows.Count);
                if (header > 0)
                {
                    IXLRange h = ws.Range(1, 1, header, maxCol);
                    h.Style.Font.Bold = true;
                    h.Style.Fill.BackgroundColor = XLColor.FromHtml("#E7E6E6");
                    h.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    h.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    ws.SheetView.FreezeRows(header);
                }
            }
            foreach (ExportMerge m in sheet.Merges ?? [])
            {
                if (m.Row < 0 || m.Col < 0 || m.RowSpan < 1 || m.ColSpan < 1 || (m.RowSpan == 1 && m.ColSpan == 1)) continue;
                if (m.Row + m.RowSpan > rows.Count || m.Col + m.ColSpan > maxCol) continue;
                ws.Range(m.Row + 1, m.Col + 1, m.Row + m.RowSpan, m.Col + m.ColSpan).Merge();
            }
            if (sheet.ColWidths is { } widths)
                for (int c = 0; c < widths.Count && c < 200; c++)
                    if (widths[c] > 0) ws.Column(c + 1).Width = Math.Min(widths[c], 255);
        }
        using var ms = new MemoryStream();
        book.SaveAs(ms);
        return ms.ToArray();
    }
}
