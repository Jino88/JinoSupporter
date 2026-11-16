using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace PortalGNLAB.Web.Services;

/// <summary>A row of the PROCESSING REQUEST: the JIG / equipment asked for. Numbers are quantities (decimals allowed, never negative).</summary>
public sealed class WorkingRequestItem
{
    public int No { get; set; }
    public string Model { get; set; } = "";
    public string Name { get; set; } = "";
    public double LineSpec { get; set; }
    public double Stock { get; set; }
    public double Order { get; set; }
    public string DrawingNo { get; set; } = "";
    public string Picture { get; set; } = "";
    public string Note { get; set; } = "";
}

/// <summary>A JIG/MC processing request as the form sends it: everything but the id and the created / updated stamps. Dates are "" or yyyy-MM-dd.</summary>
public class WorkingRequestSave
{
    public string Kind { get; set; } = WorkingRequestService.DefaultKind;
    public string Requester { get; set; } = "";
    public string Phone { get; set; } = "";
    public string Checker { get; set; } = "";
    public string Confirm { get; set; } = "";
    public string Method { get; set; } = "new";
    public string DepartmentName { get; set; } = WorkingRequestService.DefaultDepartment;
    public string EquipmentName { get; set; } = "";
    public string PartName { get; set; } = "";
    public string ItemApplication { get; set; } = "";
    public string UsingPlace { get; set; } = "";
    public string RequestDate { get; set; } = "";
    public string FinishDate1 { get; set; } = "";
    public string FinishDate2 { get; set; } = "";
    public string Purpose { get; set; } = WorkingRequestService.DefaultPurpose;
    public string Others { get; set; } = "";
    public string DeptCode { get; set; } = "";
    public string Qty1 { get; set; } = "";
    public string Qty2 { get; set; } = "";
    public List<WorkingRequestItem> Items { get; set; } = [];
}

public sealed class WorkingRequest : WorkingRequestSave
{
    public string Id { get; set; } = "";
    public string CreatedBy { get; set; } = "";
    public string CreatedAt { get; set; } = "";
    public string UpdatedBy { get; set; } = "";
    public string UpdatedAt { get; set; } = "";
}

/// <summary>
/// Working > Request: the JIG / MC processing requests, kept as JSON in the WorkingRequests table of the portal DB.
/// Every request is cleaned (<see cref="Normalize"/>) before it is stored.
/// </summary>
public sealed partial class WorkingRequestService(string? dbPath = null)
{
    public const string DefaultKind = "JIG/MC";
    public const string DefaultDepartment = "ME GROUP";
    public const string DefaultPurpose = "New request";
    public const int MaxItems = 500;

    private static readonly string[] Methods = ["new", "repair", "scrap"];
    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web);
    private static readonly object Gate = new();
    private static readonly object SchemaGate = new();
    private static readonly HashSet<string> SchemaReady = [];

    private const string SchemaSql = """
        CREATE TABLE IF NOT EXISTS WorkingRequests (
            Id TEXT PRIMARY KEY,
            Json TEXT NOT NULL,
            CreatedBy TEXT NOT NULL DEFAULT '',
            CreatedAt TEXT NOT NULL,
            UpdatedBy TEXT NOT NULL DEFAULT '',
            UpdatedAt TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS IX_WorkingRequests_UpdatedAt ON WorkingRequests (UpdatedAt);
        """;

    private SqliteConnection Open()
    {
        SqliteConnection conn = dbPath is null ? PortalDb.Open() : PortalDb.Open(dbPath);
        lock (SchemaGate)
        {
            if (SchemaReady.Add(conn.DataSource))
            {
                using SqliteCommand cmd = conn.CreateCommand();
                cmd.CommandText = SchemaSql;
                cmd.ExecuteNonQuery();
            }
        }
        return conn;
    }

    private static string Now() => DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);

    /// <summary>Every request, the most recently updated first.</summary>
    public List<WorkingRequest> List()
    {
        using SqliteConnection conn = Open();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id, Json, CreatedBy, CreatedAt, UpdatedBy, UpdatedAt FROM WorkingRequests ORDER BY UpdatedAt DESC, CreatedAt DESC";
        var list = new List<WorkingRequest>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) list.Add(Read(r));
        return list;
    }

    public WorkingRequest Create(WorkingRequestSave? body, string user)
    {
        WorkingRequestSave save = Normalize(body);
        string now = Now();
        lock (Gate)
        {
            using SqliteConnection conn = Open();
            using SqliteTransaction tx = conn.BeginTransaction();
            if (save.DeptCode.Length == 0) save.DeptCode = NextDeptCode(conn, tx, save.RequestDate);
            string id = Guid.NewGuid().ToString("N");
            using (SqliteCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "INSERT INTO WorkingRequests (Id, Json, CreatedBy, CreatedAt, UpdatedBy, UpdatedAt) VALUES ($id, $json, $by, $at, $by, $at)";
                cmd.Parameters.AddWithValue("$id", id);
                cmd.Parameters.AddWithValue("$json", JsonSerializer.Serialize<WorkingRequestSave>(save, Json));
                cmd.Parameters.AddWithValue("$by", user);
                cmd.Parameters.AddWithValue("$at", now);
                cmd.ExecuteNonQuery();
            }
            tx.Commit();
            return Build(id, save, user, now, user, now);
        }
    }

    /// <summary>Replaces the request; null when there is none with that id. An empty deptCode keeps the stored one.</summary>
    public WorkingRequest? Update(string id, WorkingRequestSave? body, string user)
    {
        WorkingRequestSave save = Normalize(body);
        string now = Now();
        lock (Gate)
        {
            using SqliteConnection conn = Open();
            using SqliteTransaction tx = conn.BeginTransaction();
            WorkingRequest? old = Get(conn, tx, id);
            if (old is null) return null;
            if (save.DeptCode.Length == 0)
                save.DeptCode = old.DeptCode.Length > 0 ? old.DeptCode : NextDeptCode(conn, tx, save.RequestDate);
            using (SqliteCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "UPDATE WorkingRequests SET Json = $json, UpdatedBy = $by, UpdatedAt = $at WHERE Id = $id";
                cmd.Parameters.AddWithValue("$id", id);
                cmd.Parameters.AddWithValue("$json", JsonSerializer.Serialize<WorkingRequestSave>(save, Json));
                cmd.Parameters.AddWithValue("$by", user);
                cmd.Parameters.AddWithValue("$at", now);
                cmd.ExecuteNonQuery();
            }
            tx.Commit();
            return Build(id, save, old.CreatedBy, old.CreatedAt, user, now);
        }
    }

    public bool Delete(string id)
    {
        lock (Gate)
        {
            using SqliteConnection conn = Open();
            using SqliteCommand cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM WorkingRequests WHERE Id = $id";
            cmd.Parameters.AddWithValue("$id", id);
            return cmd.ExecuteNonQuery() > 0;
        }
    }

    private static WorkingRequest? Get(SqliteConnection conn, SqliteTransaction tx, string id)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.Transaction = tx;
        cmd.CommandText = "SELECT Id, Json, CreatedBy, CreatedAt, UpdatedBy, UpdatedAt FROM WorkingRequests WHERE Id = $id";
        cmd.Parameters.AddWithValue("$id", id);
        using SqliteDataReader r = cmd.ExecuteReader();
        return r.Read() ? Read(r) : null;
    }

    private static WorkingRequest Read(SqliteDataReader r)
    {
        WorkingRequestSave save;
        try { save = Normalize(JsonSerializer.Deserialize<WorkingRequestSave>(r.GetString(1), Json)); }
        catch (JsonException) { save = Normalize(null); }
        return Build(r.GetString(0), save, r.GetString(2), r.GetString(3), r.GetString(4), r.GetString(5));
    }

    private static WorkingRequest Build(string id, WorkingRequestSave s, string createdBy, string createdAt, string updatedBy, string updatedAt) => new()
    {
        Id = id, Kind = s.Kind, Requester = s.Requester, Phone = s.Phone, Checker = s.Checker, Confirm = s.Confirm, Method = s.Method,
        DepartmentName = s.DepartmentName, EquipmentName = s.EquipmentName, PartName = s.PartName, ItemApplication = s.ItemApplication,
        UsingPlace = s.UsingPlace, RequestDate = s.RequestDate, FinishDate1 = s.FinishDate1, FinishDate2 = s.FinishDate2, Purpose = s.Purpose,
        Others = s.Others, DeptCode = s.DeptCode, Qty1 = s.Qty1, Qty2 = s.Qty2, Items = s.Items,
        CreatedBy = createdBy, CreatedAt = createdAt, UpdatedBy = updatedBy, UpdatedAt = updatedAt,
    };

    /// <summary>"ME-" + yyMM of the request date (today when none) + the next 3-digit number among the codes of that yyMM.</summary>
    private static string NextDeptCode(SqliteConnection conn, SqliteTransaction tx, string requestDate)
    {
        DateTime date = DateTime.TryParseExact(requestDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime d) ? d : DateTime.Today;
        string prefix = "ME-" + date.ToString("yyMM", CultureInfo.InvariantCulture);
        int max = 0;
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.Transaction = tx;
        cmd.CommandText = "SELECT json_extract(Json, '$.deptCode') FROM WorkingRequests WHERE json_extract(Json, '$.deptCode') LIKE $like";
        cmd.Parameters.AddWithValue("$like", prefix + "%");
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            string code = r.GetString(0);
            if (code.StartsWith(prefix, StringComparison.Ordinal) && code.Length == prefix.Length + 3
                && int.TryParse(code.AsSpan(prefix.Length), NumberStyles.None, CultureInfo.InvariantCulture, out int n) && n > max) max = n;
        }
        return prefix + (max + 1).ToString("000", CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// A cleaned copy: trimmed, length-limited text, defaults for kind / method / department / purpose, dates as yyyy-MM-dd
    /// ("" when unreadable), quantities never negative or NaN, at most <see cref="MaxItems"/> rows, rows without a name dropped,
    /// the rest numbered from 1.
    /// </summary>
    public static WorkingRequestSave Normalize(WorkingRequestSave? s)
    {
        s ??= new WorkingRequestSave();
        string method = (s.Method ?? "").Trim().ToLowerInvariant();
        var items = new List<WorkingRequestItem>();
        foreach (WorkingRequestItem? i in s.Items ?? [])
        {
            if (i is null) continue;
            string name = Text(i.Name, 200);
            if (name.Length == 0) continue;
            items.Add(new WorkingRequestItem
            {
                No = items.Count + 1, Model = Text(i.Model, 100), Name = name, LineSpec = Qty(i.LineSpec), Stock = Qty(i.Stock), Order = Qty(i.Order),
                DrawingNo = Text(i.DrawingNo, 100), Picture = Text(i.Picture, 200), Note = Text(i.Note, 500),
            });
            if (items.Count >= MaxItems) break;
        }
        return new WorkingRequestSave
        {
            Kind = Or(Text(s.Kind, 50), DefaultKind),
            Requester = Text(s.Requester, 100), Phone = Text(s.Phone, 50), Checker = Text(s.Checker, 100), Confirm = Text(s.Confirm, 100),
            Method = Methods.Contains(method) ? method : "new",
            DepartmentName = Or(Text(s.DepartmentName, 100), DefaultDepartment),
            EquipmentName = Text(s.EquipmentName, 200), PartName = Text(s.PartName, 200),
            ItemApplication = Text(s.ItemApplication, 200), UsingPlace = Text(s.UsingPlace, 200),
            RequestDate = DateOf(s.RequestDate), FinishDate1 = DateOf(s.FinishDate1), FinishDate2 = DateOf(s.FinishDate2),
            Purpose = Or(Text(s.Purpose, 200), DefaultPurpose), Others = Text(s.Others, 1000),
            DeptCode = Text(s.DeptCode, 50), Qty1 = Text(s.Qty1, 50), Qty2 = Text(s.Qty2, 50), Items = items,
        };
    }

    private static string Or(string v, string fallback) => v.Length == 0 ? fallback : v;

    public static string Text(string? value, int max)
    {
        string v = (value ?? "").Replace("\r\n", "\n").Trim();
        return v.Length > max ? v[..max].TrimEnd() : v;
    }

    /// <summary>A quantity: decimals kept; negative, NaN or infinite becomes 0.</summary>
    public static double Qty(double v) => double.IsFinite(v) && v > 0 ? v : 0;

    /// <summary>yyyy-MM-dd, yyyy.MM.dd, yyyy/M/d (with a time after it, or without separators as yyyyMMdd) as yyyy-MM-dd; anything else "".</summary>
    public static string DateOf(string? value)
    {
        string v = (value ?? "").Trim();
        int y, mo, d;
        if (DatePattern().Match(v) is { Success: true } m)
        { y = int.Parse(m.Groups[1].Value); mo = int.Parse(m.Groups[2].Value); d = int.Parse(m.Groups[3].Value); }
        else if (Compact().Match(v) is { Success: true } c)
        { y = int.Parse(c.Groups[1].Value); mo = int.Parse(c.Groups[2].Value); d = int.Parse(c.Groups[3].Value); }
        else return "";
        return y is >= 1 and <= 9999 && mo is >= 1 and <= 12 && d >= 1 && d <= DateTime.DaysInMonth(y, mo)
            ? new DateTime(y, mo, d).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : "";
    }

    [GeneratedRegex(@"^(\d{4})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})(?!\d)")]
    private static partial Regex DatePattern();

    [GeneratedRegex(@"^(\d{4})(\d{2})(\d{2})$")]
    private static partial Regex Compact();
}
