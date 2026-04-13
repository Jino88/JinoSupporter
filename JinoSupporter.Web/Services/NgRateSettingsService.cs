using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

/// <summary>
/// NG Rate 전용 설정 DB 서비스.
/// 설정 파일: D:\000. MyWorks\000. 일일업무\04. DB\01. NGRATE\ModelBmes\ngrate_settings.db
/// </summary>
public sealed class NgRateSettingsService
{
    // ── 기본값 ──────────────────────────────────────────────────────────────────
    public static readonly string DefaultDbSaveDirectory =
        @"D:\000. MyWorks\000. 일일업무\04. DB\01. NGRATE";

    public static readonly string DefaultRoutingFilePath =
        @"D:\000. MyWorks\000. 일일업무\04. DB\01. NGRATE\Routing.txt";

    public static readonly string DefaultReasonFilePath =
        @"D:\000. MyWorks\000. 일일업무\04. DB\01. NGRATE\reason.txt";

    // ── Settings DB 위치 ─────────────────────────────────────────────────────────
    public static readonly string SettingsDbDirectory =
        @"D:\000. MyWorks\000. 일일업무\04. DB\01. NGRATE\ModelBmes";

    public static readonly string SettingsDbPath =
        Path.Combine(SettingsDbDirectory, "ngrate_settings.db");

    // ── Setting keys ─────────────────────────────────────────────────────────────
    public const string KeyDbSaveDirectory  = "NgRate:DbSaveDirectory";
    public const string KeyRoutingFilePath  = "NgRate:RoutingFilePath";
    public const string KeyReasonFilePath   = "NgRate:ReasonFilePath";
    public const string KeyLoginId          = "NgRate:LoginId";
    public const string KeyPassword         = "NgRate:Password";

    // ── Constructor ──────────────────────────────────────────────────────────────
    public NgRateSettingsService()
    {
        EnsureDatabase();
    }

    // ── Typed accessors ──────────────────────────────────────────────────────────

    public string DbSaveDirectory =>
        GetSetting(KeyDbSaveDirectory) ?? DefaultDbSaveDirectory;

    public string RoutingFilePath =>
        GetSetting(KeyRoutingFilePath) ?? DefaultRoutingFilePath;

    public string ReasonFilePath =>
        GetSetting(KeyReasonFilePath) ?? DefaultReasonFilePath;

    public string LoginId =>
        GetSetting(KeyLoginId) ?? string.Empty;

    public string Password =>
        GetSetting(KeyPassword) ?? string.Empty;

    public bool IsCredentialsConfigured =>
        !string.IsNullOrWhiteSpace(LoginId) && !string.IsNullOrWhiteSpace(Password);

    // ── CRUD ─────────────────────────────────────────────────────────────────────

    public string? GetSetting(string key)
    {
        using var conn = Open();
        using var cmd  = conn.CreateCommand();
        cmd.CommandText = "SELECT Value FROM NgRateSettings WHERE Key = @k;";
        cmd.Parameters.AddWithValue("@k", key);
        var result = cmd.ExecuteScalar();
        return result is string s && s.Length > 0 ? s : null;
    }

    public void SetSetting(string key, string value)
    {
        using var conn = Open();
        using var cmd  = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO NgRateSettings (Key, Value) VALUES (@k, @v)
            ON CONFLICT(Key) DO UPDATE SET Value = excluded.Value;
            """;
        cmd.Parameters.AddWithValue("@k", key);
        cmd.Parameters.AddWithValue("@v", value ?? string.Empty);
        cmd.ExecuteNonQuery();
    }

    /// <summary>현재 저장된 전체 설정을 반환 (UI 바인딩용)</summary>
    public NgRateSettingsSnapshot GetSnapshot() => new()
    {
        DbSaveDirectory = DbSaveDirectory,
        RoutingFilePath = RoutingFilePath,
        ReasonFilePath  = ReasonFilePath,
        LoginId         = LoginId,
        Password        = Password,
    };

    public void ApplySnapshot(NgRateSettingsSnapshot snap)
    {
        SetSetting(KeyDbSaveDirectory, snap.DbSaveDirectory.Trim());
        SetSetting(KeyRoutingFilePath, snap.RoutingFilePath.Trim());
        SetSetting(KeyReasonFilePath,  snap.ReasonFilePath.Trim());
        SetSetting(KeyLoginId,         snap.LoginId.Trim());
        SetSetting(KeyPassword,        snap.Password);
    }

    // ── Private ──────────────────────────────────────────────────────────────────

    private static SqliteConnection Open()
    {
        var conn = new SqliteConnection($"Data Source={SettingsDbPath}");
        conn.Open();
        return conn;
    }

    // ── Routing Table ─────────────────────────────────────────────────────────────

    public List<RoutingRow> GetRoutingRows()
    {
        using var conn = Open();
        using var cmd  = conn.CreateCommand();
        cmd.CommandText = "SELECT Id, ModelName, ProcessCode, ProcessName, ProcessType FROM RoutingTable ORDER BY Id;";
        using var rdr = cmd.ExecuteReader();
        var list = new List<RoutingRow>();
        while (rdr.Read())
            list.Add(new RoutingRow { Id = rdr.GetInt64(0), ModelName = rdr.GetString(1),
                ProcessCode = rdr.GetString(2), ProcessName = rdr.GetString(3), ProcessType = rdr.GetString(4) });
        return list;
    }

    public long AddRoutingRow(RoutingRow r)
    {
        using var conn = Open();
        using var cmd  = conn.CreateCommand();
        cmd.CommandText = "INSERT INTO RoutingTable (ModelName,ProcessCode,ProcessName,ProcessType) VALUES (@m,@pc,@pn,@pt); SELECT last_insert_rowid();";
        cmd.Parameters.AddWithValue("@m",  r.ModelName);
        cmd.Parameters.AddWithValue("@pc", r.ProcessCode);
        cmd.Parameters.AddWithValue("@pn", r.ProcessName);
        cmd.Parameters.AddWithValue("@pt", r.ProcessType);
        return Convert.ToInt64(cmd.ExecuteScalar() ?? 0L);
    }

    public void UpdateRoutingRow(RoutingRow r)
    {
        using var conn = Open();
        using var cmd  = conn.CreateCommand();
        cmd.CommandText = "UPDATE RoutingTable SET ModelName=@m, ProcessCode=@pc, ProcessName=@pn, ProcessType=@pt WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@m",  r.ModelName);
        cmd.Parameters.AddWithValue("@pc", r.ProcessCode);
        cmd.Parameters.AddWithValue("@pn", r.ProcessName);
        cmd.Parameters.AddWithValue("@pt", r.ProcessType);
        cmd.Parameters.AddWithValue("@id", r.Id);
        cmd.ExecuteNonQuery();
    }

    public void DeleteRoutingRow(long id)
    {
        using var conn = Open();
        using var cmd  = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM RoutingTable WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    public int ImportRoutingFromFile(string filePath)
    {
        if (!File.Exists(filePath)) return 0;
        var lines = File.ReadAllLines(filePath, System.Text.Encoding.UTF8);
        if (lines.Length < 2) return 0;
        using var conn = Open();
        using var tx   = conn.BeginTransaction();
        using (var del = conn.CreateCommand()) { del.CommandText = "DELETE FROM RoutingTable;"; del.ExecuteNonQuery(); }
        int count = 0;
        for (int i = 1; i < lines.Length; i++)
        {
            var cols = lines[i].Split('\t');
            if (cols.Length < 4) continue;
            using var ins = conn.CreateCommand();
            ins.CommandText = "INSERT INTO RoutingTable (ModelName,ProcessCode,ProcessName,ProcessType) VALUES (@m,@pc,@pn,@pt);";
            ins.Parameters.AddWithValue("@m",  cols[0].Trim());
            ins.Parameters.AddWithValue("@pc", cols[1].Trim());
            ins.Parameters.AddWithValue("@pn", cols[2].Trim());
            ins.Parameters.AddWithValue("@pt", cols[3].Trim());
            ins.ExecuteNonQuery();
            count++;
        }
        tx.Commit();
        return count;
    }

    // ── Reason Table ──────────────────────────────────────────────────────────────

    public List<ReasonRow> GetReasonRows()
    {
        using var conn = Open();
        using var cmd  = conn.CreateCommand();
        cmd.CommandText = "SELECT Id, ProcessName, NgName, Reason FROM ReasonTable ORDER BY Id;";
        using var rdr = cmd.ExecuteReader();
        var list = new List<ReasonRow>();
        while (rdr.Read())
            list.Add(new ReasonRow { Id = rdr.GetInt64(0), ProcessName = rdr.GetString(1),
                NgName = rdr.GetString(2), Reason = rdr.GetString(3) });
        return list;
    }

    public long AddReasonRow(ReasonRow r)
    {
        using var conn = Open();
        using var cmd  = conn.CreateCommand();
        cmd.CommandText = "INSERT INTO ReasonTable (ProcessName,NgName,Reason) VALUES (@pn,@ng,@rs); SELECT last_insert_rowid();";
        cmd.Parameters.AddWithValue("@pn", r.ProcessName);
        cmd.Parameters.AddWithValue("@ng", r.NgName);
        cmd.Parameters.AddWithValue("@rs", r.Reason);
        return Convert.ToInt64(cmd.ExecuteScalar() ?? 0L);
    }

    public void UpdateReasonRow(ReasonRow r)
    {
        using var conn = Open();
        using var cmd  = conn.CreateCommand();
        cmd.CommandText = "UPDATE ReasonTable SET ProcessName=@pn, NgName=@ng, Reason=@rs WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@pn", r.ProcessName);
        cmd.Parameters.AddWithValue("@ng", r.NgName);
        cmd.Parameters.AddWithValue("@rs", r.Reason);
        cmd.Parameters.AddWithValue("@id", r.Id);
        cmd.ExecuteNonQuery();
    }

    public void DeleteReasonRow(long id)
    {
        using var conn = Open();
        using var cmd  = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM ReasonTable WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    public int ImportReasonFromFile(string filePath)
    {
        if (!File.Exists(filePath)) return 0;
        var lines = File.ReadAllLines(filePath, System.Text.Encoding.UTF8);
        if (lines.Length < 2) return 0;
        using var conn = Open();
        using var tx   = conn.BeginTransaction();
        using (var del = conn.CreateCommand()) { del.CommandText = "DELETE FROM ReasonTable;"; del.ExecuteNonQuery(); }
        int count = 0;
        for (int i = 1; i < lines.Length; i++)
        {
            var cols = lines[i].Split('\t');
            if (cols.Length < 3) continue;
            using var ins = conn.CreateCommand();
            ins.CommandText = "INSERT INTO ReasonTable (ProcessName,NgName,Reason) VALUES (@pn,@ng,@rs);";
            ins.Parameters.AddWithValue("@pn", cols[0].Trim());
            ins.Parameters.AddWithValue("@ng", cols[1].Trim());
            ins.Parameters.AddWithValue("@rs", cols[2].Trim());
            ins.ExecuteNonQuery();
            count++;
        }
        tx.Commit();
        return count;
    }

    // ── Private ──────────────────────────────────────────────────────────────────

    private static void EnsureDatabase()
    {
        Directory.CreateDirectory(SettingsDbDirectory);
        using var conn = new SqliteConnection($"Data Source={SettingsDbPath}");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            CREATE TABLE IF NOT EXISTS NgRateSettings (
                Key   TEXT PRIMARY KEY NOT NULL,
                Value TEXT NOT NULL DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS RoutingTable (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                ModelName   TEXT NOT NULL DEFAULT '',
                ProcessCode TEXT NOT NULL DEFAULT '',
                ProcessName TEXT NOT NULL DEFAULT '',
                ProcessType TEXT NOT NULL DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS ReasonTable (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                ProcessName TEXT NOT NULL DEFAULT '',
                NgName      TEXT NOT NULL DEFAULT '',
                Reason      TEXT NOT NULL DEFAULT ''
            );
            """;
        cmd.ExecuteNonQuery();
    }
}

public sealed class RoutingRow
{
    public long   Id          { get; init; }
    public string ModelName   { get; set; } = "";
    public string ProcessCode { get; set; } = "";
    public string ProcessName { get; set; } = "";
    public string ProcessType { get; set; } = "";
}

public sealed class ReasonRow
{
    public long   Id          { get; init; }
    public string ProcessName { get; set; } = "";
    public string NgName      { get; set; } = "";
    public string Reason      { get; set; } = "";
}

/// <summary>UI 바인딩용 스냅샷 (mutable)</summary>
public sealed class NgRateSettingsSnapshot
{
    public string DbSaveDirectory { get; set; } = NgRateSettingsService.DefaultDbSaveDirectory;
    public string RoutingFilePath { get; set; } = NgRateSettingsService.DefaultRoutingFilePath;
    public string ReasonFilePath  { get; set; } = NgRateSettingsService.DefaultReasonFilePath;
    public string LoginId         { get; set; } = string.Empty;
    public string Password        { get; set; } = string.Empty;
}
