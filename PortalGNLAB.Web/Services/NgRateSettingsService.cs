using Microsoft.Data.Sqlite;
using System.Text.RegularExpressions;

namespace PortalGNLAB.Web.Services;

/// <summary>
/// NG Rate settings service (BMES web / API logins, routing and reason tables).
/// Everything lives in the unified <see cref="PortalDb"/> file; the old ngrate_settings.db
/// is only read by <see cref="PortalDbMigrator"/>.
/// </summary>
public sealed class NgRateSettingsService
{
    private readonly AppPathsService _appPaths;
    private static readonly string BootstrapSettingsDbDirectory =
        AppStoragePaths.Combine("standalone-bootstrap", "ModelBmes");

    // ── Defaults (sourced from AppPathsService → webapp-paths.json) ─────────────
    public string DefaultDbSaveDirectory     => _appPaths.Current.NgRateDbSaveDirectory;
    public string DefaultFCostDbSaveDirectory =>
        HasUnifiedDbDirectory
            ? AppPathsService.GetFCostDirectory(_appPaths.Current.DbRootDirectory)
            : DbSaveDirectory;
    public string DefaultRoutingFilePath     => _appPaths.Current.NgRateRoutingFilePath;
    public string DefaultReasonFilePath      => _appPaths.Current.NgRateReasonFilePath;
    public string WorkerStatusCacheDirectory =>
        HasUnifiedDbDirectory
            ? AppPathsService.GetWorkerStatusDirectory(_appPaths.Current.DbRootDirectory)
            : Path.Combine(DbSaveDirectory, "3. Wk Status");

    // ── Settings DB location ─────────────────────────────────────────────────────
    public string SettingsDbDirectory =>
        string.IsNullOrWhiteSpace(_appPaths.Current.NgRateSettingsDbDirectory)
            ? BootstrapSettingsDbDirectory
            : _appPaths.Current.NgRateSettingsDbDirectory;
    public string SettingsDbPath      => PortalDb.FilePath;
    public bool IsNgRateStorageConfigured => _appPaths.IsNgRateStorageConfigured;
    private bool HasUnifiedDbDirectory =>
        !string.IsNullOrWhiteSpace(_appPaths.Current.DbRootDirectory);

    // ── Setting keys ─────────────────────────────────────────────────────────────
    public const string KeyDbSaveDirectory  = "NgRate:DbSaveDirectory";
    public const string KeyFCostDbSaveDirectory = "FCost:DbSaveDirectory";
    public const string KeyRoutingFilePath  = "NgRate:RoutingFilePath";
    public const string KeyReasonFilePath   = "NgRate:ReasonFilePath";
    public const string KeyLoginId          = "NgRate:LoginId";
    public const string KeyPassword         = "NgRate:Password";
    // Request Work Time-specific credentials
    public const string KeyRwtLoginId       = "Rwt:LoginId";
    public const string KeyRwtPassword      = "Rwt:Password";
    // BMES PLATFORM_API (https://api.bujeon.com) login: facility code (BN / GN), user id, password
    public const string KeyApiFacility      = "BmesApi:Facility";
    public const string KeyApiLoginId       = "BmesApi:LoginId";
    public const string KeyApiPassword      = "BmesApi:Password";

    // ── Constructor ──────────────────────────────────────────────────────────────
    public NgRateSettingsService(AppPathsService appPaths)
    {
        _appPaths = appPaths;
        EnsureDatabase();
    }

    // ── Typed accessors ──────────────────────────────────────────────────────────

    public string DbSaveDirectory =>
        HasUnifiedDbDirectory ? DefaultDbSaveDirectory : GetSetting(KeyDbSaveDirectory) ?? DefaultDbSaveDirectory;

    public string FCostDbSaveDirectory =>
        HasUnifiedDbDirectory ? DefaultFCostDbSaveDirectory : GetSetting(KeyFCostDbSaveDirectory) ?? DefaultFCostDbSaveDirectory;

    public string RoutingFilePath =>
        HasUnifiedDbDirectory ? DefaultRoutingFilePath : GetSetting(KeyRoutingFilePath) ?? DefaultRoutingFilePath;

    public string ReasonFilePath =>
        HasUnifiedDbDirectory ? DefaultReasonFilePath : GetSetting(KeyReasonFilePath) ?? DefaultReasonFilePath;

    public string LoginId =>
        GetSetting(KeyLoginId) ?? string.Empty;

    public string Password =>
        GetSetting(KeyPassword) ?? string.Empty;

    public bool IsCredentialsConfigured =>
        !string.IsNullOrWhiteSpace(LoginId) && !string.IsNullOrWhiteSpace(Password);

    public string RwtLoginId =>
        GetSetting(KeyRwtLoginId) ?? string.Empty;

    public string RwtPassword =>
        GetSetting(KeyRwtPassword) ?? string.Empty;

    public bool IsRwtCredentialsConfigured =>
        !string.IsNullOrWhiteSpace(RwtLoginId) && !string.IsNullOrWhiteSpace(RwtPassword);

    public string ApiFacility =>
        GetSetting(KeyApiFacility) ?? string.Empty;

    public string ApiLoginId =>
        GetSetting(KeyApiLoginId) ?? string.Empty;

    public string ApiPassword =>
        GetSetting(KeyApiPassword) ?? string.Empty;

    public bool IsApiCredentialsConfigured =>
        !string.IsNullOrWhiteSpace(ApiFacility) && !string.IsNullOrWhiteSpace(ApiLoginId) && !string.IsNullOrWhiteSpace(ApiPassword);

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

    /// <summary>Return the full saved settings (for UI binding)</summary>
    public NgRateSettingsSnapshot GetSnapshot() => new()
    {
        DbSaveDirectory      = DbSaveDirectory,
        FCostDbSaveDirectory = FCostDbSaveDirectory,
        RoutingFilePath      = RoutingFilePath,
        ReasonFilePath       = ReasonFilePath,
        LoginId              = LoginId,
        Password             = Password,
        RwtLoginId           = RwtLoginId,
        RwtPassword          = RwtPassword,
        ApiFacility          = ApiFacility,
        ApiLoginId           = ApiLoginId,
        ApiPassword          = ApiPassword,
    };

    public void ApplySnapshot(NgRateSettingsSnapshot snap)
    {
        if (HasUnifiedDbDirectory)
        {
            SetSetting(KeyDbSaveDirectory,      DefaultDbSaveDirectory);
            SetSetting(KeyFCostDbSaveDirectory, DefaultFCostDbSaveDirectory);
            SetSetting(KeyRoutingFilePath,      DefaultRoutingFilePath);
            SetSetting(KeyReasonFilePath,       DefaultReasonFilePath);
        }
        else
        {
            SetSetting(KeyDbSaveDirectory,      snap.DbSaveDirectory.Trim());
            SetSetting(KeyFCostDbSaveDirectory, snap.FCostDbSaveDirectory.Trim());
            SetSetting(KeyRoutingFilePath,      snap.RoutingFilePath.Trim());
            SetSetting(KeyReasonFilePath,       snap.ReasonFilePath.Trim());
        }

        SetSetting(KeyLoginId,               snap.LoginId.Trim());
        SetSetting(KeyPassword,              snap.Password);
        SetSetting(KeyRwtLoginId,            snap.RwtLoginId.Trim());
        SetSetting(KeyRwtPassword,           snap.RwtPassword);
        SetSetting(KeyApiFacility,           snap.ApiFacility.Trim());
        SetSetting(KeyApiLoginId,            snap.ApiLoginId.Trim());
        SetSetting(KeyApiPassword,           snap.ApiPassword);
    }

    // ── Private ──────────────────────────────────────────────────────────────────

    private SqliteConnection Open()
    {
        EnsureDatabase();
        return PortalDb.Open();
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

    public void ReplaceRoutingRows(IReadOnlyList<RoutingRow> rows)
    {
        using var conn = Open();
        using var tx   = conn.BeginTransaction();
        using (var del = conn.CreateCommand())
        {
            del.Transaction = tx;
            del.CommandText = "DELETE FROM RoutingTable;";
            del.ExecuteNonQuery();
        }

        foreach (var r in rows)
        {
            using var ins = conn.CreateCommand();
            ins.Transaction = tx;
            ins.CommandText = """
                INSERT INTO RoutingTable (Id, ModelName, ProcessCode, ProcessName, ProcessType)
                VALUES (@id, @m, @pc, @pn, @pt);
                """;
            ins.Parameters.AddWithValue("@id", r.Id);
            ins.Parameters.AddWithValue("@m",  r.ModelName ?? string.Empty);
            ins.Parameters.AddWithValue("@pc", r.ProcessCode ?? string.Empty);
            ins.Parameters.AddWithValue("@pn", r.ProcessName ?? string.Empty);
            ins.Parameters.AddWithValue("@pt", r.ProcessType ?? string.Empty);
            ins.ExecuteNonQuery();
        }

        tx.Commit();
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

    public void ReplaceReasonRows(IReadOnlyList<ReasonRow> rows)
    {
        using var conn = Open();
        using var tx   = conn.BeginTransaction();
        using (var del = conn.CreateCommand())
        {
            del.Transaction = tx;
            del.CommandText = "DELETE FROM ReasonTable;";
            del.ExecuteNonQuery();
        }

        foreach (var r in rows)
        {
            using var ins = conn.CreateCommand();
            ins.Transaction = tx;
            ins.CommandText = """
                INSERT INTO ReasonTable (Id, ProcessName, NgName, Reason)
                VALUES (@id, @pn, @ng, @rs);
                """;
            ins.Parameters.AddWithValue("@id", r.Id);
            ins.Parameters.AddWithValue("@pn", r.ProcessName ?? string.Empty);
            ins.Parameters.AddWithValue("@ng", r.NgName ?? string.Empty);
            ins.Parameters.AddWithValue("@rs", r.Reason ?? string.Empty);
            ins.ExecuteNonQuery();
        }

        tx.Commit();
    }

    /// <summary>
    /// Inserts routing rows missing from RoutingTable (BMES API routing, see BmesRoutingScrapeService).
    /// Input per row: material name (MAKTX) → ModelName, process code (KTSCH = routing VLSCH) → ProcessCode,
    /// process name → ProcessName, line stage name (C0127 = routing LGUBN_TX) → ProcessType
    /// (Main Assy→MAIN, Sub Assy→SUB, Final Visual→VISUAL, Function→FUNCTION, others: raw value).
    /// Model/Code/Name are NormalizeText-ed on both sides before comparison
    /// (and on insert), matching the normalization applied downstream in
    /// Get Data's temp DB. Existing rows are never changed or removed.
    /// </summary>
    public int MergeRoutingRows(
        IEnumerable<(string Model, string Code, string Name, string StageName)> rows, IProgress<string>? progress = null)
    {
        var dedup = new Dictionary<(string Model, string Code, string Name), string>();
        int total = 0;
        foreach (var (rawModel, rawCode, rawName, stageName) in rows)
        {
            total++;
            string model = NormalizeText(rawModel);
            string code  = NormalizeText(rawCode);
            string name  = NormalizeText(rawName);
            if (model.Length == 0 || code.Length == 0 || name.Length == 0) continue;
            dedup.TryAdd((model, code, name), MapProcessType(stageName ?? string.Empty));
        }

        progress?.Report($"Routing: {total:N0} rows → {dedup.Count:N0} unique (Model, Code, Name) tuples.");

        var existing = GetRoutingRows()
            .Select(r => (
                Model: NormalizeText(r.ModelName),
                Code:  NormalizeText(r.ProcessCode),
                Name:  NormalizeText(r.ProcessName)))
            .ToHashSet();

        var missing = dedup.Where(kv => !existing.Contains(kv.Key)).ToList();
        if (missing.Count == 0)
        {
            progress?.Report("No new routing rows to add.");
            return 0;
        }

        using var wconn = Open();
        using var tx    = wconn.BeginTransaction();
        foreach (var kv in missing)
        {
            var (model, code, name) = kv.Key;
            using var ins = wconn.CreateCommand();
            ins.CommandText =
                "INSERT INTO RoutingTable (ModelName, ProcessCode, ProcessName, ProcessType) " +
                "VALUES (@m, @pc, @pn, @pt);";
            ins.Parameters.AddWithValue("@m",  model);
            ins.Parameters.AddWithValue("@pc", code);
            ins.Parameters.AddWithValue("@pn", name);
            ins.Parameters.AddWithValue("@pt", kv.Value);
            ins.ExecuteNonQuery();
        }
        tx.Commit();

        progress?.Report($"Added {missing.Count:N0} new row(s) to RoutingTable.");
        return missing.Count;
    }

    /// <summary>
    /// Material code → material name already in PortalGNLAB.db: the NG Rate cache (MATERIALCODE / MATERIALNAME,
    /// from AI_TOOL_00055 view S), then the old MES020240 routing dump (BmesRouting MATNR / MAKTX) when present.
    /// </summary>
    public Dictionary<string, string> GetKnownMaterialNames()
    {
        var names = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            using var conn = Open();
            void Read(string table, string sql)
            {
                using var check = conn.CreateCommand();
                check.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@t;";
                check.Parameters.AddWithValue("@t", table);
                if (Convert.ToInt64(check.ExecuteScalar() ?? 0L) == 0) return;

                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                using var rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    string code = rdr.IsDBNull(0) ? string.Empty : rdr.GetString(0).Trim();
                    string name = rdr.IsDBNull(1) ? string.Empty : rdr.GetString(1).Trim();
                    if (code.Length > 0 && name.Length > 0) names.TryAdd(code, name);
                }
            }

            Read("NgRateCache", "SELECT DISTINCT [MATERIALCODE], [MATERIALNAME] FROM [NgRateCache];");
            Read("BmesRouting", "SELECT DISTINCT [MATNR], [MAKTX] FROM [BmesRouting];");
        }
        catch
        {
            // Unreadable tables just give fewer names; the caller asks the API for the rest.
        }
        return names;
    }

    private static string MapProcessType(string lgubnTx) => lgubnTx.Trim() switch
    {
        "Main Assy"    => "MAIN",
        "Sub Assy"     => "SUB",
        "Final Visual" => "VISUAL",
        "Function"     => "FUNCTION",
        _              => lgubnTx.Trim(),
    };

    /// <summary>
    /// Scans the NG Rate cache table (NgRateCache in PortalGNLAB.db),
    /// unions DISTINCT (PROCESSNAME, NGNAME) pairs,
    /// applies NormalizeText to both columns, then inserts any pair not already
    /// present in ReasonTable (comparison is normalized on both sides).
    /// Returns the number of newly inserted rows.
    /// </summary>
    public int RefreshReasonFromDailyDbs(IProgress<string>? progress = null)
    {
        progress?.Report("Scanning NG Rate cache…");
        var dailyPairs = ReadCachedProcessNgPairs(null);
        progress?.Report($"Collected {dailyPairs.Count} distinct pair(s) from the NG Rate cache.");
        return MergeReasonPairs(dailyPairs, progress);
    }

    /// <summary>
    /// Inserts (ProcessName, NgName) pairs missing from ReasonTable with an empty Reason. Both columns are
    /// NormalizeText-ed and compared normalized; existing rows are never changed or removed.
    /// Returns the number of newly inserted rows.
    /// </summary>
    public int MergeReasonPairs(IEnumerable<(string ProcessName, string NgName)> pairs, IProgress<string>? progress = null)
    {
        var existing = GetReasonRows()
            .Select(r => (Pn: NormalizeText(r.ProcessName), Ng: NormalizeText(r.NgName)))
            .ToHashSet();

        var missing = pairs
            .Select(p => (Pn: NormalizeText(p.ProcessName), Ng: NormalizeText(p.NgName)))
            .Where(p => (p.Pn.Length > 0 || p.Ng.Length > 0) && !existing.Contains(p))
            .Distinct()
            .ToList();
        if (missing.Count == 0)
        {
            progress?.Report("No new pairs to add.");
            return 0;
        }

        using var wconn = Open();
        using var tx    = wconn.BeginTransaction();
        foreach (var (pn, ng) in missing)
        {
            using var ins = wconn.CreateCommand();
            ins.CommandText = "INSERT INTO ReasonTable (ProcessName,NgName,Reason) VALUES (@pn,@ng,'');";
            ins.Parameters.AddWithValue("@pn", pn);
            ins.Parameters.AddWithValue("@ng", ng);
            ins.ExecuteNonQuery();
        }
        tx.Commit();

        progress?.Report($"Added {missing.Count} new row(s) to ReasonTable.");
        return missing.Count;
    }

    /// <summary>
    /// Scans the NG Rate cache for DISTINCT (PROCESSNAME, NGNAME) pairs whose
    /// [MATERIALNAME] contains <paramref name="materialNamePattern"/> (case-insensitive,
    /// SQLite LIKE). Used by the Reason Table page to discover which reasons are still
    /// missing for a given model.
    /// </summary>
    public List<(string ProcessName, string NgName)> GetProcessNgPairsByMaterial(
        string materialNamePattern, IProgress<string>? progress = null)
    {
        if (string.IsNullOrWhiteSpace(materialNamePattern)) return [];

        progress?.Report($"Scanning NG Rate cache for material LIKE '{materialNamePattern}'…");
        var pairs = ReadCachedProcessNgPairs("%" + materialNamePattern.Trim() + "%");
        progress?.Report($"Found {pairs.Count} distinct pair(s).");
        return pairs
            .OrderBy(p => p.Pn, StringComparer.Ordinal)
            .ThenBy(p => p.Ng, StringComparer.Ordinal)
            .Select(p => (p.Pn, p.Ng))
            .ToList();
    }

    private HashSet<(string Pn, string Ng)> ReadCachedProcessNgPairs(string? materialLikePattern)
    {
        var pairs = new HashSet<(string Pn, string Ng)>();
        try
        {
            using var conn = Open();
            if (!CacheTableExists(conn)) return pairs;

            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                "SELECT DISTINCT [PROCESSNAME], [NGNAME] FROM [NgRateCache] " +
                "WHERE [PROCESSNAME] IS NOT NULL AND [NGNAME] IS NOT NULL" +
                (materialLikePattern is null
                    ? ";"
                    : " AND [MATERIALNAME] IS NOT NULL AND [MATERIALNAME] LIKE @mn COLLATE NOCASE;");
            if (materialLikePattern is not null)
                cmd.Parameters.AddWithValue("@mn", materialLikePattern);
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                string pn = NormalizeText(rdr.IsDBNull(0) ? string.Empty : rdr.GetString(0));
                string ng = NormalizeText(rdr.IsDBNull(1) ? string.Empty : rdr.GetString(1));
                if (pn.Length == 0 && ng.Length == 0) continue;
                pairs.Add((pn, ng));
            }
        }
        catch
        {
            // An unreadable cache just yields no pairs.
        }
        return pairs;
    }

    private static bool CacheTableExists(SqliteConnection conn)
    {
        using var check = conn.CreateCommand();
        check.CommandText =
            "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='NgRateCache';";
        return Convert.ToInt64(check.ExecuteScalar() ?? 0L) > 0;
    }

    /// <summary>
    /// Same logic as NgRateService.NormalizeText — duplicated here to avoid a
    /// circular dependency (NgRateService already depends on this class).
    /// </summary>
    private static string NormalizeText(string input)
    {
        if (string.IsNullOrWhiteSpace(input)) return string.Empty;
        input = input.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
        input = input.Replace("‘", "'").Replace("’", "'")
                     .Replace("“", "\"").Replace("”", "\"");
        input = input.Replace("'", " ").Replace("\"", " ").Replace("~", " ");
        input = input.Replace("[", "").Replace("]", "_").Replace("+", " ");
        input = Regex.Replace(input, @"\s{2,}", " ");
        return input.Trim();
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

    // ── LineShift Scan ────────────────────────────────────────────────────────────

    public List<string> GetAllLineShifts() => ReadCachedLineShifts(null);

    /// <summary>
    /// Returns distinct LineShift values (MATERIALNAME_PRODUCTION_LINE) from the NG Rate cache
    /// rows whose MATERIALNAME matches the given value, so a model that only ran in an earlier
    /// period still offers its shift groups to the Model Group editor.
    /// </summary>
    public List<string> GetLineShiftsByMaterialName(string materialName)
        => string.IsNullOrWhiteSpace(materialName) ? [] : ReadCachedLineShifts(materialName);

    private List<string> ReadCachedLineShifts(string? materialName)
    {
        var result = new HashSet<string>(StringComparer.Ordinal);
        try
        {
            using var conn = Open();
            if (!CacheTableExists(conn)) return [];

            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                "SELECT DISTINCT ([MATERIALNAME] || '_' || [PRODUCTION_LINE]) FROM [NgRateCache] " +
                "WHERE [MATERIALNAME] IS NOT NULL AND [MATERIALNAME] != ''" +
                (materialName is null ? ";" : " AND [MATERIALNAME] = @mn;");
            if (materialName is not null)
                cmd.Parameters.AddWithValue("@mn", materialName);
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                if (!rdr.IsDBNull(0))
                {
                    string val = rdr.GetString(0);
                    if (!string.IsNullOrEmpty(val)) result.Add(val);
                }
            }
        }
        catch
        {
            // swallow: an unreadable cache simply offers no shifts
        }
        return result.OrderBy(x => x).ToList();
    }

    // ── Private ──────────────────────────────────────────────────────────────────

    public void EnsureDatabase()
    {
        using var conn = PortalDb.Open();
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

/// <summary>Mutable snapshot used for UI binding. Path defaults are filled in by
/// <see cref="NgRateSettingsService.GetSnapshot"/> (sourced from AppPathsService).</summary>
public sealed class NgRateSettingsSnapshot
{
    public string DbSaveDirectory      { get; set; } = string.Empty;
    public string FCostDbSaveDirectory { get; set; } = string.Empty;
    public string RoutingFilePath      { get; set; } = string.Empty;
    public string ReasonFilePath       { get; set; } = string.Empty;
    public string LoginId              { get; set; } = string.Empty;
    public string Password             { get; set; } = string.Empty;
    public string RwtLoginId           { get; set; } = string.Empty;
    public string RwtPassword          { get; set; } = string.Empty;
    public string ApiFacility          { get; set; } = string.Empty;
    public string ApiLoginId           { get; set; } = string.Empty;
    public string ApiPassword          { get; set; } = string.Empty;
}
