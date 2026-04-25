using Microsoft.Data.Sqlite;
using System.Text.RegularExpressions;

namespace JinoSupporter.Web.Services;

/// <summary>
/// NG Rate settings DB service.
/// Settings file location is now configurable via <see cref="AppPathsService"/>
/// (see Admin → Paths). Defaults: %LOCALAPPDATA%\…\NGRATE\ModelBmes\ngrate_settings.db.
/// </summary>
public sealed class NgRateSettingsService
{
    private readonly AppPathsService _appPaths;

    // ── Defaults (sourced from AppPathsService → webapp-paths.json) ─────────────
    public string DefaultDbSaveDirectory     => _appPaths.Current.NgRateDbSaveDirectory;
    public string DefaultRoutingFilePath     => _appPaths.Current.NgRateRoutingFilePath;
    public string DefaultReasonFilePath      => _appPaths.Current.NgRateReasonFilePath;

    // ── Settings DB location ─────────────────────────────────────────────────────
    public string SettingsDbDirectory => _appPaths.Current.NgRateSettingsDbDirectory;
    public string SettingsDbPath      => Path.Combine(SettingsDbDirectory, "ngrate_settings.db");

    // ── Setting keys ─────────────────────────────────────────────────────────────
    public const string KeyDbSaveDirectory  = "NgRate:DbSaveDirectory";
    public const string KeyRoutingFilePath  = "NgRate:RoutingFilePath";
    public const string KeyReasonFilePath   = "NgRate:ReasonFilePath";
    public const string KeyLoginId          = "NgRate:LoginId";
    public const string KeyPassword         = "NgRate:Password";
    // Request Work Time-specific credentials
    public const string KeyRwtLoginId       = "Rwt:LoginId";
    public const string KeyRwtPassword      = "Rwt:Password";

    // ── Constructor ──────────────────────────────────────────────────────────────
    public NgRateSettingsService(AppPathsService appPaths)
    {
        _appPaths = appPaths;
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

    public string RwtLoginId =>
        GetSetting(KeyRwtLoginId) ?? string.Empty;

    public string RwtPassword =>
        GetSetting(KeyRwtPassword) ?? string.Empty;

    public bool IsRwtCredentialsConfigured =>
        !string.IsNullOrWhiteSpace(RwtLoginId) && !string.IsNullOrWhiteSpace(RwtPassword);

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
        DbSaveDirectory = DbSaveDirectory,
        RoutingFilePath = RoutingFilePath,
        ReasonFilePath  = ReasonFilePath,
        LoginId         = LoginId,
        Password        = Password,
        RwtLoginId      = RwtLoginId,
        RwtPassword     = RwtPassword,
    };

    public void ApplySnapshot(NgRateSettingsSnapshot snap)
    {
        SetSetting(KeyDbSaveDirectory, snap.DbSaveDirectory.Trim());
        SetSetting(KeyRoutingFilePath, snap.RoutingFilePath.Trim());
        SetSetting(KeyReasonFilePath,  snap.ReasonFilePath.Trim());
        SetSetting(KeyLoginId,         snap.LoginId.Trim());
        SetSetting(KeyPassword,        snap.Password);
        SetSetting(KeyRwtLoginId,      snap.RwtLoginId.Trim());
        SetSetting(KeyRwtPassword,     snap.RwtPassword);
    }

    // ── Private ──────────────────────────────────────────────────────────────────

    private SqliteConnection Open()
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

    /// <summary>
    /// Reads the standalone BMES routing dump (bmes_routing_raw.db) produced by
    /// BmesRoutingScrapeService and inserts rows missing from RoutingTable.
    /// Mapping:
    ///   MAKTX    → ModelName
    ///   VLSCH    → ProcessCode
    ///   VLSCH_TX → ProcessName
    ///   LGUBN_TX → ProcessType (Main Assy→MAIN, Sub Assy→SUB,
    ///                           Final Visual→VISUAL, Function→FUNCTION,
    ///                           others: raw value)
    /// Model/Code/Name are NormalizeText-ed on both sides before comparison
    /// (and on insert), matching the normalization applied downstream in
    /// Get Data's temp DB.
    /// </summary>
    public int MergeRoutingFromRawDb(string rawDbPath, IProgress<string>? progress = null)
    {
        if (!File.Exists(rawDbPath))
        {
            progress?.Report($"[WARN] Raw routing DB not found: {rawDbPath}");
            return 0;
        }

        var rawDedup = new Dictionary<(string Model, string Code, string Name), string>();
        int rawTotal = 0;
        try
        {
            var cs = new SqliteConnectionStringBuilder
            {
                DataSource = rawDbPath,
                Mode       = SqliteOpenMode.ReadOnly,
            }.ToString();
            using var conn = new SqliteConnection(cs);
            conn.Open();

            using (var check = conn.CreateCommand())
            {
                check.CommandText =
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='BmesRouting';";
                if (Convert.ToInt64(check.ExecuteScalar() ?? 0L) == 0)
                {
                    progress?.Report("[WARN] BmesRouting table not found in raw DB.");
                    return 0;
                }
            }

            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT MAKTX, VLSCH, VLSCH_TX, LGUBN_TX FROM BmesRouting;";
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                rawTotal++;
                string model = NormalizeText(rdr.IsDBNull(0) ? string.Empty : rdr.GetString(0));
                string code  = NormalizeText(rdr.IsDBNull(1) ? string.Empty : rdr.GetString(1));
                string name  = NormalizeText(rdr.IsDBNull(2) ? string.Empty : rdr.GetString(2));
                string lgTx  = rdr.IsDBNull(3) ? string.Empty : rdr.GetString(3).Trim();
                if (model.Length == 0 || code.Length == 0 || name.Length == 0) continue;
                rawDedup.TryAdd((model, code, name), MapProcessType(lgTx));
            }
        }
        catch (Exception ex)
        {
            progress?.Report($"[ERROR] Read raw DB failed: {ex.Message}");
            return 0;
        }

        progress?.Report($"Raw: {rawTotal:N0} rows → {rawDedup.Count:N0} unique (Model, Code, Name) tuples.");

        var existing = GetRoutingRows()
            .Select(r => (
                Model: NormalizeText(r.ModelName),
                Code:  NormalizeText(r.ProcessCode),
                Name:  NormalizeText(r.ProcessName)))
            .ToHashSet();

        var missing = rawDedup.Where(kv => !existing.Contains(kv.Key)).ToList();
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

    private static string MapProcessType(string lgubnTx) => lgubnTx.Trim() switch
    {
        "Main Assy"    => "MAIN",
        "Sub Assy"     => "SUB",
        "Final Visual" => "VISUAL",
        "Function"     => "FUNCTION",
        _              => lgubnTx.Trim(),
    };

    /// <summary>
    /// Scans all *.db files in DbSaveDirectory (top-level + "daily" subfolder),
    /// unions DISTINCT (PROCESSNAME, NGNAME) pairs from each OrginalTable,
    /// applies NormalizeText to both columns, then inserts any pair not already
    /// present in ReasonTable (comparison is normalized on both sides).
    /// Returns the number of newly inserted rows.
    /// </summary>
    public int RefreshReasonFromDailyDbs(IProgress<string>? progress = null)
    {
        var dailyPairs = new HashSet<(string Pn, string Ng)>();
        var files      = EnumerateDbFilesForScan();

        progress?.Report($"Scanning {files.Count} DB file(s)…");
        int scanned = 0;
        foreach (var file in files)
        {
            scanned++;
            try
            {
                var cs = new SqliteConnectionStringBuilder
                {
                    DataSource = file,
                    Mode       = SqliteOpenMode.ReadOnly,
                }.ToString();

                using var conn = new SqliteConnection(cs);
                conn.Open();

                using (var check = conn.CreateCommand())
                {
                    check.CommandText =
                        "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='OrginalTable';";
                    if (Convert.ToInt64(check.ExecuteScalar() ?? 0L) == 0) continue;
                }

                using var cmd = conn.CreateCommand();
                cmd.CommandText =
                    "SELECT DISTINCT [PROCESSNAME], [NGNAME] FROM [OrginalTable] " +
                    "WHERE [PROCESSNAME] IS NOT NULL AND [NGNAME] IS NOT NULL;";
                using var rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    string pn = NormalizeText(rdr.IsDBNull(0) ? string.Empty : rdr.GetString(0));
                    string ng = NormalizeText(rdr.IsDBNull(1) ? string.Empty : rdr.GetString(1));
                    if (pn.Length == 0 && ng.Length == 0) continue;
                    dailyPairs.Add((pn, ng));
                }
            }
            catch (Exception ex)
            {
                progress?.Report($"[WARN] {Path.GetFileName(file)}: {ex.Message}");
            }
        }

        progress?.Report($"Collected {dailyPairs.Count} distinct pair(s) from {scanned} DB file(s).");

        var existing = GetReasonRows()
            .Select(r => (Pn: NormalizeText(r.ProcessName), Ng: NormalizeText(r.NgName)))
            .ToHashSet();

        var missing = dailyPairs.Where(p => !existing.Contains(p)).ToList();
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
    /// Scans all daily *.db files for DISTINCT (PROCESSNAME, NGNAME) pairs whose
    /// [MATERIALNAME] contains <paramref name="materialNamePattern"/> (case-insensitive,
    /// SQLite LIKE). Used by the Reason Table page to discover which reasons are still
    /// missing for a given model.
    /// </summary>
    public List<(string ProcessName, string NgName)> GetProcessNgPairsByMaterial(
        string materialNamePattern, IProgress<string>? progress = null)
    {
        var pairs = new HashSet<(string Pn, string Ng)>();
        if (string.IsNullOrWhiteSpace(materialNamePattern)) return pairs.ToList();

        var files = EnumerateDbFilesForScan();
        progress?.Report($"Scanning {files.Count} DB file(s) for material LIKE '{materialNamePattern}'…");

        string pattern = "%" + materialNamePattern.Trim() + "%";
        int scanned = 0;
        foreach (var file in files)
        {
            scanned++;
            try
            {
                var cs = new SqliteConnectionStringBuilder
                {
                    DataSource = file,
                    Mode       = SqliteOpenMode.ReadOnly,
                }.ToString();

                using var conn = new SqliteConnection(cs);
                conn.Open();

                using (var check = conn.CreateCommand())
                {
                    check.CommandText =
                        "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='OrginalTable';";
                    if (Convert.ToInt64(check.ExecuteScalar() ?? 0L) == 0) continue;
                }

                using var cmd = conn.CreateCommand();
                cmd.CommandText =
                    "SELECT DISTINCT [PROCESSNAME], [NGNAME] FROM [OrginalTable] " +
                    "WHERE [PROCESSNAME] IS NOT NULL AND [NGNAME] IS NOT NULL " +
                    "AND [MATERIALNAME] IS NOT NULL " +
                    "AND [MATERIALNAME] LIKE @mn COLLATE NOCASE;";
                cmd.Parameters.AddWithValue("@mn", pattern);
                using var rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    string pn = NormalizeText(rdr.IsDBNull(0) ? string.Empty : rdr.GetString(0));
                    string ng = NormalizeText(rdr.IsDBNull(1) ? string.Empty : rdr.GetString(1));
                    if (pn.Length == 0 && ng.Length == 0) continue;
                    pairs.Add((pn, ng));
                }
            }
            catch (Exception ex)
            {
                progress?.Report($"[WARN] {Path.GetFileName(file)}: {ex.Message}");
            }
        }

        progress?.Report($"Found {pairs.Count} distinct pair(s) across {scanned} DB file(s).");
        return pairs
            .OrderBy(p => p.Pn, StringComparer.Ordinal)
            .ThenBy(p => p.Ng, StringComparer.Ordinal)
            .ToList();
    }

    /// <summary>
    /// Collect *.db files from DbSaveDirectory top-level + "daily" subfolder,
    /// excluding ngrate_settings.db.
    /// </summary>
    private List<string> EnumerateDbFilesForScan()
    {
        var dbDir = DbSaveDirectory;
        var files = new List<string>();

        if (Directory.Exists(dbDir))
        {
            files.AddRange(
                Directory.EnumerateFiles(dbDir, "*.db", SearchOption.TopDirectoryOnly)
                         .Where(f => !string.Equals(
                             Path.GetFileName(f), "ngrate_settings.db",
                             StringComparison.OrdinalIgnoreCase)));
        }
        var dailyDir = Path.Combine(dbDir, "daily");
        if (Directory.Exists(dailyDir))
        {
            files.AddRange(
                Directory.EnumerateFiles(dailyDir, "*.db", SearchOption.TopDirectoryOnly)
                         .Where(f => !string.Equals(
                             Path.GetFileName(f), "ngrate_settings.db",
                             StringComparison.OrdinalIgnoreCase)));
        }
        return files;
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

    public List<string> GetAllLineShifts()
    {
        var dbDir  = DbSaveDirectory;
        var result = new HashSet<string>(StringComparer.Ordinal);

        // Collect candidate DB files: top-level + optional "daily" subdirectory
        var files = new List<string>();

        if (Directory.Exists(dbDir))
        {
            files.AddRange(
                Directory.EnumerateFiles(dbDir, "*.db", SearchOption.TopDirectoryOnly)
                         .Where(f => !string.Equals(
                             Path.GetFileName(f), "ngrate_settings.db",
                             StringComparison.OrdinalIgnoreCase)));
        }

        var dailyDir = Path.Combine(dbDir, "daily");
        if (Directory.Exists(dailyDir))
        {
            files.AddRange(
                Directory.EnumerateFiles(dailyDir, "*.db", SearchOption.TopDirectoryOnly)
                         .Where(f => !string.Equals(
                             Path.GetFileName(f), "ngrate_settings.db",
                             StringComparison.OrdinalIgnoreCase)));
        }

        foreach (var file in files)
        {
            try
            {
                var cs = new SqliteConnectionStringBuilder
                {
                    DataSource = file,
                    Mode       = SqliteOpenMode.ReadOnly,
                }.ToString();

                using var conn = new SqliteConnection(cs);
                conn.Open();

                // Check OrginalTable exists
                using var checkCmd = conn.CreateCommand();
                checkCmd.CommandText =
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='OrginalTable';";
                var tableCount = Convert.ToInt64(checkCmd.ExecuteScalar() ?? 0L);
                if (tableCount == 0) continue;

                // Check LineShift column exists
                bool hasLineShift = false;
                using (var pragmaCmd = conn.CreateCommand())
                {
                    pragmaCmd.CommandText = "PRAGMA table_info([OrginalTable]);";
                    using var rdr = pragmaCmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        if (string.Equals(rdr.GetString(1), "LineShift", StringComparison.Ordinal))
                        {
                            hasLineShift = true;
                            break;
                        }
                    }
                }

                using var selCmd = conn.CreateCommand();
                if (hasLineShift)
                {
                    selCmd.CommandText =
                        "SELECT DISTINCT [LineShift] FROM [OrginalTable] " +
                        "WHERE [LineShift] IS NOT NULL AND [LineShift] != '';";
                }
                else
                {
                    selCmd.CommandText =
                        "SELECT DISTINCT ([MATERIALNAME] || '_' || [PRODUCTION_LINE]) FROM [OrginalTable] " +
                        "WHERE [MATERIALNAME] IS NOT NULL AND [MATERIALNAME] != '';";
                }

                using var selRdr = selCmd.ExecuteReader();
                while (selRdr.Read())
                {
                    if (!selRdr.IsDBNull(0))
                    {
                        var val = selRdr.GetString(0);
                        if (!string.IsNullOrEmpty(val))
                            result.Add(val);
                    }
                }
            }
            catch
            {
                // swallow per-file exceptions
            }
        }

        return result.OrderBy(x => x).ToList();
    }

    /// <summary>
    /// Returns distinct LineShift values from OrginalTable rows whose MATERIALNAME
    /// matches the given value. Scans the same set of *.db files as GetAllLineShifts.
    /// </summary>
    public List<string> GetLineShiftsByMaterialName(string materialName)
    {
        var result = new HashSet<string>(StringComparer.Ordinal);
        if (string.IsNullOrWhiteSpace(materialName)) return new List<string>();

        var dbDir = DbSaveDirectory;
        var files = new List<string>();
        if (Directory.Exists(dbDir))
        {
            files.AddRange(
                Directory.EnumerateFiles(dbDir, "*.db", SearchOption.TopDirectoryOnly)
                         .Where(f => !string.Equals(
                             Path.GetFileName(f), "ngrate_settings.db",
                             StringComparison.OrdinalIgnoreCase)));
        }
        var dailyDir = Path.Combine(dbDir, "daily");
        if (Directory.Exists(dailyDir))
        {
            files.AddRange(
                Directory.EnumerateFiles(dailyDir, "*.db", SearchOption.TopDirectoryOnly)
                         .Where(f => !string.Equals(
                             Path.GetFileName(f), "ngrate_settings.db",
                             StringComparison.OrdinalIgnoreCase)));
        }

        foreach (var file in files)
        {
            try
            {
                var cs = new SqliteConnectionStringBuilder
                {
                    DataSource = file,
                    Mode       = SqliteOpenMode.ReadOnly,
                }.ToString();

                using var conn = new SqliteConnection(cs);
                conn.Open();

                using var checkCmd = conn.CreateCommand();
                checkCmd.CommandText =
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='OrginalTable';";
                var tableCount = Convert.ToInt64(checkCmd.ExecuteScalar() ?? 0L);
                if (tableCount == 0) continue;

                bool hasLineShift = false;
                using (var pragmaCmd = conn.CreateCommand())
                {
                    pragmaCmd.CommandText = "PRAGMA table_info([OrginalTable]);";
                    using var rdr = pragmaCmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        if (string.Equals(rdr.GetString(1), "LineShift", StringComparison.Ordinal))
                        {
                            hasLineShift = true;
                            break;
                        }
                    }
                }

                using var selCmd = conn.CreateCommand();
                if (hasLineShift)
                {
                    selCmd.CommandText =
                        "SELECT DISTINCT [LineShift] FROM [OrginalTable] " +
                        "WHERE [MATERIALNAME] = @mn " +
                        "AND [LineShift] IS NOT NULL AND [LineShift] != '';";
                }
                else
                {
                    selCmd.CommandText =
                        "SELECT DISTINCT ([MATERIALNAME] || '_' || [PRODUCTION_LINE]) FROM [OrginalTable] " +
                        "WHERE [MATERIALNAME] = @mn " +
                        "AND [MATERIALNAME] IS NOT NULL AND [MATERIALNAME] != '';";
                }
                selCmd.Parameters.AddWithValue("@mn", materialName);

                using var selRdr = selCmd.ExecuteReader();
                while (selRdr.Read())
                {
                    if (!selRdr.IsDBNull(0))
                    {
                        var val = selRdr.GetString(0);
                        if (!string.IsNullOrEmpty(val)) result.Add(val);
                    }
                }
            }
            catch { }
        }

        return result.OrderBy(x => x).ToList();
    }

    // ── Private ──────────────────────────────────────────────────────────────────

    private void EnsureDatabase()
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

/// <summary>Mutable snapshot used for UI binding. Path defaults are filled in by
/// <see cref="NgRateSettingsService.GetSnapshot"/> (sourced from AppPathsService).</summary>
public sealed class NgRateSettingsSnapshot
{
    public string DbSaveDirectory { get; set; } = string.Empty;
    public string RoutingFilePath { get; set; } = string.Empty;
    public string ReasonFilePath  { get; set; } = string.Empty;
    public string LoginId         { get; set; } = string.Empty;
    public string Password        { get; set; } = string.Empty;
    public string RwtLoginId      { get; set; } = string.Empty;
    public string RwtPassword     { get; set; } = string.Empty;
}
