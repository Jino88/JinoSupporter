using System.Globalization;
using Microsoft.Data.Sqlite;

namespace PortalGNLAB.Web.Services;

/// <summary>Where the pre-unification stores live. Any member may be null/missing.</summary>
public sealed record PortalDbMigrationSources(
    string? MainDb,
    string? SettingsDbDirectory,
    string? NgRateDbDirectory,
    string? FCostDbDirectory,
    string? ReviewReportDb)
{
    /// <summary>Resolves the old locations from the app path config (which still points at them).</summary>
    public static PortalDbMigrationSources FromAppPaths(AppPathsConfig cfg)
    {
        string root = (cfg.DbRootDirectory ?? string.Empty).Trim();
        string settingsDir = string.IsNullOrWhiteSpace(cfg.NgRateSettingsDbDirectory)
            ? AppStoragePaths.Combine("standalone-bootstrap", "ModelBmes")
            : cfg.NgRateSettingsDbDirectory;
        string fcostDir = root.Length > 0
            ? AppPathsService.GetFCostDirectory(root)
            : cfg.NgRateDbSaveDirectory;

        string? main = new[]
            {
                cfg.MainDbPath,
                root.Length > 0 ? Path.Combine(root, "process-review.db") : null,
                AppStoragePaths.Combine("process-review.db"),
            }
            .Where(p => !string.IsNullOrWhiteSpace(p) && File.Exists(p)
                && !string.Equals(Path.GetFileName(p), PortalDb.FileName, StringComparison.OrdinalIgnoreCase))
            .FirstOrDefault();

        return new PortalDbMigrationSources(
            main,
            settingsDir,
            cfg.NgRateDbSaveDirectory,
            fcostDir,
            AppStoragePaths.Combine("review-report.db"));
    }
}

public sealed record PortalDbMigrationLine(
    string Source, string SourceTable, string TargetTable, long SourceRows, long InsertedRows);

public sealed class PortalDbMigrationResult
{
    public bool AlreadyMigrated { get; init; }
    public bool CreatedFromMainDb { get; init; }
    public string TargetPath { get; init; } = string.Empty;
    public List<PortalDbMigrationLine> Lines { get; } = [];
}

/// <summary>Table and index layout of the NG Rate daily/monthly cache inside PortalGNLAB.db.
/// One row table keyed by PRODUCT_DATE replaces the per-month yyyyMM.db / per-day yyyyMMdd.db files.</summary>
internal static class NgRateCacheSchema
{
    public const string RowsTable = "NgRateCache";
    public const string DatesTable = "NgRateCachedDates";

    public static readonly string[] Columns =
    {
        "PRODUCTION_LINE", "PROCESSCODE", "PROCESSNAME", "NGCODE", "NGNAME",
        "USE", "QTYINPUT", "QTYNG", "INPUTQUANTITY", "NGQUANTITY",
        "MATERIALCODE", "MATERIALNAME", "PLANT", "WORKORDER", "Shift",
        "PRODUCT_DATE", "WORKODER", "ACTIVITYNO", "LASTREGPERSON", "LASTREGDATE",
    };

    public static void Ensure(SqliteConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            $"CREATE TABLE IF NOT EXISTS [{RowsTable}] ({string.Join(", ", Columns.Select(c => $"[{c}] TEXT"))});" +
            $"CREATE INDEX IF NOT EXISTS [IX_{RowsTable}_PRODUCT_DATE] ON [{RowsTable}]([PRODUCT_DATE]);" +
            $"CREATE TABLE IF NOT EXISTS [{DatesTable}] " +
            "([PRODUCT_DATE] TEXT PRIMARY KEY, [RowCount] INTEGER NOT NULL, [CachedAt] TEXT NOT NULL);";
        cmd.ExecuteNonQuery();
    }
}

/// <summary>
/// One-time move of every old SQLite store into <see cref="PortalDb.FilePath"/>.
/// The main DB is snapshotted (SQLite backup API) when the unified file does not exist yet, the
/// others are ATTACHed read-only and copied table by table. Old files are never modified or deleted,
/// so they stay as backups. Completion is recorded in <c>PortalDbMigrations</c> so it never re-runs.
/// </summary>
public static class PortalDbMigrator
{
    public const string CompletedKey = "UnifiedDb.v1";

    private sealed record TableRename(string SourceTable, string TargetTable, string? CreateSql);

    public static PortalDbMigrationResult EnsureMigrated(
        string targetPath,
        PortalDbMigrationSources sources,
        Action<string>? log = null)
    {
        log ??= _ => { };
        string? dir = Path.GetDirectoryName(targetPath);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);

        bool created = false;
        if (!File.Exists(targetPath))
        {
            created = CreateTargetFromMain(targetPath, sources.MainDb, log);
        }

        using var conn = PortalDb.Open(targetPath);
        Exec(conn, "PRAGMA foreign_keys = OFF;");
        Exec(conn, "PRAGMA journal_mode = WAL;");
        Exec(conn,
            "CREATE TABLE IF NOT EXISTS PortalDbMigrations (" +
            "Name TEXT PRIMARY KEY NOT NULL, CompletedAt TEXT NOT NULL, Detail TEXT NOT NULL DEFAULT '')");

        if (IsDone(conn, CompletedKey))
        {
            return new PortalDbMigrationResult { AlreadyMigrated = true, TargetPath = targetPath };
        }

        var res = new PortalDbMigrationResult { TargetPath = targetPath, CreatedFromMainDb = created };

        string settingsDir = sources.SettingsDbDirectory ?? string.Empty;
        string fcostDir = sources.FCostDbDirectory ?? string.Empty;
        string ngDir = sources.NgRateDbDirectory ?? string.Empty;

        ImportDatabase(conn, res, log, "ngrate_settings", Combine(settingsDir, "ngrate_settings.db"));
        ImportDatabase(conn, res, log, "bmes_routing_raw", Combine(settingsDir, "bmes_routing_raw.db"));
        ImportDatabase(conn, res, log, "bmes_pdm_raw", Combine(settingsDir, "bmes_pdm_raw.db"));
        ImportDatabase(conn, res, log, "bmes_bom_cache", Combine(settingsDir, "bmes_bom_cache.db"));
        ImportDatabase(conn, res, log, "fcost_raw", Combine(fcostDir, "fcost_raw.db"));
        ImportDatabase(conn, res, log, "ipg_defect_cache", Combine(fcostDir, "ipg_defect_cache.db"));
        ImportDatabase(conn, res, log, "review-report", sources.ReviewReportDb,
            new TableRename("reports", "ReviewReports", ReviewReportService.SchemaSql),
            new TableRename("report_images", "ReviewReportImages", null),
            new TableRename("report_values", "ReviewReportValues", null));
        ImportNgRateCaches(conn, res, log, ngDir);

        string detail = string.Join("; ", res.Lines.Select(l =>
            $"{l.Source}.{l.SourceTable}->{l.TargetTable}:{l.SourceRows}/{l.InsertedRows}"));
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText =
                "INSERT OR REPLACE INTO PortalDbMigrations (Name, CompletedAt, Detail) VALUES (@n, @t, @d)";
            cmd.Parameters.AddWithValue("@n", CompletedKey);
            cmd.Parameters.AddWithValue("@t", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
            cmd.Parameters.AddWithValue("@d", detail);
            cmd.ExecuteNonQuery();
        }

        log($"Unified DB migration complete: {res.Lines.Count} table copy step(s).");
        return res;
    }

    // ── History cache reset ────────────────────────────────────────────────────

    public const string HistoryCacheResetKey = "HistoryCacheReset.v1";

    /// <summary>BMES copies kept for NG Rate and F-COST. Most rows came over from the old DB
    /// files and were never re-checked, so they are dropped once and refilled under
    /// <see cref="HistoryCachePolicy"/>. Schemas stay; only rows go.</summary>
    internal static readonly string[] HistoryCacheTables =
    [
        NgRateCacheSchema.RowsTable, NgRateCacheSchema.DatesTable,
        "FCostRawPulls", "FCostRawRows", "FCostRawColumns", "FCostRawBreakdownCache",
        "MES072410Pulls", "MES072410Rows", "MES072410Columns",
    ];

    /// <summary>Empties <see cref="HistoryCacheTables"/> once per database; later calls only
    /// check the marker. Returns the number of rows removed.</summary>
    public static long EnsureHistoryCacheReset(string targetPath, Action<string>? log = null)
    {
        log ??= _ => { };
        using var conn = PortalDb.Open(targetPath);
        Exec(conn,
            "CREATE TABLE IF NOT EXISTS PortalDbMigrations (" +
            "Name TEXT PRIMARY KEY NOT NULL, CompletedAt TEXT NOT NULL, Detail TEXT NOT NULL DEFAULT '')");
        if (IsDone(conn, HistoryCacheResetKey))
            return 0;

        long total = 0;
        var detail = new List<string>();
        using var tx = conn.BeginTransaction();
        foreach (string table in HistoryCacheTables)
        {
            if (!TableExists(conn, tx, table)) continue;
            long rows = Scalar(conn, $"SELECT COUNT(*) FROM [{table}]", tx);
            Exec(conn, $"DELETE FROM [{table}]", tx);
            total += rows;
            detail.Add($"{table}:{rows}");
            log($"Cleared {table}: {rows:N0} row(s)");
        }
        Exec(conn,
            "INSERT OR REPLACE INTO PortalDbMigrations (Name, CompletedAt, Detail) VALUES (@n, @t, @d)", tx,
            ("@n", HistoryCacheResetKey),
            ("@t", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)),
            ("@d", string.Join("; ", detail)));
        tx.Commit();

        log($"History cache reset complete: {total:N0} row(s) removed.");
        return total;
    }

    // ── Target creation ────────────────────────────────────────────────────────

    private static bool CreateTargetFromMain(string targetPath, string? mainDb, Action<string> log)
    {
        if (string.IsNullOrWhiteSpace(mainDb) || !File.Exists(mainDb)
            || string.Equals(Path.GetFullPath(mainDb), Path.GetFullPath(targetPath), StringComparison.OrdinalIgnoreCase))
            return false;

        string tmp = targetPath + ".migrating";
        if (File.Exists(tmp)) File.Delete(tmp);

        log($"Creating {Path.GetFileName(targetPath)} from {mainDb} …");
        using (var source = new SqliteConnection(PortalDb.ConnectionString(mainDb, readOnly: true)))
        using (var dest = new SqliteConnection($"Data Source={tmp}"))
        {
            source.Open();
            dest.Open();
            source.BackupDatabase(dest);
        }
        SqliteConnection.ClearAllPools();
        File.Move(tmp, targetPath);
        return true;
    }

    // ── Generic table-by-table import ──────────────────────────────────────────

    private static void ImportDatabase(
        SqliteConnection conn,
        PortalDbMigrationResult res,
        Action<string> log,
        string label,
        string? path,
        params TableRename[] renames)
    {
        string marker = "Source:" + label;
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path) || IsDone(conn, marker))
            return;

        Attach(conn, path);
        try
        {
            using var tx = conn.BeginTransaction();
            var tables = ReadList(conn, tx,
                "SELECT name, sql FROM src.sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY rowid");
            var steps = new List<(string Source, string Target)>();

            foreach (var (name, sql) in tables)
            {
                TableRename? rename = renames.FirstOrDefault(r =>
                    string.Equals(r.SourceTable, name, StringComparison.OrdinalIgnoreCase));
                string target = rename?.TargetTable ?? name;

                if (rename?.CreateSql is { } createSql)
                {
                    Exec(conn, createSql, tx);
                }
                else if (rename is not null)
                {
                    // Created together with the renamed parent table by its CreateSql.
                }
                else if (TableExists(conn, tx, name))
                {
                    // Same name already used by another store (SQLite names are case-insensitive): prefix it.
                    target = label.Replace('-', '_') + "_" + name;
                    if (TableExists(conn, tx, target))
                        continue;
                    Exec(conn, $"CREATE TABLE [{target}] AS SELECT * FROM src.[{name}] WHERE 0", tx);
                }
                else
                {
                    Exec(conn, sql!, tx);
                    foreach (string indexSql in ReadList(conn, tx,
                                 $"SELECT sql, name FROM src.sqlite_master WHERE type='index' AND tbl_name=@t AND sql IS NOT NULL",
                                 ("@t", name)).Select(x => x.Item1))
                    {
                        try { Exec(conn, indexSql, tx); } catch (SqliteException) { /* index name taken */ }
                    }
                }

                steps.Add((name, target));
            }

            foreach (var (source, target) in steps)
                res.Lines.Add(CopyRows(conn, tx, label, source, target, whereClause: null));

            Exec(conn,
                "INSERT OR REPLACE INTO PortalDbMigrations (Name, CompletedAt, Detail) VALUES (@n, @t, @d)", tx,
                ("@n", marker),
                ("@t", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)),
                ("@d", path));
            tx.Commit();
            log($"  {label}: {steps.Count} table(s) copied.");
        }
        finally
        {
            Detach(conn);
        }
    }

    private static PortalDbMigrationLine CopyRows(
        SqliteConnection conn, SqliteTransaction tx, string label,
        string source, string target, string? whereClause)
    {
        var srcCols = ReadList(conn, tx, $"SELECT name, '' FROM pragma_table_info('{source}', 'src')").Select(x => x.Item1);
        var dstCols = ReadList(conn, tx, $"SELECT name, '' FROM pragma_table_info('{target}', 'main')")
            .Select(x => x.Item1).ToHashSet(StringComparer.OrdinalIgnoreCase);
        string cols = string.Join(", ", srcCols.Where(dstCols.Contains).Select(c => $"[{c}]"));

        long srcRows = Scalar(conn, $"SELECT COUNT(*) FROM src.[{source}]", tx);
        long inserted = 0;
        if (cols.Length > 0)
        {
            using var cmd = conn.CreateCommand();
            cmd.Transaction = tx;
            cmd.CommandText =
                $"INSERT OR IGNORE INTO main.[{target}] ({cols}) SELECT {cols} FROM src.[{source}] {whereClause}";
            inserted = cmd.ExecuteNonQuery();
        }
        return new PortalDbMigrationLine(label, source, target, srcRows, inserted);
    }

    // ── NG Rate monthly/daily caches ───────────────────────────────────────────

    private static void ImportNgRateCaches(
        SqliteConnection conn, PortalDbMigrationResult res, Action<string> log, string ngDir)
    {
        if (string.IsNullOrWhiteSpace(ngDir) || !Directory.Exists(ngDir))
            return;

        NgRateCacheSchema.Ensure(conn);

        var files = new List<(string Path, bool Daily)>();
        string monthly = Path.Combine(ngDir, "monthly");
        string daily = Path.Combine(ngDir, "daily");
        if (Directory.Exists(monthly))
            files.AddRange(Directory.EnumerateFiles(monthly, "??????.db").OrderBy(f => f, StringComparer.Ordinal)
                .Where(f => IsDigits(System.IO.Path.GetFileNameWithoutExtension(f), 6)).Select(f => (f, false)));
        if (Directory.Exists(daily))
            files.AddRange(Directory.EnumerateFiles(daily, "????????.db").OrderBy(f => f, StringComparer.Ordinal)
                .Where(f => IsDigits(System.IO.Path.GetFileNameWithoutExtension(f), 8)).Select(f => (f, true)));

        int imported = 0;
        foreach (var (file, isDaily) in files)
        {
            string label = $"ngcache:{(isDaily ? "daily" : "monthly")}/{System.IO.Path.GetFileNameWithoutExtension(file)}";
            string marker = "Source:" + label;
            if (IsDone(conn, marker))
                continue;

            Attach(conn, file);
            try
            {
                using var tx = conn.BeginTransaction();
                if (TableExists(conn, tx, "OrginalTable", "src"))
                {
                    var srcCols = ReadList(conn, tx, "SELECT name, '' FROM pragma_table_info('OrginalTable', 'src')")
                        .Select(x => x.Item1).ToList();
                    string cols = string.Join(", ",
                        NgRateCacheSchema.Columns.Where(c => srcCols.Contains(c, StringComparer.OrdinalIgnoreCase))
                            .Select(c => $"[{c}]"));

                    long srcRows = Scalar(conn, "SELECT COUNT(*) FROM src.[OrginalTable]", tx);
                    long inserted = 0;
                    if (cols.Length > 0)
                    {
                        // A daily file is only a staging area for the monthly cache: skip dates already cached.
                        string where = isDaily
                            ? $"WHERE COALESCE([PRODUCT_DATE], '') NOT IN (SELECT [PRODUCT_DATE] FROM [{NgRateCacheSchema.DatesTable}])"
                            : string.Empty;
                        using var cmd = conn.CreateCommand();
                        cmd.Transaction = tx;
                        cmd.CommandText =
                            $"INSERT INTO [{NgRateCacheSchema.RowsTable}] ({cols}) SELECT {cols} FROM src.[OrginalTable] {where}";
                        inserted = cmd.ExecuteNonQuery();
                    }

                    if (TableExists(conn, tx, "__NgRateCachedDates", "src"))
                    {
                        Exec(conn,
                            $"INSERT OR REPLACE INTO [{NgRateCacheSchema.DatesTable}] ([PRODUCT_DATE], [RowCount], [CachedAt]) " +
                            "SELECT [PRODUCT_DATE], [RowCount], [CachedAt] FROM src.[__NgRateCachedDates]", tx);
                    }
                    else if (isDaily)
                    {
                        string stamp = File.GetLastWriteTime(file).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                        Exec(conn,
                            $"INSERT OR IGNORE INTO [{NgRateCacheSchema.DatesTable}] ([PRODUCT_DATE], [RowCount], [CachedAt]) " +
                            "SELECT [PRODUCT_DATE], COUNT(*), @stamp FROM src.[OrginalTable] " +
                            "WHERE COALESCE([PRODUCT_DATE], '') <> '' GROUP BY [PRODUCT_DATE]", tx, ("@stamp", stamp));
                    }

                    res.Lines.Add(new PortalDbMigrationLine(label, "OrginalTable", NgRateCacheSchema.RowsTable, srcRows, inserted));
                }

                Exec(conn,
                    "INSERT OR REPLACE INTO PortalDbMigrations (Name, CompletedAt, Detail) VALUES (@n, @t, @d)", tx,
                    ("@n", marker),
                    ("@t", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)),
                    ("@d", file));
                tx.Commit();
                imported++;
            }
            finally
            {
                Detach(conn);
            }
        }

        log($"  NG Rate cache files imported: {imported}.");
    }

    // ── helpers ────────────────────────────────────────────────────────────────

    private static string? Combine(string dir, string file) =>
        string.IsNullOrWhiteSpace(dir) ? null : Path.Combine(dir, file);

    private static bool IsDigits(string s, int length) => s.Length == length && s.All(char.IsDigit);

    private static void Attach(SqliteConnection conn, string path)
    {
        // Read-only URI so the old files are guaranteed untouched.
        string uri = new Uri(Path.GetFullPath(path)).AbsoluteUri + "?mode=ro";
        try
        {
            Exec(conn, "ATTACH DATABASE @p AS src", null, ("@p", uri));
        }
        catch (SqliteException)
        {
            Exec(conn, "ATTACH DATABASE @p AS src", null, ("@p", path));
        }
    }

    private static void Detach(SqliteConnection conn)
    {
        try { Exec(conn, "DETACH DATABASE src"); } catch (SqliteException) { }
    }

    private static bool IsDone(SqliteConnection conn, string name) =>
        Scalar(conn, "SELECT COUNT(*) FROM PortalDbMigrations WHERE Name = @n", null, ("@n", name)) > 0;

    private static bool TableExists(SqliteConnection conn, SqliteTransaction? tx, string name, string schema = "main") =>
        Scalar(conn, $"SELECT COUNT(*) FROM {schema}.sqlite_master WHERE type='table' AND name = @n COLLATE NOCASE",
            tx, ("@n", name)) > 0;

    private static List<(string, string)> ReadList(
        SqliteConnection conn, SqliteTransaction? tx, string sql, params (string, object)[] args)
    {
        using var cmd = conn.CreateCommand();
        cmd.Transaction = tx;
        cmd.CommandText = sql;
        foreach (var (k, v) in args) cmd.Parameters.AddWithValue(k, v);
        var list = new List<(string, string)>();
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add((r.IsDBNull(0) ? string.Empty : r.GetString(0), r.IsDBNull(1) ? string.Empty : r.GetString(1)));
        return list;
    }

    private static long Scalar(
        SqliteConnection conn, string sql, SqliteTransaction? tx = null, params (string, object)[] args)
    {
        using var cmd = conn.CreateCommand();
        cmd.Transaction = tx;
        cmd.CommandText = sql;
        foreach (var (k, v) in args) cmd.Parameters.AddWithValue(k, v);
        return Convert.ToInt64(cmd.ExecuteScalar() ?? 0L, CultureInfo.InvariantCulture);
    }

    private static void Exec(
        SqliteConnection conn, string sql, SqliteTransaction? tx = null, params (string, object)[] args)
    {
        using var cmd = conn.CreateCommand();
        cmd.Transaction = tx;
        cmd.CommandText = sql;
        foreach (var (k, v) in args) cmd.Parameters.AddWithValue(k, v);
        cmd.ExecuteNonQuery();
    }
}
