using Microsoft.Data.Sqlite;
using PortalGNLAB.Web.Services;
using Xunit;

namespace PortalGNLAB.Db.Tests;

/// <summary>Migration logic against small throw-away SQLite files (never the real data folder).</summary>
public sealed class PortalDbMigratorTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "portaldb-tests-" + Guid.NewGuid().ToString("N"));

    public PortalDbMigratorTests() => Directory.CreateDirectory(_root);

    public void Dispose()
    {
        SqliteConnection.ClearAllPools();
        try { Directory.Delete(_root, recursive: true); } catch { }
    }

    private string P(params string[] parts) => Path.Combine([_root, .. parts]);

    private static void Run(string dbPath, params string[] statements)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(dbPath)!);
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        foreach (string sql in statements)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            cmd.ExecuteNonQuery();
        }
    }

    private static long Count(string dbPath, string table)
    {
        using var conn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"SELECT COUNT(*) FROM [{table}]";
        return Convert.ToInt64(cmd.ExecuteScalar());
    }

    private static string? Scalar(string dbPath, string sql)
    {
        using var conn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        return cmd.ExecuteScalar()?.ToString();
    }

    private const string NgColumns =
        "[PRODUCTION_LINE] TEXT, [PROCESSCODE] TEXT, [PROCESSNAME] TEXT, [NGCODE] TEXT, [NGNAME] TEXT, [USE] TEXT, " +
        "[QTYINPUT] TEXT, [QTYNG] TEXT, [INPUTQUANTITY] TEXT, [NGQUANTITY] TEXT, [MATERIALCODE] TEXT, [MATERIALNAME] TEXT, " +
        "[PLANT] TEXT, [WORKORDER] TEXT, [Shift] TEXT, [PRODUCT_DATE] TEXT, [WORKODER] TEXT, [ACTIVITYNO] TEXT, " +
        "[LASTREGPERSON] TEXT, [LASTREGDATE] TEXT";

    private static string NgRow(string date, string material) =>
        $"INSERT INTO [OrginalTable] ([PRODUCTION_LINE],[MATERIALNAME],[PRODUCT_DATE],[QTYNG]) VALUES ('L1','{material}','{date}','1')";

    private PortalDbMigrationSources BuildSources()
    {
        Run(P("old", "process-review.db"),
            "CREATE TABLE AppSettings (Key TEXT PRIMARY KEY, Value TEXT)",
            "INSERT INTO AppSettings VALUES ('a','1'), ('b','2')",
            "CREATE TABLE Users (Id INTEGER PRIMARY KEY, Name TEXT)",
            "INSERT INTO Users VALUES (1,'jino')",
            "CREATE TABLE Reports (Id INTEGER PRIMARY KEY, Title TEXT)",
            "INSERT INTO Reports VALUES (1,'main report')");

        Run(P("old", "ng", "ngrate_settings.db"),
            "CREATE TABLE NgRateSettings (Key TEXT PRIMARY KEY NOT NULL, Value TEXT NOT NULL DEFAULT '')",
            "INSERT INTO NgRateSettings VALUES ('NgRate:LoginId','web-id'), ('NgRate:Password','web-pw'), " +
            "('BmesApi:Facility','GN'), ('BmesApi:LoginId','api-id'), ('BmesApi:Password','api-pw')",
            "CREATE TABLE RoutingTable (Id INTEGER PRIMARY KEY AUTOINCREMENT, ModelName TEXT NOT NULL DEFAULT '')",
            "INSERT INTO RoutingTable (Id, ModelName) VALUES (7,'M1'), (9,'M2')",
            // same name as a main-DB table: must be prefixed, not merged
            "CREATE TABLE Users (Id INTEGER PRIMARY KEY, Login TEXT)",
            "INSERT INTO Users VALUES (1,'legacy')");

        Run(P("old", "ng", "bmes_pdm_raw.db"),
            "CREATE TABLE BmesPdm (Id INTEGER PRIMARY KEY, PDRTX TEXT)",
            "CREATE INDEX IX_BmesPdm_PDRTX ON BmesPdm(PDRTX)",
            "INSERT INTO BmesPdm VALUES (1,'x'), (2,'y'), (3,'z')");

        Run(P("old", "fcost", "fcost_raw.db"),
            "CREATE TABLE FCostRawRows (Id INTEGER PRIMARY KEY, SDate TEXT)",
            "INSERT INTO FCostRawRows VALUES (1,'2026-01-01'), (2,'2026-01-02')");

        // review-report.db uses lower-case names that collide with the main DB's Reports
        Run(P("old", "review-report.db"),
            "CREATE TABLE reports (id TEXT PRIMARY KEY, title TEXT NOT NULL DEFAULT '', created_by TEXT NOT NULL DEFAULT '', " +
            "created_at TEXT NOT NULL, updated_at TEXT NOT NULL, language TEXT NOT NULL DEFAULT '', input_text TEXT NOT NULL DEFAULT '', " +
            "extra_instruction TEXT NOT NULL DEFAULT '', report_json TEXT, model TEXT NOT NULL DEFAULT '')",
            "INSERT INTO reports (id,title,created_at,updated_at) VALUES ('r1','T1','2026-01-01','2026-01-01'), ('r2','T2','2026-01-02','2026-01-02')",
            "CREATE TABLE report_images (report_id TEXT NOT NULL, no INTEGER NOT NULL, file_name TEXT NOT NULL DEFAULT '', " +
            "media_type TEXT NOT NULL DEFAULT 'image/png', data BLOB NOT NULL, PRIMARY KEY (report_id, no))",
            "INSERT INTO report_images (report_id,no,data) VALUES ('r1',1,x'0102')",
            "CREATE TABLE report_values (report_id TEXT NOT NULL, metric_id TEXT NOT NULL, arm_id TEXT NOT NULL, " +
            "metric_name TEXT NOT NULL DEFAULT '', unit TEXT NOT NULL DEFAULT '', kind TEXT NOT NULL DEFAULT '', rate_of TEXT NOT NULL DEFAULT '', " +
            "arm_label TEXT NOT NULL DEFAULT '', arm_role TEXT NOT NULL DEFAULT '', source_display TEXT NOT NULL DEFAULT '', number REAL, " +
            "PRIMARY KEY (report_id, metric_id, arm_id))",
            "INSERT INTO report_values (report_id,metric_id,arm_id,number) VALUES ('r1','m','a',1.5)");

        // Two monthly caches (3 + 2 rows) and one daily file that repeats a monthly date and adds a new one.
        string ngDir = P("old", "ng");
        foreach (var (month, rows) in new[]
                 {
                     ("202601", new[] { ("2026-01-05", "A"), ("2026-01-05", "B"), ("2026-01-06", "A") }),
                     ("202602", new[] { ("2026-02-01", "A"), ("2026-02-01", "C") }),
                 })
        {
            var stmts = new List<string>
            {
                $"CREATE TABLE [OrginalTable] ({NgColumns})",
                "CREATE TABLE [__NgRateCachedDates] ([PRODUCT_DATE] TEXT PRIMARY KEY, [RowCount] INTEGER NOT NULL, [CachedAt] TEXT NOT NULL)",
            };
            stmts.AddRange(rows.Select(r => NgRow(r.Item1, r.Item2)));
            foreach (var g in rows.GroupBy(r => r.Item1))
                stmts.Add($"INSERT INTO [__NgRateCachedDates] VALUES ('{g.Key}', {g.Count()}, '2026-03-01 10:00:00')");
            Run(Path.Combine(ngDir, "monthly", month + ".db"), stmts.ToArray());
        }
        Run(Path.Combine(ngDir, "daily", "20260106.db"),
            $"CREATE TABLE [OrginalTable] ({NgColumns})",
            NgRow("2026-01-06", "DUP"),
            NgRow("2026-01-07", "NEW"));

        return new PortalDbMigrationSources(
            P("old", "process-review.db"),
            P("old", "ng"),
            P("old", "ng"),
            P("old", "fcost"),
            P("old", "review-report.db"));
    }

    [Fact]
    public void Migrates_every_store_and_row_counts_match()
    {
        var sources = BuildSources();
        string target = P("data", "PortalGNLAB.db");

        var result = PortalDbMigrator.EnsureMigrated(target, sources);

        Assert.False(result.AlreadyMigrated);
        Assert.True(result.CreatedFromMainDb);

        // main DB copied as is
        Assert.Equal(2, Count(target, "AppSettings"));
        Assert.Equal(1, Count(target, "Reports"));

        // key/value settings keep their keys and values (web login + API login)
        Assert.Equal("web-id", Scalar(target, "SELECT Value FROM NgRateSettings WHERE Key='NgRate:LoginId'"));
        Assert.Equal("GN", Scalar(target, "SELECT Value FROM NgRateSettings WHERE Key='BmesApi:Facility'"));
        Assert.Equal("api-pw", Scalar(target, "SELECT Value FROM NgRateSettings WHERE Key='BmesApi:Password'"));

        // ids survive and AUTOINCREMENT continues after the highest copied id
        Assert.Equal(2, Count(target, "RoutingTable"));
        Run(target, "INSERT INTO RoutingTable (ModelName) VALUES ('next')");
        Assert.Equal("10", Scalar(target, "SELECT MAX(Id) FROM RoutingTable"));

        // plain stores: same table name, same rows, index recreated
        Assert.Equal(3, Count(target, "BmesPdm"));
        Assert.Equal("1", Scalar(target, "SELECT COUNT(*) FROM sqlite_master WHERE type='index' AND name='IX_BmesPdm_PDRTX'"));
        Assert.Equal(2, Count(target, "FCostRawRows"));

        // name collision with the main DB -> prefixed copy, main table untouched
        Assert.Equal(1, Count(target, "Users"));
        Assert.Equal("jino", Scalar(target, "SELECT Name FROM Users"));
        Assert.Equal(1, Count(target, "ngrate_settings_Users"));

        // review-report.db: reports collides with Reports (case-insensitive) -> ReviewReports*
        Assert.Equal(2, Count(target, "ReviewReports"));
        Assert.Equal(1, Count(target, "ReviewReportImages"));
        Assert.Equal(1, Count(target, "ReviewReportValues"));
        Assert.Equal("[]", Scalar(target, "SELECT revisions_json FROM ReviewReports WHERE id='r1'"));

        // NG caches: 3 + 2 monthly rows + 1 daily row for the date not cached yet (its duplicate date is skipped)
        Assert.Equal(6, Count(target, "NgRateCache"));
        Assert.Equal("0", Scalar(target, "SELECT COUNT(*) FROM NgRateCache WHERE MATERIALNAME='DUP'"));
        Assert.Equal("1", Scalar(target, "SELECT COUNT(*) FROM NgRateCache WHERE MATERIALNAME='NEW'"));
        Assert.Equal(4, Count(target, "NgRateCachedDates"));
        Assert.Equal("2026-03-01 10:00:00", Scalar(target, "SELECT CachedAt FROM NgRateCachedDates WHERE PRODUCT_DATE='2026-01-05'"));

        // every step's reported counts agree with what landed in the target
        foreach (var line in result.Lines.Where(l => l.TargetTable != "NgRateCache"))
            Assert.Equal(line.SourceRows, line.InsertedRows);
    }

    [Fact]
    public void Second_run_is_a_noop_and_old_files_are_left_in_place()
    {
        var sources = BuildSources();
        string target = P("data", "PortalGNLAB.db");
        string[] oldFiles =
        [
            sources.MainDb!, P("old", "ng", "ngrate_settings.db"), P("old", "fcost", "fcost_raw.db"),
            P("old", "ng", "monthly", "202601.db"), sources.ReviewReportDb!,
        ];
        var before = oldFiles.ToDictionary(f => f, f => (new FileInfo(f).Length, File.GetLastWriteTimeUtc(f)));

        PortalDbMigrator.EnsureMigrated(target, sources);
        long cacheRows = Count(target, "NgRateCache");

        // The old files change afterwards; a re-run must not pull anything in again.
        Run(P("old", "ng", "ngrate_settings.db"), "INSERT INTO NgRateSettings VALUES ('late','x')");
        var second = PortalDbMigrator.EnsureMigrated(target, sources);

        Assert.True(second.AlreadyMigrated);
        Assert.Equal(cacheRows, Count(target, "NgRateCache"));
        Assert.Equal(5, Count(target, "NgRateSettings"));
        Assert.Equal("1", Scalar(target, "SELECT COUNT(*) FROM PortalDbMigrations WHERE Name='UnifiedDb.v1'"));

        foreach (string f in oldFiles.Where(f => f != P("old", "ng", "ngrate_settings.db")))
        {
            Assert.True(File.Exists(f));
            Assert.Equal(before[f].Item1, new FileInfo(f).Length);
        }
        Assert.True(File.Exists(P("old", "ng", "ngrate_settings.db")));
    }

    [Fact]
    public void Missing_sources_still_create_a_usable_empty_unified_db()
    {
        string target = P("fresh", "PortalGNLAB.db");

        var result = PortalDbMigrator.EnsureMigrated(target,
            new PortalDbMigrationSources(null, P("nope"), P("nope"), P("nope"), P("nope", "x.db")));

        Assert.False(result.CreatedFromMainDb);
        Assert.True(File.Exists(target));
        Assert.Empty(result.Lines);
        Assert.Equal("1", Scalar(target, "SELECT COUNT(*) FROM PortalDbMigrations WHERE Name='UnifiedDb.v1'"));
    }

    [Fact]
    public void Review_report_schema_matches_the_migration_target()
    {
        string target = P("schema", "PortalGNLAB.db");
        Run(target, ReviewReportService.SchemaSql);

        Assert.Equal(0, Count(target, "ReviewReports"));
        Assert.Equal("1", Scalar(target,
            "SELECT COUNT(*) FROM pragma_table_info('ReviewReports') WHERE name='revisions_json'"));
    }
}
