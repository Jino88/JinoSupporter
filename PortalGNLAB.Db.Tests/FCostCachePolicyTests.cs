using Microsoft.Data.Sqlite;
using PortalGNLAB.Web.Services;
using Xunit;

namespace PortalGNLAB.Db.Tests;

public sealed class FCostCachePolicyTests : IDisposable
{
    private readonly string _dbPath = Path.Combine(Path.GetTempPath(), $"fcost-cache-{Guid.NewGuid():N}.db");
    private static readonly DateTime Now = new(2026, 10, 30, 9, 0, 0);

    public FCostCachePolicyTests()
    {
        Exec("""
            CREATE TABLE FCostRawPulls (QueryDate TEXT PRIMARY KEY, Status TEXT NOT NULL, FetchedAt TEXT NOT NULL,
                RowCount INTEGER NOT NULL DEFAULT 0, ColumnCount INTEGER NOT NULL DEFAULT 0, ErrorMessage TEXT NOT NULL DEFAULT '');
            CREATE TABLE MES072410Pulls (QueryDate TEXT PRIMARY KEY, Status TEXT NOT NULL, FetchedAt TEXT NOT NULL);
            """);
    }

    public void Dispose()
    {
        SqliteConnection.ClearAllPools();
        try { File.Delete(_dbPath); } catch { }
    }

    private void Exec(string sql)
    {
        using var c = new SqliteConnection($"Data Source={_dbPath}");
        c.Open();
        using var cmd = c.CreateCommand();
        cmd.CommandText = sql;
        cmd.ExecuteNonQuery();
    }

    private void Put(DateTime day, DateTime fetchedLocal, string status = "OK", int rows = 5)
    {
        string fetched = new DateTimeOffset(fetchedLocal, TimeZoneInfo.Local.GetUtcOffset(fetchedLocal)).ToString("o");
        Exec($"""
            INSERT OR REPLACE INTO FCostRawPulls (QueryDate, Status, FetchedAt, RowCount) VALUES ('{day:yyyy-MM-dd}', '{status}', '{fetched}', {rows});
            INSERT OR REPLACE INTO MES072410Pulls (QueryDate, Status, FetchedAt) VALUES ('{day:yyyy-MM-dd}', '{status}', '{fetched}');
            """);
    }

    private bool Raw(DateTime day) => FCostService.RawPullSucceeded(_dbPath, day, Now);
    private bool Kpi(DateTime day) => FCostCorePartsService.PullSucceeded(_dbPath, day, Now);

    [Fact]
    public void CopyFetchedWhileStillLiveIsFetchedAgain()
    {
        var day = new DateTime(2026, 10, 1);
        Put(day, new DateTime(2026, 10, 3, 8, 0, 0));
        Assert.False(Raw(day));
        Assert.False(Kpi(day));
    }

    [Fact]
    public void SettledCopyIsSkipped()
    {
        var day = new DateTime(2026, 10, 1);
        Put(day, new DateTime(2026, 10, 9, 8, 0, 0));
        Assert.True(Raw(day));
        Assert.True(Kpi(day));
    }

    [Fact]
    public void RecentDayIsReusedOnlyFor15Minutes()
    {
        var day = Now.Date.AddDays(-1);
        Put(day, Now.AddMinutes(-10));
        Assert.True(Raw(day));
        Assert.True(Kpi(day));
        Put(day, Now.AddMinutes(-20));
        Assert.False(Raw(day));
        Assert.False(Kpi(day));
    }

    [Fact]
    public void FailedOrEmptyCopyIsNeverReused()
    {
        var day = new DateTime(2026, 10, 1);
        Put(day, new DateTime(2026, 10, 9), status: "FAIL");
        Assert.False(Raw(day));
        Put(day, new DateTime(2026, 10, 9), rows: 0);
        Assert.False(Raw(day));
        Assert.False(Raw(new DateTime(2026, 9, 1))); // no row
    }

    [Fact]
    public async Task ForceAlwaysFetchesAgain()
    {
        // force=true must skip nothing; with no credentials the backfill reports the missing login
        // after deciding every day needs a fetch.
        var day = new DateTime(2026, 10, 1);
        Put(day, new DateTime(2026, 10, 9, 8, 0, 0));
        Assert.True(Raw(day));
        // BackfillRawAsync writes into the configured DB, so the force rule is checked by its guard:
        Assert.True(FCostService.ShouldSkipRaw(force: false, cachedUsable: true));
        Assert.False(FCostService.ShouldSkipRaw(force: true, cachedUsable: true));
        await Task.CompletedTask;
    }
}
