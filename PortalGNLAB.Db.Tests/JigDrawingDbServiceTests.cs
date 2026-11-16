using Microsoft.Data.Sqlite;
using PortalGNLAB.Web.Services;
using Xunit;
using Xunit.Abstractions;

namespace PortalGNLAB.Db.Tests;

public sealed class JigDrawingDbServiceTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "jigdb-" + Guid.NewGuid().ToString("N"));
    private readonly ITestOutputHelper _out;

    public JigDrawingDbServiceTests(ITestOutputHelper output)
    {
        _out = output;
        Directory.CreateDirectory(_dir);
    }

    public void Dispose()
    {
        SqliteConnection.ClearAllPools();
        try { Directory.Delete(_dir, true); } catch { }
    }

    private string MakeDb()
    {
        string file = Path.Combine(_dir, "ORIGINAL_DRAWING.db");
        using var conn = new SqliteConnection($"Data Source={file};Pooling=False");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            CREATE TABLE drawings (model TEXT, drawing_no TEXT, base TEXT, name TEXT, assy INTEGER, category TEXT, dir TEXT, latest INTEGER);
            INSERT INTO drawings VALUES
             ('BRS-161242','BJKP18-08-I-0000-00','BJKP18-08-I-0000','CD LASER GUIDE BASE',1,'MC','',1),
             ('BRS-161242','BJKP18-08-I-0000-01','BJKP18-08-I-0000','CD LASER GUIDE BASE',1,'MC','',0),
             ('OTHER-1','BJKP18-09-I-0000-00','BJKP18-09-I-0000','CD LASER GUIDE BASE',0,'MC','',1),
             ('MSU-L15S10-20','BJGP26-22-T-0001-00','BJGP26-22-T-0001','(SUB1)DOME LASER CUTTING JIG',1,'JIG','',1),
             ('X','ZZ1-00','ZZ1','UNRELATED PART',0,'ETC','',1);
            """;
        cmd.ExecuteNonQuery();
        return file;
    }

    private JigDrawingDbService Service(string path) => new(() => path, _ => { });

    [Fact]
    public void Check_FoundRevisionMissing()
    {
        MakeDb();
        var r = Service(_dir).Check([" bjkp18-08-i-0000-00 ", "BJKP18-08-I-0000-05", "NOPE-99-00"]);
        Assert.True(r.Ok);
        Assert.Equal(["found", "revision", "missing"], r.Results.Select(x => x.Status).ToArray());
        Assert.Equal("CD LASER GUIDE BASE", r.Results[0].Match!.Name);
        Assert.Equal("BJKP18-08-I-0000-00", r.Results[1].Match!.DrawingNo); // latest=1 of the base
        Assert.Null(r.Results[2].Match);
    }

    [Fact]
    public void Check_NoDb_ReturnsOkFalse()
    {
        var r = Service(Path.Combine(_dir, "nothing")).Check(["A-00"]);
        Assert.False(r.Ok);
        Assert.NotEmpty(r.Error);
        Assert.Empty(r.Results);
    }

    [Fact]
    public void Suggest_RanksByTokens_DedupesBase_ModelBonus()
    {
        MakeDb();
        var svc = Service(Path.Combine(_dir, "ORIGINAL_DRAWING.db"));
        var r = svc.Suggest("DOME LASER GUIDE BASE", null);
        Assert.True(r.Ok);
        Assert.Contains("BJKP18", r.Suggestions[0].DrawingNo);
        Assert.Single(r.Suggestions, s => s.DrawingNo.StartsWith("BJKP18-08-I-0000")); // base deduped
        Assert.DoesNotContain(r.Suggestions, s => s.DrawingNo.StartsWith("ZZ1"));
        Assert.Equal(r.Suggestions.OrderByDescending(s => s.Score).Select(s => s.DrawingNo), r.Suggestions.Select(s => s.DrawingNo));

        // model bonus lifts the matching model above an otherwise equal drawing
        var withModel = svc.Suggest("DOME LASER GUIDE BASE", "OTHER-1");
        Assert.Equal("BJKP18-09-I-0000-00", withModel.Suggestions[0].DrawingNo);
    }

    [Fact]
    public void Suggest_JigNameFavorsJigCategory_AndTest_ReportsCounts()
    {
        MakeDb();
        var svc = Service(_dir);
        var r = svc.Suggest("LASER JIG", null);
        Assert.Equal("JIG", r.Suggestions[0].Category);
        var (ok, msg) = svc.Test();
        Assert.True(ok);
        Assert.Equal("5 drawings · 4 models", msg);
    }

    [Fact]
    public void RealDb_CheckAndSuggest_SkippedWhenMissing()
    {
        string real = JigDrawingDbService.Resolve(JigDrawingDbService.DefaultPath);
        if (real.Length == 0) return;
        var svc = Service(real);
        var r = svc.Check(["BJKP18-08-I-0000-00", "NOPE-NOT-A-DRAWING-00"]);
        _out.WriteLine($"real DB check: ok={r.Ok} {string.Join(" | ", r.Results.Select(x => $"{x.DrawingNo}={x.Status}({x.Match?.Name})"))}");
        Assert.True(r.Ok, r.Error);
        Assert.Equal("found", r.Results[0].Status);
        Assert.Equal("missing", r.Results[1].Status);
        var s = svc.Suggest("DOME LASER GUIDE BASE", null, 5);
        _out.WriteLine("real DB suggest: " + string.Join(" | ", s.Suggestions.Select(x => $"{x.DrawingNo} {x.Name} {x.Score}")));
        Assert.True(s.Ok, s.Error);
    }
}
