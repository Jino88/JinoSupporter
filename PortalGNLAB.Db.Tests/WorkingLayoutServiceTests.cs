using System.Text;
using System.Text.Json;
using PortalGNLAB.Web.Services;
using Xunit;

namespace PortalGNLAB.Db.Tests;

public sealed class WorkingLayoutServiceTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "working-layout-tests-" + Guid.NewGuid().ToString("N"));

    public void Dispose()
    {
        try { Directory.Delete(_dir, true); } catch { }
    }

    static WorkingLayoutServiceTests() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    /// <summary>A small R2000 drawing in CP949: layers, a block put in twice (scaled, turned), a polyline with a bulge, texts.</summary>
    private static byte[] Drawing(string codePage = "ANSI_949", string version = "AC1015")
    {
        string[] pairs =
        [
            "0", "SECTION", "2", "HEADER", "9", "$ACADVER", "1", version, "9", "$DWGCODEPAGE", "3", codePage, "0", "ENDSEC",
            "0", "SECTION", "2", "TABLES", "0", "TABLE", "2", "LAYER",
            "0", "LAYER", "2", "0", "70", "0", "62", "7",
            "0", "LAYER", "2", "WALL", "70", "0", "62", "1",
            "0", "LAYER", "2", "OFF", "70", "0", "62", "-3",
            "0", "ENDTAB", "0", "ENDSEC",
            "0", "SECTION", "2", "BLOCKS",
            "0", "BLOCK", "8", "0", "2", "MAN", "70", "0", "10", "100", "20", "100",
            "0", "LINE", "8", "0", "62", "0", "10", "100", "20", "100", "11", "110", "21", "100",
            "0", "ENDBLK", "8", "0",
            "0", "ENDSEC",
            "0", "SECTION", "2", "ENTITIES",
            "0", "LINE", "8", "WALL", "10", "0", "20", "0", "11", "50", "21", "0",
            "0", "LINE", "8", "OFF", "10", "0", "20", "0", "11", "1", "21", "1",
            "0", "LWPOLYLINE", "8", "WALL", "62", "5", "90", "2", "70", "0", "10", "0", "20", "0", "42", "1", "10", "10", "20", "0",
            "0", "CIRCLE", "8", "0", "10", "5", "20", "5", "40", "2",
            "0", "INSERT", "8", "WALL", "62", "3", "2", "MAN", "10", "200", "20", "0",
            "0", "INSERT", "8", "WALL", "2", "MAN", "10", "300", "20", "0", "41", "2", "42", "2", "50", "90",
            "0", "TEXT", "8", "0", "10", "1", "20", "2", "40", "2.5", "1", "작업자 %%p0.05",
            "0", "MTEXT", "8", "WALL", "10", "7", "20", "8", "40", "3", "71", "5",
            "1", "{\\fGulim|b1|i0|c129|p50;UPPER & SPK\\P ARRAY}",
            "0", "TEXT", "8", "0", "10", "0", "20", "0", "11", "40", "21", "9", "72", "1", "73", "2", "40", "1", "1", "\\U+AC00 MID",
            "0", "ENDSEC", "0", "EOF",
        ];
        Encoding enc = codePage == "UTF8" ? new UTF8Encoding(false) : Encoding.GetEncoding(949);
        return enc.GetBytes(string.Join("\r\n", pairs) + "\r\n");
    }

    private static WorkingLayout Parse(byte[] bytes) =>
        WorkingLayoutService.Parse(new MemoryStream(bytes), "C:\\x\\line.dxf", bytes.Length, "jhbyun");

    [Fact]
    public void Parse_FlattensLinesArcsAndBlocksIntoWorldCoordinates()
    {
        WorkingLayout layout = Parse(Drawing());

        Assert.Equal("line.dxf", layout.FileName);
        Assert.Equal("jhbyun", layout.UploadedBy);
        // the line on the layer that is off is left out
        Assert.DoesNotContain(layout.Paths, p => layout.Layers[p.L] == "OFF");
        var line = layout.Paths[0];
        Assert.Equal("WALL", layout.Layers[line.L]);
        Assert.Equal("#ff0000", line.C);  // BYLAYER: red layer
        Assert.Equal([0d, 0, 50, 0], line.P);

        // a bulge of 1 is a half circle from (0,0) to (10,0) through (5,-5)
        var arc = layout.Paths[1];
        Assert.Equal("#0000ff", arc.C);
        Assert.True(arc.P.Length > 8);
        Assert.Equal(0, arc.P[0]);
        Assert.Equal(10, arc.P[^2]);
        Assert.Contains(Enumerable.Range(0, arc.P.Length / 2), i => Math.Abs(arc.P[2 * i] - 5) < 0.01 && Math.Abs(arc.P[2 * i + 1] + 5) < 0.01);

        var circle = layout.Paths[2];
        Assert.Equal(1, circle.Z);
        Assert.Null(circle.C);  // colour 7: the screen's foreground

        // block line (100,100)-(110,100) with base (100,100): at (200,0) as is, in BYBLOCK green on the INSERT's layer
        var first = layout.Paths[3];
        Assert.Equal("WALL", layout.Layers[first.L]);
        Assert.Equal("#00ff00", first.C);
        Assert.Equal([200d, 0, 210, 0], first.P);
        // twice as large and turned 90 degrees at (300,0): goes up 20
        Assert.Equal([300d, 0, 300, 20], layout.Paths[4].P);

        Assert.Equal([0d, -5, 300, 20], layout.Bounds);
    }

    [Fact]
    public void Parse_ReadsKoreanTextsAndDropsMTextFormatting()
    {
        WorkingLayout layout = Parse(Drawing());

        Assert.Equal(3, layout.Texts.Count);
        var korean = layout.Texts[0];
        Assert.Equal("작업자 ±0.05", korean.T);
        Assert.Equal(2.5, korean.H);
        Assert.Null(korean.A);

        var mtext = layout.Texts[1];
        Assert.Equal("UPPER & SPK\n ARRAY", mtext.T);
        Assert.Equal(1, mtext.A);  // attachment 5: middle centre
        Assert.Equal(2, mtext.V);

        // a centred, middle TEXT sits on its second point
        var mid = layout.Texts[2];
        Assert.Equal("가 MID", mid.T);
        Assert.Equal((40d, 9d, 1, 2), (mid.X, mid.Y, mid.A, mid.V));
    }

    [Fact]
    public void EncodingOf_UsesUtf8FromAutoCad2007AndTheCodePageBefore()
    {
        Assert.Equal(949, WorkingLayoutService.EncodingOf(Drawing()).CodePage);
        Assert.Equal(65001, WorkingLayoutService.EncodingOf(Drawing("UTF8", "AC1032")).CodePage);
        Assert.Equal("작업자 ±0.05", Parse(Drawing("UTF8", "AC1032")).Texts[0].T);
    }

    [Fact]
    public void Parse_RejectsWhatIsNotADrawing()
    {
        Assert.Throws<ArgumentException>(() => Parse(Encoding.ASCII.GetBytes("AutoCAD Binary DXF\r\n\u001a\0")));
        Assert.Throws<ArgumentException>(() => Parse(Encoding.ASCII.GetBytes("AC1032\0\0\0")));
        Assert.Throws<ArgumentException>(() => Parse(Encoding.ASCII.GetBytes("hello")));
    }

    [Fact]
    public void CleanMText_DropsCodesAndKeepsStacksAndEscapes()
    {
        Assert.Equal("1.작업자 : 54명\n2.OQC: 0명", WorkingLayoutService.CleanMText("1.{\\f굴림|b0|i0|c129|p50;작업자} : 54{\\f굴림;명}\\P2.{\\Fsimplex.shx|c129;OQC}: 0명"));
        Assert.Equal("1/2 {a}", WorkingLayoutService.CleanMText("\\S1^2; \\{a\\}"));
        Assert.Equal("12시간 자연", WorkingLayoutService.CleanMText("{\\fGulim|b1|i0|c129|p50;\\C2;12시간 자연 }"));
    }

    [Fact]
    public void Aci_GivesTheStandardColoursAndNullForWhite()
    {
        Assert.Null(WorkingLayoutService.Aci(7));
        Assert.Equal("#ff0000", WorkingLayoutService.Aci(1));
        Assert.Equal("#ff0000", WorkingLayoutService.Aci(10));
        Assert.Equal("#808080", WorkingLayoutService.Aci(8));
        Assert.Equal("#ffffff", WorkingLayoutService.Aci(255));
    }

    [Fact]
    public void Save_KeepsOneLayoutPerModelAndDeletes()
    {
        var svc = new WorkingLayoutService(_dir);
        byte[] bytes = Drawing();
        svc.Save("MSM-G778B", "a.dxf", bytes.Length, new MemoryStream(bytes), "u1");
        Assert.Null(svc.LoadGzip("MSM-X316B"));
        byte[] gz = svc.LoadGzip("msm-g778b ")!;
        WorkingLayout back = WorkingLayoutService.FromGzip(gz);
        Assert.Equal("a.dxf", back.FileName);
        Assert.Equal(Parse(bytes).Paths.Count, back.Paths.Count);

        // the JSON the screen reads: camelCase short keys, z/a/v only when set
        string json = Encoding.UTF8.GetString(new System.IO.Compression.GZipStream(new MemoryStream(gz), System.IO.Compression.CompressionMode.Decompress).ReadAllBytesForTest());
        using var doc = JsonDocument.Parse(json);
        JsonElement firstPath = doc.RootElement.GetProperty("paths")[0];
        Assert.Equal("[\"l\",\"c\",\"p\"]", JsonSerializer.Serialize(firstPath.EnumerateObject().Select(p => p.Name)));
        Assert.True(doc.RootElement.GetProperty("bounds").GetArrayLength() == 4);
        Assert.Equal("u1", doc.RootElement.GetProperty("uploadedBy").GetString());

        Assert.True(svc.Delete("MSM-G778B"));
        Assert.False(svc.Delete("MSM-G778B"));
        Assert.Null(svc.LoadGzip("MSM-G778B"));
    }

    /// <summary>The real line layout, when LAYOUT_DXF points at it (not in the repository: 90 MB).</summary>
    [Fact]
    public void Parse_RealLayout_WhenGiven()
    {
        string? path = Environment.GetEnvironmentVariable("LAYOUT_DXF");
        if (string.IsNullOrEmpty(path) || !File.Exists(path)) return;
        using FileStream fs = File.OpenRead(path);
        var watch = System.Diagnostics.Stopwatch.StartNew();
        WorkingLayout layout = WorkingLayoutService.Parse(fs, path, fs.Length, "real");
        byte[] gz = WorkingLayoutService.Gzip(layout);
        Console.WriteLine($"paths {layout.Paths.Count}, points {layout.Paths.Sum(p => p.P.Length / 2)}, texts {layout.Texts.Count}, " +
                          $"layers {layout.Layers.Count}, bounds {string.Join(",", layout.Bounds)}, gzip {gz.Length:N0} bytes, {watch.ElapsedMilliseconds} ms");
        Assert.Contains(layout.Texts, t => t.T == "DT LINE LAY-OUT [MSM-G778B]");
        Assert.Contains(layout.Texts, t => t.T.Contains("작업자 : 54명"));
        Assert.Contains(layout.Texts, t => t.T == "SPK ADHESION");
        Assert.True(layout.Paths.Count > 10_000);
    }
}

internal static class StreamTestExtensions
{
    public static byte[] ReadAllBytesForTest(this Stream s)
    {
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }
}
