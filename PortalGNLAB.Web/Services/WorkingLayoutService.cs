using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PortalGNLAB.Web.Services;

/// <summary>A line of the drawing in world coordinates: layer index, colour (null = the screen's foreground), x,y pairs, closed.</summary>
public sealed record WorkingLayoutPath(
    int L, string? C, double[] P, [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] int? Z);

/// <summary>
/// A text of the drawing: base point, height and rotation (degrees, counter-clockwise) in world units, the plain text
/// (lines split by \n), horizontal alignment (0 left, 1 centre, 2 right) and vertical one (0 baseline, 1 bottom, 2 middle, 3 top).
/// </summary>
public sealed record WorkingLayoutText(
    int L, string? C, double X, double Y, double H, double R, string T,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] int? A,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] int? V);

/// <summary>The process layout of a model: a DXF drawing flattened to lines and texts (blocks, arcs and splines spread out).</summary>
public sealed record WorkingLayout(
    string FileName, long Size, DateTime UploadedAt, string UploadedBy, double[] Bounds, List<string> Layers,
    List<WorkingLayoutPath> Paths, List<WorkingLayoutText> Texts);

/// <summary>
/// Reads a DXF (ASCII, R12 to current) into a <see cref="WorkingLayout"/> and keeps one per model as gzip JSON under
/// &lt;data&gt;\PortalGNLAB\working-layouts. The drawing itself is not kept, only what is shown.
/// </summary>
public sealed class WorkingLayoutService(string? root = null)
{
    public const long MaxUploadBytes = 300L * 1024 * 1024;
    private const int MaxPaths = 600_000;
    private const int MaxDepth = 12;
    /// <summary>How far (world units, mm in a line layout) a curve's lines may be from the true curve.</summary>
    private const double Tolerance = 0.5;

    public static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web);

    static WorkingLayoutService() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    private string Root => root ?? AppStoragePaths.AppFolder("working-layouts");

    private string FileOf(string modelName)
    {
        byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes(modelName.Trim().ToLowerInvariant()));
        return Path.Combine(Root, Convert.ToHexString(hash, 0, 16).ToLowerInvariant() + ".json.gz");
    }

    /// <summary>The stored layout as gzip JSON, or null.</summary>
    public byte[]? LoadGzip(string modelName)
    {
        string path = FileOf(modelName);
        return File.Exists(path) ? File.ReadAllBytes(path) : null;
    }

    /// <summary>Reads the DXF and stores what is shown for the model; returns the stored gzip JSON.</summary>
    public byte[] Save(string modelName, string fileName, long size, Stream dxf, string user)
    {
        WorkingLayout layout = Parse(dxf, fileName, size, user);
        byte[] gz = Gzip(layout);
        string path = FileOf(modelName);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        string tmp = path + ".tmp";
        File.WriteAllBytes(tmp, gz);
        File.Move(tmp, path, overwrite: true);
        return gz;
    }

    public bool Delete(string modelName)
    {
        string path = FileOf(modelName);
        if (!File.Exists(path)) return false;
        File.Delete(path);
        return true;
    }

    public static byte[] Gzip(WorkingLayout layout)
    {
        using var ms = new MemoryStream();
        using (var gz = new GZipStream(ms, CompressionLevel.Optimal, leaveOpen: true))
            JsonSerializer.Serialize(gz, layout, Json);
        return ms.ToArray();
    }

    public static WorkingLayout FromGzip(byte[] gz)
    {
        using var s = new GZipStream(new MemoryStream(gz), CompressionMode.Decompress);
        return JsonSerializer.Deserialize<WorkingLayout>(s, Json)!;
    }

    // ---------------------------------------------------------------- reading

    private sealed record Pair(int Code, string Value);

    private sealed class Ent(string type)
    {
        public readonly string Type = type;
        public readonly List<Pair> Pairs = [];
        public List<Ent>? Attribs;  // ATTRIB after an INSERT, VERTEX of a POLYLINE

        public string? S(int code) => Pairs.FirstOrDefault(p => p.Code == code)?.Value;
        public double D(int code, double fallback = 0) => S(code) is { } v ? Num(v) : fallback;
        public int I(int code, int fallback = 0) => S(code) is { } v && int.TryParse(v.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int n) ? n : fallback;
        public IEnumerable<double> All(int code) => Pairs.Where(p => p.Code == code).Select(p => Num(p.Value));
    }

    private sealed record Block(string Name, double BaseX, double BaseY, List<Ent> Ents);
    private sealed record Layer(string Name, int Color, string? True, bool Hidden);

    private static double Num(string v) =>
        double.TryParse(v.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double d) && double.IsFinite(d) ? d : 0;

    /// <summary>The text encoding of the drawing: UTF-8 from AutoCAD 2007 (AC1021) on, else the code page in the header.</summary>
    public static Encoding EncodingOf(byte[] head)
    {
        string text = Encoding.ASCII.GetString(head);
        string? Header(string name)
        {
            int i = text.IndexOf("\n" + name, StringComparison.Ordinal);
            if (i < 0) return null;
            string[] lines = text[(i + 1)..].Split('\n', 5);
            return lines.Length >= 3 ? lines[2].Trim() : null;
        }
        string ver = Header("$ACADVER") ?? "";
        if (string.CompareOrdinal(ver, "AC1021") >= 0) return new UTF8Encoding(false);
        string page = (Header("$DWGCODEPAGE") ?? "").ToUpperInvariant();
        int cp = page.StartsWith("ANSI_", StringComparison.Ordinal) && int.TryParse(page[5..], out int n) ? n
            : page switch { "DOS949" or "KSC5601" => 949, "DOS932" => 932, "GB2312" => 936, "BIG5" => 950, _ => 1252 };
        try { return Encoding.GetEncoding(cp); }
        catch (ArgumentException) { return Encoding.Latin1; }
        catch (NotSupportedException) { return Encoding.Latin1; }
    }

    /// <summary>Reads an ASCII DXF into lines and texts in world coordinates.</summary>
    public static WorkingLayout Parse(Stream dxf, string fileName, long size, string user)
    {
        Stream src = dxf;
        if (!src.CanSeek)
        {
            var copy = new MemoryStream();
            src.CopyTo(copy);
            src = copy;
        }
        src.Position = 0;
        byte[] head = new byte[Math.Min(256 * 1024, src.Length)];
        src.ReadExactly(head);
        src.Position = 0;
        if (Encoding.ASCII.GetString(head, 0, Math.Min(22, head.Length)).StartsWith("AutoCAD Binary DXF", StringComparison.Ordinal))
            throw new ArgumentException("This is a binary DXF. Save the drawing as an ASCII DXF and upload it again.");
        if (head.Length >= 4 && head[0] == 'A' && head[1] == 'C' && head[2] == '1' && head[3] == '0')
            throw new ArgumentException("This is a DWG file. Save the drawing as DXF and upload it again.");

        Encoding enc = EncodingOf(head);
        using var reader = new StreamReader(src, enc, detectEncodingFromByteOrderMarks: true, bufferSize: 1 << 16, leaveOpen: true);
        var doc = new Reader(reader);
        doc.Read();
        if (doc.Pairs == 0) throw new ArgumentException("The file is not a DXF drawing.");

        var b = new Builder(doc);
        b.Draw(doc.Entities, Affine.Identity, null, null, 0);
        if (b.Paths.Count == 0 && b.Texts.Count == 0) throw new ArgumentException("The drawing has nothing to show in its model space.");
        double[] bounds = b.MinX <= b.MaxX ? [R(b.MinX), R(b.MinY), R(b.MaxX), R(b.MaxY)] : [0, 0, 1, 1];
        return new WorkingLayout(Path.GetFileName(fileName), size, DateTime.UtcNow, user, bounds, b.LayerNames, b.Paths, b.Texts);
    }

    private static double R(double v) => Math.Round(v, 2);

    /// <summary>Group code / value pairs into the layer table, the blocks and the model-space entities.</summary>
    private sealed class Reader(StreamReader r)
    {
        public int Pairs;
        public readonly Dictionary<string, Layer> Layers = new(StringComparer.OrdinalIgnoreCase);
        public readonly Dictionary<string, Block> Blocks = new(StringComparer.OrdinalIgnoreCase);
        public readonly List<Ent> Entities = [];

        private Pair? _peek;

        private Pair? Next()
        {
            if (_peek is { } p) { _peek = null; return p; }
            string? code = r.ReadLine();
            while (code is not null && code.Trim().Length == 0) code = r.ReadLine();
            string? value = r.ReadLine();
            if (code is null || value is null) return null;
            if (!int.TryParse(code.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int c)) return null;
            Pairs++;
            return new Pair(c, value.TrimEnd('\r'));
        }

        public void Read()
        {
            while (Next() is { } p)
            {
                if (p.Code != 0 || p.Value.Trim() != "SECTION") continue;
                if (Next() is not { Code: 2 } name) continue;
                switch (name.Value.Trim())
                {
                    case "TABLES": ReadTables(); break;
                    case "BLOCKS": ReadBlocks(); break;
                    case "ENTITIES": Entities.AddRange(ReadEntities(out _)); break;
                }
            }
        }

        /// <summary>Reads one entity (from just after its 0 code) up to the next 0 code.</summary>
        private Ent ReadEnt(string type)
        {
            var e = new Ent(type);
            while (Next() is { } p)
            {
                if (p.Code == 0) { _peek = p; break; }
                e.Pairs.Add(p.Code is 1 or 3 ? p : p with { Value = p.Value.Trim() });
            }
            return e;
        }

        /// <summary>Entities up to ENDSEC or ENDBLK, with the ATTRIBs of an INSERT and the VERTEXes of a POLYLINE folded in.</summary>
        private List<Ent> ReadEntities(out Ent? end)
        {
            var list = new List<Ent>();
            end = null;
            while (Next() is { } p)
            {
                if (p.Code != 0) continue;
                string type = p.Value.Trim();
                if (type is "ENDSEC" or "ENDBLK") { end = ReadEnt(type); break; }
                Ent e = ReadEnt(type);
                if ((type == "INSERT" && e.I(66) == 1) || type == "POLYLINE")
                {
                    e.Attribs = [];
                    while (Next() is { } q)
                    {
                        if (q.Code != 0) continue;
                        string t = q.Value.Trim();
                        Ent sub = ReadEnt(t);
                        if (t == "SEQEND") break;
                        if (t is "ATTRIB" or "VERTEX") e.Attribs.Add(sub);
                        else { list.Add(sub); break; }  // no SEQEND: the next entity already started
                    }
                }
                list.Add(e);
            }
            return list;
        }

        private void ReadTables()
        {
            while (Next() is { } p)
            {
                if (p.Code != 0) continue;
                string type = p.Value.Trim();
                if (type == "ENDSEC") return;
                if (type != "LAYER") continue;
                Ent e = ReadEnt(type);
                string name = e.S(2) ?? "0";
                int color = e.I(62, 7);
                Layers[name] = new Layer(name, Math.Abs(color), e.S(420), color < 0 || (e.I(70) & 1) != 0);
            }
        }

        private void ReadBlocks()
        {
            while (Next() is { } p)
            {
                if (p.Code != 0) continue;
                string type = p.Value.Trim();
                if (type == "ENDSEC") return;
                if (type != "BLOCK") continue;
                Ent head = ReadEnt(type);
                List<Ent> ents = ReadEntities(out Ent? end);
                string name = head.S(2) ?? "";
                if (name.Length > 0) Blocks[name] = new Block(name, head.D(10), head.D(20), ents);
                if (end?.Type == "ENDSEC") return;
            }
        }
    }

    /// <summary>x' = a x + c y + e, y' = b x + d y + f.</summary>
    private readonly record struct Affine(double A, double B, double C, double D, double E, double F)
    {
        public static readonly Affine Identity = new(1, 0, 0, 1, 0, 0);
        public (double X, double Y) Apply(double x, double y) => (A * x + C * y + E, B * x + D * y + F);
        public (double X, double Y) Vector(double x, double y) => (A * x + C * y, B * x + D * y);
        /// <summary>This after <paramref name="inner"/>: apply inner first.</summary>
        public Affine Times(Affine inner) => new(
            A * inner.A + C * inner.B, B * inner.A + D * inner.B,
            A * inner.C + C * inner.D, B * inner.C + D * inner.D,
            A * inner.E + C * inner.F + E, B * inner.E + D * inner.F + F);
        public double Scale => Math.Sqrt(Math.Abs(A * D - B * C));
    }

    /// <summary>Spreads blocks out and turns every entity into lines and texts in world coordinates.</summary>
    private sealed class Builder(Reader doc)
    {
        public readonly List<WorkingLayoutPath> Paths = [];
        public readonly List<WorkingLayoutText> Texts = [];
        public readonly List<string> LayerNames = [];
        private readonly Dictionary<string, int> _layerIndex = new(StringComparer.OrdinalIgnoreCase);
        public double MinX = double.MaxValue, MinY = double.MaxValue, MaxX = double.MinValue, MaxY = double.MinValue;

        private int LayerOf(string name)
        {
            if (_layerIndex.TryGetValue(name, out int i)) return i;
            _layerIndex[name] = LayerNames.Count;
            LayerNames.Add(name);
            return LayerNames.Count - 1;
        }

        private void Grow(double x, double y)
        {
            if (x < MinX) MinX = x;
            if (x > MaxX) MaxX = x;
            if (y < MinY) MinY = y;
            if (y > MaxY) MaxY = y;
        }

        /// <param name="blockLayer">The layer of the INSERT a block is drawn by: entities of layer 0 take it.</param>
        /// <param name="blockColor">The colour of that INSERT: BYBLOCK entities take it.</param>
        public void Draw(List<Ent> ents, Affine m, string? blockLayer, string? blockColor, int depth)
        {
            foreach (Ent e in ents)
            {
                if (Paths.Count > MaxPaths) throw new ArgumentException("The drawing is too large to show (over 600,000 lines).");
                if (e.I(67) == 1) continue;  // paper space
                string layer = e.S(8) ?? "0";
                if (blockLayer is not null && layer == "0") layer = blockLayer;
                if (doc.Layers.TryGetValue(layer, out Layer? lay) && lay.Hidden) continue;
                if (e.I(60) == 1) continue;  // invisible
                string? color = ColorOf(e, lay, blockColor);
                int l = LayerOf(layer);
                switch (e.Type)
                {
                    case "LINE": Add(l, color, m, [e.D(10), e.D(20), e.D(11), e.D(21)], false); break;
                    case "LWPOLYLINE": LwPolyline(e, l, color, m); break;
                    case "POLYLINE": Polyline(e, l, color, m); break;
                    case "CIRCLE": Arc(e, l, color, m, 0, 360, true); break;
                    case "ARC": Arc(e, l, color, m, e.D(50), e.D(51), false); break;
                    case "ELLIPSE": Ellipse(e, l, color, m); break;
                    case "SPLINE": Spline(e, l, color, m); break;
                    case "SOLID" or "TRACE" or "3DFACE": Solid(e, l, color, m); break;
                    case "TEXT": Text(e, l, color, m, mtext: false); break;
                    case "MTEXT": Text(e, l, color, m, mtext: true); break;
                    case "INSERT": Insert(e, layer, color, m, depth); break;
                    case "DIMENSION":
                        if (depth < MaxDepth && e.S(2) is { } dim && doc.Blocks.TryGetValue(dim, out Block? db))
                            Draw(db.Ents, m, layer, color, depth + 1);
                        break;
                }
            }
        }

        private string? ColorOf(Ent e, Layer? lay, string? blockColor)
        {
            if (e.S(420) is { } tc) return TrueColor(tc);
            int aci = e.I(62, 256);
            if (aci == 0) return blockColor;
            if (aci == 256) return lay is null ? null : lay.True is { } lt ? TrueColor(lt) : Aci(lay.Color);
            return Aci(aci);
        }

        /// <summary>Points in the entity's own coordinates (OCS); a drawing mirrored by extrusion 0,0,-1 has x turned over.</summary>
        private static double Ocs(Ent e) => e.D(230, 1) < 0 ? -1 : 1;

        private void Add(int l, string? c, Affine m, IReadOnlyList<double> pts, bool closed)
        {
            if (pts.Count < 4) return;
            var p = new List<double>(pts.Count);
            for (int i = 0; i < pts.Count; i += 2)
            {
                (double x, double y) = m.Apply(pts[i], pts[i + 1]);
                double rx = R(x), ry = R(y);
                if (p.Count >= 2 && p[^2] == rx && p[^1] == ry) continue;  // the same point once rounded
                Grow(x, y);
                p.Add(rx);
                p.Add(ry);
            }
            if (p.Count < 4) return;
            Paths.Add(new WorkingLayoutPath(l, c, [.. p], closed ? 1 : null));
        }

        /// <summary>
        /// Pieces for an arc: no further than <see cref="Tolerance"/> (world units) from the true arc, so a small circle is a
        /// few lines and a big one a smooth curve; at most 64 pieces a full turn.
        /// </summary>
        private static int Segments(double sweep, double worldRadius, int min)
        {
            double step = worldRadius <= Tolerance ? Math.PI / 2 : 2 * Math.Acos(1 - Tolerance / worldRadius);
            step = Math.Max(step, Math.PI / 32);
            return Math.Max(min, (int)Math.Ceiling(sweep / step));
        }

                /// <summary>The points of a vertex list with bulges (arc segments) spread out.</summary>
        private static List<double> Bulged(List<(double X, double Y, double Bulge)> v, bool closed, double scale)
        {
            var pts = new List<double>();
            int n = v.Count;
            for (int i = 0; i < n; i++)
            {
                pts.Add(v[i].X);
                pts.Add(v[i].Y);
                if (i == n - 1 && !closed) break;
                var a = v[i];
                var b = v[(i + 1) % n];
                if (Math.Abs(a.Bulge) < 1e-9) continue;
                double theta = 4 * Math.Atan(a.Bulge);
                double dx = b.X - a.X, dy = b.Y - a.Y, chord = Math.Sqrt(dx * dx + dy * dy);
                if (chord < 1e-12) continue;
                double radius = chord / (2 * Math.Sin(theta / 2));
                double mx = (a.X + b.X) / 2, my = (a.Y + b.Y) / 2;
                double h = radius * Math.Cos(theta / 2);
                double cx = mx - h * dy / chord, cy = my + h * dx / chord;
                double start = Math.Atan2(a.Y - cy, a.X - cx);
                double r = Math.Abs(radius);
                int seg = Segments(Math.Abs(theta), r * scale, 2);
                for (int k = 1; k < seg; k++)
                {
                    double t = start + theta * k / seg;
                    pts.Add(cx + r * Math.Cos(t));
                    pts.Add(cy + r * Math.Sin(t));
                }
            }
            if (closed && n > 1) { pts.Add(v[0].X); pts.Add(v[0].Y); }
            return pts;
        }

        private void LwPolyline(Ent e, int l, string? c, Affine m)
        {
            var v = new List<(double, double, double)>();
            double flip = Ocs(e);
            double? x = null, y = null, bulge = 0;
            foreach (Pair p in e.Pairs)
            {
                if (p.Code == 10)
                {
                    if (x is not null && y is not null) v.Add((x.Value, y.Value, bulge ?? 0));
                    x = Num(p.Value) * flip; y = null; bulge = 0;
                }
                else if (p.Code == 20) y = Num(p.Value);
                else if (p.Code == 42) bulge = Num(p.Value) * flip;
            }
            if (x is not null && y is not null) v.Add((x.Value, y.Value, bulge ?? 0));
            bool closed = (e.I(70) & 1) != 0;
            Add(l, c, m, Bulged(v, closed, m.Scale), closed);
        }

        private void Polyline(Ent e, int l, string? c, Affine m)
        {
            int flags = e.I(70);
            if ((flags & (16 | 64)) != 0) return;  // meshes
            double flip = Ocs(e);
            var v = (e.Attribs ?? [])
                .Where(x => (x.I(70) & 16) == 0)  // spline frame points
                .Select(x => (x.D(10) * flip, x.D(20), x.D(42) * flip)).ToList();
            bool closed = (flags & 1) != 0;
            Add(l, c, m, Bulged(v, closed, m.Scale), closed);
        }

        private void Arc(Ent e, int l, string? c, Affine m, double startDeg, double endDeg, bool full)
        {
            double flip = Ocs(e);
            double cx = e.D(10) * flip, cy = e.D(20), r = e.D(40);
            if (r <= 0) return;
            double a0 = startDeg * Math.PI / 180, a1 = endDeg * Math.PI / 180;
            if (flip < 0) (a0, a1) = (Math.PI - a1, Math.PI - a0);
            double sweep = full ? 2 * Math.PI : a1 - a0;
            while (!full && sweep <= 0) sweep += 2 * Math.PI;
            int seg = Segments(sweep, r * m.Scale, full ? 12 : 3);
            var pts = new List<double>((seg + 1) * 2);
            for (int k = 0; k <= seg; k++)
            {
                double t = a0 + sweep * k / seg;
                pts.Add(cx + r * Math.Cos(t));
                pts.Add(cy + r * Math.Sin(t));
            }
            Add(l, c, m, pts, full);
        }

        private void Ellipse(Ent e, int l, string? c, Affine m)
        {
            double cx = e.D(10), cy = e.D(20), mx = e.D(11), my = e.D(21), ratio = e.D(40, 1);
            double t0 = e.D(41), t1 = e.D(42, 2 * Math.PI);
            double sweep = t1 - t0;
            while (sweep <= 0) sweep += 2 * Math.PI;
            bool full = Math.Abs(sweep - 2 * Math.PI) < 1e-6;
            // centre and major axis are in world coordinates; the minor axis is extrusion × major, times the ratio
            double flip = Ocs(e);
            double nx = -my * ratio * flip, ny = mx * ratio * flip;
            int seg = Segments(sweep, Math.Sqrt(mx * mx + my * my) * m.Scale, full ? 12 : 3);
            var pts = new List<double>((seg + 1) * 2);
            for (int k = 0; k <= seg; k++)
            {
                double t = t0 + sweep * k / seg;
                pts.Add(cx + mx * Math.Cos(t) + nx * Math.Sin(t));
                pts.Add(cy + my * Math.Cos(t) + ny * Math.Sin(t));
            }
            Add(l, c, m, pts, full);
        }

        private void Spline(Ent e, int l, string? c, Affine m)
        {
            int degree = e.I(71, 3);
            double[] knots = e.All(40).ToArray();
            double[] w = e.All(41).ToArray();
            double[] cxs = e.All(10).ToArray(), cys = e.All(20).ToArray();
            int n = Math.Min(cxs.Length, cys.Length);
            bool closed = (e.I(70) & 1) != 0;
            if (n < 2 || knots.Length != n + degree + 1 || degree < 1)
            {
                double[] fx = e.All(11).ToArray(), fy = e.All(21).ToArray();
                int fn = Math.Min(fx.Length, fy.Length);
                double[] xs = fn >= 2 ? fx : cxs, ys = fn >= 2 ? fy : cys;
                int k = fn >= 2 ? fn : n;
                var raw = new List<double>(k * 2);
                for (int i = 0; i < k; i++) { raw.Add(xs[i]); raw.Add(ys[i]); }
                Add(l, c, m, raw, closed);
                return;
            }
            int samples = Math.Clamp(n * 4, 8, 200);
            double u0 = knots[degree], u1 = knots[n];
            var pts = new List<double>((samples + 1) * 2);
            for (int s = 0; s <= samples; s++)
            {
                double u = u0 + (u1 - u0) * s / samples;
                (double x, double y) = DeBoor(degree, knots, cxs, cys, w.Length == n ? w : null, n, u);
                pts.Add(x);
                pts.Add(y);
            }
            Add(l, c, m, pts, closed);
        }

        private static (double, double) DeBoor(int p, double[] t, double[] xs, double[] ys, double[]? w, int n, double u)
        {
            int k = p;
            while (k < n - 1 && u >= t[k + 1]) k++;
            var dx = new double[p + 1];
            var dy = new double[p + 1];
            var dw = new double[p + 1];
            for (int j = 0; j <= p; j++)
            {
                int i = Math.Clamp(j + k - p, 0, n - 1);
                double wi = w?[i] ?? 1;
                dx[j] = xs[i] * wi; dy[j] = ys[i] * wi; dw[j] = wi;
            }
            for (int r = 1; r <= p; r++)
            for (int j = p; j >= r; j--)
            {
                int i = j + k - p;
                double den = t[i + p - r + 1] - t[i];
                double a = den == 0 ? 0 : (u - t[i]) / den;
                dx[j] = (1 - a) * dx[j - 1] + a * dx[j];
                dy[j] = (1 - a) * dy[j - 1] + a * dy[j];
                dw[j] = (1 - a) * dw[j - 1] + a * dw[j];
            }
            return dw[p] == 0 ? (dx[p], dy[p]) : (dx[p] / dw[p], dy[p] / dw[p]);
        }

        private void Solid(Ent e, int l, string? c, Affine m)
        {
            // SOLID corners go 1, 2, 4, 3 around
            double[] p = e.Type == "SOLID"
                ? [e.D(10), e.D(20), e.D(11), e.D(21), e.D(13, e.D(12)), e.D(23, e.D(22)), e.D(12), e.D(22)]
                : [e.D(10), e.D(20), e.D(11), e.D(21), e.D(12), e.D(22), e.D(13, e.D(12)), e.D(23, e.D(22))];
            Add(l, c, m, [.. p, p[0], p[1]], true);
        }

        private void Text(Ent e, int l, string? c, Affine m, bool mtext)
        {
            string raw = mtext ? string.Concat(e.Pairs.Where(p => p.Code is 3 or 1).Select(p => p.Value)) : e.S(1) ?? "";
            string t = mtext ? CleanMText(raw) : CleanText(raw);
            if (t.Trim().Length == 0) return;
            double flip = mtext ? 1 : Ocs(e);  // an MTEXT is placed in world coordinates, a TEXT in its own
            double x = e.D(10) * flip, y = e.D(20), h = e.D(40, 1), rot = e.D(50);
            int? a = null, v = null;
            if (mtext)
            {
                if (e.S(11) is not null && e.S(50) is null) rot = Math.Atan2(e.D(21), e.D(11)) * 180 / Math.PI;
                int at = Math.Clamp(e.I(71, 1), 1, 9);
                int col = (at - 1) % 3, row = (at - 1) / 3;
                a = col == 0 ? null : col;
                v = row switch { 0 => 3, 1 => 2, _ => 1 };
            }
            else
            {
                int ha = e.I(72), va = e.I(73);
                if (ha != 0 || va != 0)
                {
                    double ax = e.D(11) * flip, ay = e.D(21);
                    if (ha is 3 or 5) { x = (x + ax) / 2; y = (y + ay) / 2; a = 1; }
                    else
                    {
                        x = ax; y = ay;
                        a = ha switch { 1 or 4 => 1, 2 => 2, _ => null };
                        if (ha == 4) v = 2;
                    }
                    if (va != 0) v = va;
                }
            }
            if (flip < 0) rot = 180 - rot;
            double rad = rot * Math.PI / 180;
            (double wx, double wy) = m.Apply(x, y);
            (double ux, double uy) = m.Vector(Math.Cos(rad), Math.Sin(rad));
            (double vx, double vy) = m.Vector(-Math.Sin(rad), Math.Cos(rad));
            double height = h * Math.Sqrt(vx * vx + vy * vy);
            if (height <= 0) return;
            Grow(wx, wy);
            Texts.Add(new WorkingLayoutText(l, c, R(wx), R(wy), Math.Round(height, 3), Math.Round(Math.Atan2(uy, ux) * 180 / Math.PI, 2), t, a, v));
        }

        private void Insert(Ent e, string layer, string? color, Affine m, int depth)
        {
            if (depth >= MaxDepth || e.S(2) is not { } name || !doc.Blocks.TryGetValue(name, out Block? blk)) return;
            double flip = Ocs(e);
            double sx = e.D(41, 1) * flip, sy = e.D(42, 1), rot = e.D(50) * Math.PI / 180 * flip;
            int cols = Math.Max(1, e.I(70, 1)), rows = Math.Max(1, e.I(71, 1));
            double dc = e.D(44), dr = e.D(45);
            for (int r = 0; r < rows; r++)
            for (int k = 0; k < cols; k++)
            {
                double ox = k * dc, oy = r * dr;
                double ix = e.D(10) * flip + ox * Math.Cos(rot) - oy * Math.Sin(rot);
                double iy = e.D(20) + ox * Math.Sin(rot) + oy * Math.Cos(rot);
                double cos = Math.Cos(rot), sin = Math.Sin(rot);
                // translate(ins) * rotate * scale * translate(-base)
                var local = new Affine(cos * sx, sin * sx, -sin * sy, cos * sy, 0, 0);
                (double bx, double by) = local.Apply(-blk.BaseX, -blk.BaseY);
                var t = local with { E = bx + ix, F = by + iy };
                Draw(blk.Ents, m.Times(t), layer, color, depth + 1);
            }
            foreach (Ent att in e.Attribs ?? [])
                if ((att.I(70) & 1) == 0 && att.Type == "ATTRIB")
                    Draw([AsText(att)], m, null, color, depth);
        }

        private static Ent AsText(Ent att)
        {
            var t = new Ent("TEXT");
            t.Pairs.AddRange(att.Pairs);
            return t;
        }
    }

    // ---------------------------------------------------------------- texts and colours

    private static string Specials(string s)
    {
        if (!s.Contains("%%")) return s;
        var sb = new StringBuilder(s.Length);
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] == '%' && i + 2 < s.Length && s[i + 1] == '%')
            {
                char k = char.ToLowerInvariant(s[i + 2]);
                string? rep = k switch { 'c' => "Ø", 'd' => "°", 'p' => "±", '%' => "%", 'u' or 'o' or 'k' => "", _ => null };
                if (rep is not null) { sb.Append(rep); i += 2; continue; }
                if (char.IsDigit(s[i + 2]))
                {
                    int j = i + 2;
                    while (j < s.Length && j < i + 5 && char.IsDigit(s[j])) j++;
                    sb.Append((char)int.Parse(s[(i + 2)..j], CultureInfo.InvariantCulture));
                    i = j - 1;
                    continue;
                }
            }
            sb.Append(s[i]);
        }
        return sb.ToString();
    }

    /// <summary>\U+XXXX characters of a text.</summary>
    private static string Unicode(string s)
    {
        if (!s.Contains("\\U+", StringComparison.OrdinalIgnoreCase)) return s;
        var sb = new StringBuilder(s.Length);
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] == '\\' && i + 6 < s.Length &&(s[i + 1] == 'U' || s[i + 1] == 'u') && s[i + 2] == '+'
                && int.TryParse(s.AsSpan(i + 3, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int code))
            {
                sb.Append((char)code);
                i += 6;
                continue;
            }
            sb.Append(s[i]);
        }
        return sb.ToString();
    }

    /// <summary>The plain text of a TEXT: special characters (%%d, %%p, %%c) and \U+ characters turned into what they show.</summary>
    public static string CleanText(string s) => Specials(Unicode(s)).Trim();

    /// <summary>The plain text of an MTEXT: format codes (\f…; \H…; \C…; {…}) dropped, \P a new line, stacks a/b.</summary>
    public static string CleanMText(string s)
    {
        s = Unicode(s);
        var sb = new StringBuilder(s.Length);
        for (int i = 0; i < s.Length; i++)
        {
            char ch = s[i];
            if (ch is '{' or '}') continue;
            if (ch != '\\' || i + 1 >= s.Length) { sb.Append(ch); continue; }
            char k = s[++i];
            switch (k)
            {
                case 'P': sb.Append('\n'); break;
                case '~': sb.Append(' '); break;
                case '\\' or '{' or '}': sb.Append(k); break;
                case 'L' or 'l' or 'O' or 'o' or 'K' or 'k' or 'N': break;
                case 'S':
                {
                    int end = s.IndexOf(';', i);
                    if (end < 0) end = s.Length;
                    string stack = s[(i + 1)..end];
                    sb.Append(stack.Replace('^', '/').Replace('#', '/').Trim());
                    i = end;
                    break;
                }
                case 'A' or 'C' or 'c' or 'F' or 'f' or 'H' or 'h' or 'Q' or 'q' or 'T' or 't' or 'W' or 'w' or 'p' or 'X' or 'x':
                {
                    int end = s.IndexOf(';', i);
                    i = end < 0 ? s.Length : end;
                    break;
                }
                default: sb.Append(k); break;
            }
        }
        string text = Specials(sb.ToString());
        return string.Join('\n', text.Split('\n').Select(line => line.TrimEnd())).Trim('\n', ' ');
    }

    private static string? TrueColor(string v) =>
        int.TryParse(v.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int n) ? "#" + (n & 0xFFFFFF).ToString("x6") : null;

    /// <summary>The colour of an AutoCAD colour index; 7 (white on black, black on white) is null: the screen's foreground.</summary>
    public static string? Aci(int i)
    {
        switch (i)
        {
            case <= 0 or 7 or >= 256: return null;
            case 1: return "#ff0000";
            case 2: return "#ffff00";
            case 3: return "#00ff00";
            case 4: return "#00ffff";
            case 5: return "#0000ff";
            case 6: return "#ff00ff";
            case 8: return "#808080";
            case 9: return "#c0c0c0";
            case >= 250:
                int g = new[] { 0x33, 0x50, 0x69, 0x82, 0xbe, 0xff }[i - 250];
                return $"#{g:x2}{g:x2}{g:x2}";
        }
        // 10..249: hue in steps of 15 degrees; the last digit gives the value (0-1 bright … 8-9 dark), odd ones paler
        double hue = (i / 10 - 1) * 15;
        int digit = i % 10;
        double value = new[] { 1.0, 1.0, 0.8, 0.8, 0.6, 0.6, 0.5, 0.5, 0.3, 0.3 }[digit];
        double sat = digit % 2 == 0 ? 1.0 : 0.5;
        double c = value * sat, x = c * (1 - Math.Abs(hue / 60 % 2 - 1)), mm = value - c;
        (double r, double gg, double b) = hue switch
        {
            < 60 => (c, x, 0d), < 120 => (x, c, 0d), < 180 => (0d, c, x),
            < 240 => (0d, x, c), < 300 => (x, 0d, c), _ => (c, 0d, x),
        };
        int R8(double v) => (int)Math.Round((v + mm) * 255);
        return $"#{R8(r):x2}{R8(gg):x2}{R8(b):x2}";
    }
}
