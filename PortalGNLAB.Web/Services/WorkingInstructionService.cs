using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace PortalGNLAB.Web.Services;

/// <summary>The Work Instructions (작업지도서) of each model: one workbook, one sheet per process.</summary>
public sealed class WorkingInstructionService(string? root = null)
{
    public const long MaxUploadBytes = 100L * 1024 * 1024;

    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web);
    private readonly object gate = new();
    private string Root => root ?? AppStoragePaths.AppFolder("working-instructions");

    // Labels the sheets are read by. Add the Korean / English words of another form here.
    internal static readonly string[] NameLabels = ["Tên công đoạn", "공정명", "Process name"];
    internal static readonly string[] CodeLabels = ["Công đoạn", "공정", "Process"];
    internal static readonly string[] StandardLabels = ["Thời gian tiêu chuẩn", "표준시간", "표준 시간", "Standard time"];
    internal static readonly string[] SpareLabels = ["Thời gian thừa", "여유시간", "여유 시간", "Spare time", "Allowance"];
    internal static readonly string[] CtqLabels = ["Công đoạn CTQ", "CTQ 공정", "CTQ process"];
    internal static readonly string[] PartLabels = ["PART LIST", "부품 리스트", "부품리스트", "Part list"];
    internal static readonly string[] ToolLabels = ["Thiết bị, công cụ", "Thiết bị công cụ", "설비·공구", "설비, 공구", "설비/공구", "Equipment and tools", "Tools"];
    internal static readonly string[] KeyLabels = ["Hạng mục chú ý", "중점 관리", "중점관리", "Key points", "Points to note"];
    internal static readonly string[] StepLabels = ["Bước", "단계", "Step"];
    internal static readonly string[] DateLabels = ["Ngày sửa đổi", "개정일", "Revision date", "Date"];
    internal static readonly string[] PageLabels = ["Trang", "페이지", "Page"];
    internal static readonly string[] RevLabels = ["Mã số sửa đổi", "개정번호", "개정 번호", "Revision no"];
    internal static readonly string[] ContentLabels = ["Tóm tắt nội dung sửa đổi", "개정내용", "개정 내용", "Revision details", "Content"];
    internal static readonly string[] AuthorLabels = ["Lập", "작성", "Prepared by", "Author"];

    // ---- storage ----

    /// <summary>The model's workbook was read and kept; <paramref name="data"/> is the uploaded .xlsx.</summary>
    public InstructionDoc Save(string model, string fileName, Stream data, string user)
    {
        string name = Path.GetFileName(fileName ?? "");
        if (name.Length == 0) name = "instructions.xlsx";
        string ext = Path.GetExtension(name);
        if (!ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) && !ext.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Upload an Excel workbook (.xlsx).");
        Directory.CreateDirectory(Root);
        string tmp = Path.Combine(Root, $"{Guid.NewGuid():N}.tmp");
        try
        {
            using (var fs = File.Create(tmp)) data.CopyTo(fs);
            InstructionDoc parsed = ParseFile(tmp);
            var doc = parsed with { FileName = name, UploadedAt = DateTime.UtcNow, UploadedBy = user, Size = new FileInfo(tmp).Length };
            lock (gate)
            {
                string key = Key(model);
                File.Move(tmp, Path.Combine(Root, key + ".xlsx"), true);
                File.WriteAllText(Path.Combine(Root, key + ".json"), JsonSerializer.Serialize(doc, Json));
            }
            return doc;
        }
        finally { if (File.Exists(tmp)) File.Delete(tmp); }
    }

    public InstructionDoc? Get(string model)
    {
        string path = Path.Combine(Root, Key(model) + ".json");
        lock (gate) return File.Exists(path) ? JsonSerializer.Deserialize<InstructionDoc>(File.ReadAllText(path), Json) : null;
    }

    public bool Delete(string model)
    {
        lock (gate)
        {
            string key = Key(model);
            bool had = false;
            foreach (string ext in new[] { ".json", ".xlsx" })
            {
                string path = Path.Combine(Root, key + ext);
                if (!File.Exists(path)) continue;
                File.Delete(path);
                had = true;
            }
            return had;
        }
    }

    /// <summary>The original workbook as uploaded, with the file name it had.</summary>
    public (Stream Data, string FileName)? OpenFile(string model)
    {
        if (Get(model) is not { } doc) return null;
        string path = Path.Combine(Root, Key(model) + ".xlsx");
        return File.Exists(path) ? (new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read), doc.FileName) : null;
    }

    private static string Key(string model) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(model.Trim().ToLowerInvariant())))[..32].ToLowerInvariant();

    // ---- reading ----

    public static InstructionDoc ParseFile(string path)
    {
        byte[] head = new byte[4];
        using (var fs = File.OpenRead(path)) _ = fs.Read(head, 0, 4);
        if (head[0] == 0xD0 && head[1] == 0xCF && head[2] == 0x11 && head[3] == 0xE0)
            throw new ArgumentException("This workbook is encrypted (DRM) and cannot be opened. Upload a decrypted file.");
        if (head[0] != 'P' || head[1] != 'K')
            throw new ArgumentException("This is not an Excel workbook (.xlsx).");
        try
        {
            using var stripped = WithoutPictures(path);
            using var wb = new XLWorkbook(stripped);
            return Read(wb);
        }
        catch (ArgumentException) { throw; }
        catch (Exception)
        {
            throw new ArgumentException("The workbook could not be opened. If it is encrypted (DRM), upload a decrypted file.");
        }
    }

    /// <summary>A copy of the workbook without pictures and printer settings: they are not read, and a broken drawing link stops the reader.</summary>
    private static MemoryStream WithoutPictures(string path)
    {
        var result = new MemoryStream();
        using (var src = ZipFile.OpenRead(path))
        using (var dst = new ZipArchive(result, ZipArchiveMode.Create, true))
        {
            foreach (ZipArchiveEntry e in src.Entries)
            {
                string n = e.FullName;
                if (n.StartsWith("xl/media/") || n.StartsWith("xl/drawings/") || n.StartsWith("xl/printerSettings/")) continue;
                using Stream from = e.Open();
                using Stream to = dst.CreateEntry(n).Open();
                if (n.StartsWith("xl/worksheets/") && n.EndsWith(".xml") && !n.Contains("/_rels/"))
                    WriteText(to, Regex.Replace(ReadText(from), @"<(drawing|legacyDrawing|legacyDrawingHF|picture)\b[^>]*/>", ""));
                else if (n.StartsWith("xl/worksheets/_rels/"))
                    WriteText(to, Regex.Replace(ReadText(from), @"<Relationship\b[^>]*Type=""[^""]*/(drawing|vmlDrawing|image|printerSettings)""[^>]*/>", ""));
                else from.CopyTo(to);
            }
        }
        result.Position = 0;
        return result;
    }

    private static string ReadText(Stream s) => new StreamReader(s, Encoding.UTF8).ReadToEnd();
    private static void WriteText(Stream s, string text) { using var w = new StreamWriter(s, new UTF8Encoding(false)); w.Write(text); }

    private static InstructionDoc Read(XLWorkbook wb)
    {
        var revisions = new List<InstructionRevision>();
        var processes = new List<InstructionProcess>();
        bool first = true;
        foreach (IXLWorksheet ws in wb.Worksheets)
        {
            var grid = new Grid(ws);
            if (grid.Find(NameLabels, 10) is not null) processes.Add(ReadProcess(ws.Name.Trim(), grid));
            else if (first || revisions.Count == 0) revisions.AddRange(ReadRevisions(grid));
            first = false;
        }
        return new InstructionDoc("", DateTime.MinValue, "", 0, revisions, processes);
    }

    private static List<InstructionRevision> ReadRevisions(Grid g)
    {
        var rows = new List<InstructionRevision>();
        if (g.Find(DateLabels, int.MaxValue) is not { } head) return rows;
        int Col(string[] labels) => g.InRow(head.Row).FirstOrDefault(c => StartsWithAny(c.Text, labels))?.Col ?? 0;
        int date = head.Col, page = Col(PageLabels), rev = Col(RevLabels), content = Col(ContentLabels), author = Col(AuthorLabels);
        for (int r = head.Row + 1; r <= g.LastRow; r++)
        {
            string d = g.Text(r, date), p = page > 0 ? g.Text(r, page) : "", v = rev > 0 ? g.Text(r, rev) : "";
            string c = content > 0 ? g.Text(r, content) : "", a = author > 0 ? g.Text(r, author) : "";
            if (d.Length + p.Length + v.Length + c.Length == 0) break;
            rows.Add(new InstructionRevision(d, p, v, c, a));
        }
        return rows;
    }

    private static InstructionProcess ReadProcess(string sheet, Grid g)
    {
        Cell nameCell = g.Find(NameLabels, 10)!;
        string name = g.Right(nameCell, 6) ?? "";
        string code = "";
        foreach (Cell c in g.All.Where(c => c.Row <= 3))
        {
            string[] lines = c.Raw.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            if (lines.Length >= 2 && CodeLabels.Any(l => lines[0].Equals(l, StringComparison.OrdinalIgnoreCase))) { code = lines[1]; break; }
        }
        double? Time(string[] labels) => g.Find(labels, 10) is { } l ? g.RightNumber(l, 3) : null;
        bool ctq = g.All.Any(c => c.Row <= 3 && StartsWithAny(c.Text, CtqLabels));
        string page = g.All.FirstOrDefault(c => c.Row <= 5 && Regex.IsMatch(c.Text, @"^\d+(-\d+)?\s*/\s*\d+$"))?.Text ?? "";

        // sections: the row of each header; a section runs down to the next header
        Cell? part = g.Find(PartLabels, int.MaxValue), tool = g.Find(ToolLabels, int.MaxValue), key = g.Find(KeyLabels, int.MaxValue);
        int[] starts = new[] { part, tool, key }.Where(c => c is not null).Select(c => c!.Row).ToArray();
        int End(Cell c) => starts.Where(s => s > c.Row).DefaultIfEmpty(g.LastRow + 1).Min() - 1;

        var parts = new List<InstructionPart>();
        if (part is not null)
            foreach (string?[] row in Table(g, part, End(part), 4)) parts.Add(new InstructionPart(row[0] ?? "", row[1] ?? "", row[2] ?? "", row[3] ?? ""));
        var tools = new List<InstructionTool>();
        if (tool is not null)
            foreach (string?[] row in Table(g, tool, End(tool), 3)) tools.Add(new InstructionTool(row[0] ?? "", row[1] ?? "", row[2] ?? ""));
        var keys = new List<string>();
        if (key is not null)
            for (int r = key.Row + 1; r <= End(key); r++) if (g.Text(r, key.Col) is { Length: > 0 } t) keys.Add(t);

        var steps = new List<InstructionStep>();
        var stepRx = new Regex(@"^(?:" + string.Join("|", StepLabels.Select(Regex.Escape)) + @")\s*(\d+)\s*[:：.]?\s*(.*)$", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        foreach (Cell c in g.All.Where(c => c.Row >= 5 && c.Col <= 6).OrderBy(c => c.Row).ThenBy(c => c.Col))
        {
            Match m = stepRx.Match(c.Text);
            if (!m.Success) continue;
            string text = m.Groups[2].Value.Trim();
            if (text.Length == 0) text = g.Right(c, 6 - c.Col) ?? "";
            steps.Add(new InstructionStep(int.Parse(m.Groups[1].Value), text));
        }
        return new InstructionProcess(sheet, code, name, Time(StandardLabels), Time(SpareLabels), page, ctq, parts, tools, steps, keys);
    }

    /// <summary>The rows under a section header: the row after it names the columns (NO first), the rest are the entries with a name.</summary>
    private static List<string?[]> Table(Grid g, Cell head, int end, int width)
    {
        var rows = new List<string?[]>();
        int[] cols = g.InRow(head.Row + 1).Where(c => c.Col >= head.Col).Take(width).Select(c => c.Col).ToArray();
        if (cols.Length < 2) return rows;
        for (int r = head.Row + 2; r <= end; r++)
        {
            string?[] row = new string?[width];
            for (int i = 0; i < cols.Length; i++) row[i] = g.Text(r, cols[i]);
            if (string.IsNullOrEmpty(row[1])) continue;
            rows.Add(row);
        }
        return rows;
    }

    private static bool StartsWithAny(string text, string[] labels) =>
        text.Length > 0 && labels.Any(l => text.StartsWith(l, StringComparison.OrdinalIgnoreCase));

    private sealed record Cell(int Row, int Col, string Raw, string Text, double? Number);

    /// <summary>The filled cells of a sheet by position; merged areas count at their top-left cell.</summary>
    private sealed class Grid
    {
        private readonly Dictionary<(int, int), Cell> cells = [];
        public List<Cell> All { get; } = [];
        public int LastRow { get; }

        public Grid(IXLWorksheet ws)
        {
            foreach (IXLCell c in ws.CellsUsed())
            {
                string raw = c.DataType == XLDataType.DateTime && c.TryGetValue(out DateTime d) ? d.ToString("yyyy-MM-dd") : c.GetFormattedString();
                string text = raw.Trim();
                if (text.Length == 0) continue;
                double? number = c.DataType == XLDataType.Number || c.HasFormula ? c.TryGetValue(out double n) ? n : null : null;
                if (c.DataType == XLDataType.Number && number is { } whole && whole == Math.Floor(whole)) text = ((long)whole).ToString();
                var cell = new Cell(c.Address.RowNumber, c.Address.ColumnNumber, raw, text, number);
                cells[(cell.Row, cell.Col)] = cell;
                All.Add(cell);
                LastRow = Math.Max(LastRow, cell.Row);
            }
        }

        public string Text(int row, int col) => cells.TryGetValue((row, col), out Cell? c) ? c.Text : "";
        public IEnumerable<Cell> InRow(int row) => All.Where(c => c.Row == row).OrderBy(c => c.Col);

        /// <summary>The first cell (top to bottom) whose text begins with one of the labels, in the first <paramref name="maxRow"/> rows.</summary>
        public Cell? Find(string[] labels, int maxRow) =>
            All.Where(c => c.Row <= maxRow && StartsWithAny(c.Text, labels)).OrderBy(c => c.Row).ThenBy(c => c.Col).FirstOrDefault();

        /// <summary>The text of the nearest filled cell to the right of <paramref name="from"/>, within <paramref name="span"/> columns.</summary>
        public string? Right(Cell from, int span) =>
            InRow(from.Row).FirstOrDefault(c => c.Col > from.Col && c.Col <= from.Col + span)?.Text;

        public double? RightNumber(Cell from, int span)
        {
            Cell? c = InRow(from.Row).FirstOrDefault(c => c.Col > from.Col && c.Col <= from.Col + span && c.Number is not null);
            return c?.Number is { } v ? Math.Round(v, 3) : null;
        }
    }
}

public sealed record InstructionDoc(string FileName, DateTime UploadedAt, string UploadedBy, long Size,
    List<InstructionRevision> Revisions, List<InstructionProcess> Processes);
public sealed record InstructionRevision(string Date, string Page, string Rev, string Content, string Author);
public sealed record InstructionProcess(string Sheet, string Code, string Name, double? StandardTime, double? SpareTime, string Page, bool Ctq,
    List<InstructionPart> Parts, List<InstructionTool> Tools, List<InstructionStep> Steps, List<string> KeyPoints);
public sealed record InstructionPart(string No, string Name, string Spec, string Qty);
public sealed record InstructionTool(string No, string Name, string Note);
public sealed record InstructionStep(int No, string Text);
