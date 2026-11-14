using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace PortalGNLAB.Web.Services;

/// <summary>
/// A process of the line: its number (as written, "01"), name, people, standard time (seconds), machines, testers,
/// materials (bond / nozzle), a note and its jigs (each with drawing number, version and quantity). Machines and testers
/// are one per line. People and standard time are no longer shown. Jig, JigVer and JigQty are the jigs as text lines of
/// lists saved before Jigs: read into Jigs and then left "".
/// </summary>
public sealed record WorkingProcessRow(
    string Id, string No, string Process, int? Workers, double? StdTime, string Machine, string Jig, string Tester,
    string Material, string Note, string JigVer = "", string JigQty = "", List<WorkingJig>? Jigs = null);

/// <summary>A jig of a process: its name, drawing number (도번), version and quantity, as written ("" when not given).</summary>
public sealed record WorkingJig(string Name, string DrawingNo, string Ver, string Qty);

/// <summary>
/// A machine, jig or tester of a JIG / Machine / Tester list workbook: its name, the text shown in a Process List cell
/// (a machine with its quantity, a tester with its model), the process it is written for, its process numbers, the note
/// of its row and, for a jig, its version, quantity and drawing number (shown in their own columns).
/// </summary>
public sealed record WorkingListItem(string Id, string Kind, string Name, string Label, string Process, List<int> ProcessNos, string Note,
    string Ver = "", string Qty = "", string DrawingNo = "")
{
    public const string Machine = "machine", Jig = "jig", Tester = "tester";
}

public sealed record WorkingProcessList(List<WorkingProcessRow> Rows, int Revision, string UpdatedBy, DateTime? UpdatedAt, string Source);

/// <summary>
/// The Process List (공정편성 LIST) of each model, kept as JSON under &lt;data&gt;\PortalGNLAB\working-process-list. Rows are
/// typed in, or read from a JIG / Machine / Tester list workbook (its Machine List, JIG LIST and Tester List sheets).
/// </summary>
public sealed partial class WorkingProcessListService(string? root = null)
{
    public const long MaxUploadBytes = 50L * 1024 * 1024;
    public const int MaxRows = 500;
    public const int MaxJigs = 100;

    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web);
    private readonly Lock _gate = new();

    private string Root => root ?? AppStoragePaths.AppFolder("working-process-list");

    private string FileOf(string modelName)
    {
        byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes(modelName.Trim().ToLowerInvariant()));
        return Path.Combine(Root, Convert.ToHexString(hash, 0, 16).ToLowerInvariant() + ".json");
    }

    public WorkingProcessList Get(string modelName)
    {
        string path = FileOf(modelName);
        lock (_gate)
        {
            WorkingProcessList? list = File.Exists(path) ? JsonSerializer.Deserialize<WorkingProcessList>(File.ReadAllText(path), Json) : null;
            return list is null ? Empty : list with { Rows = list.Rows.Select(Upgrade).ToList() };
        }
    }

    /// <summary>
    /// A row with its jigs in Jigs: the jig text lines of a list saved before (one per line with JigVer / JigQty lines, or
    /// "NAME (V1) ×6" labels, older still joined with ", ") are read into Jigs and left "".
    /// </summary>
    public static WorkingProcessRow Upgrade(WorkingProcessRow r)
    {
        if (r.Jigs is { Count: > 0 } || string.IsNullOrWhiteSpace(r.Jig))
            return r with { Jig = "", JigVer = "", JigQty = "", Jigs = r.Jigs ?? [] };
        string[] names = r.Jig.Replace("\r\n", "\n").Split(r.Jig.Contains('\n') ? "\n" : ", ");
        string[] vers = (r.JigVer ?? "").Replace("\r\n", "\n").Split('\n'), qtys = (r.JigQty ?? "").Replace("\r\n", "\n").Split('\n');
        var jigs = new List<WorkingJig>();
        for (int i = 0; i < names.Length; i++)
        {
            string name = names[i].Trim(), ver = i < vers.Length ? vers[i].Trim() : "", qty = i < qtys.Length ? qtys[i].Trim() : "";
            if (name.Length == 0) continue;
            if (ver.Length == 0 && qty.Length == 0 && OldJigLabel().Match(name) is { Success: true } m)
                (name, ver, qty) = (m.Groups[1].Value.Trim(), m.Groups[2].Value, m.Groups[3].Value);
            jigs.Add(new WorkingJig(name, "", ver, qty));
        }
        return r with { Jig = "", JigVer = "", JigQty = "", Jigs = jigs };
    }

    /// <summary>"ENCLOSURE JIG (V1) ×6", "PLATE TRAY ×4", "SPK JIG (V0)": the labels jigs had before their own columns.</summary>
    [GeneratedRegex(@"^(.*?)(?:\s*\((V[^)]*)\))?(?:\s*×\s*(\S+))?$")]
    private static partial Regex OldJigLabel();

    private static WorkingProcessList Empty => new([], 0, "", null, "");

    /// <summary>Keeps the rows (cleaned, given ids) when <paramref name="revision"/> is the one stored; else a conflict.</summary>
    public WorkingProcessList Save(string modelName, IEnumerable<WorkingProcessRow?> rows, int revision, string user, string? source = null)
    {
        var clean = Clean(rows);
        lock (_gate)
        {
            WorkingProcessList now = Get(modelName);
            if (revision != now.Revision)
                throw new WorkingConflictException("Someone else saved this Process List. Reload it and make your change again.");
            var saved = new WorkingProcessList(clean, now.Revision + 1, user, DateTime.UtcNow, source ?? now.Source);
            string path = FileOf(modelName);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path + ".tmp", JsonSerializer.Serialize(saved, Json));
            File.Move(path + ".tmp", path, overwrite: true);
            return saved;
        }
    }

    public static List<WorkingProcessRow> Clean(IEnumerable<WorkingProcessRow?> rows)
    {
        var list = new List<WorkingProcessRow>();
        foreach (WorkingProcessRow? r in rows)
        {
            if (r is null) continue;
            static string T(string? s, int max) => (s ?? "").Replace("\r\n", "\n").Trim() is var t && t.Length > max ? t[..max] : t;
            List<WorkingJig> jigs = (Upgrade(r).Jigs ?? [])
                .Where(j => j is not null)
                .Select(j => new WorkingJig(T(j.Name, 200), T(j.DrawingNo, 100), T(j.Ver, 40), T(j.Qty, 40)))
                .Where(j => j.Name.Length + j.DrawingNo.Length + j.Ver.Length + j.Qty.Length > 0)
                .Take(MaxJigs).ToList();
            var row = new WorkingProcessRow(
                string.IsNullOrWhiteSpace(r.Id) ? Guid.NewGuid().ToString("N")[..12] : T(r.Id, 40),
                T(r.No, 20), T(r.Process, 200), r.Workers is >= 0 and <= 10_000 ? r.Workers : null,
                r.StdTime is { } st && double.IsFinite(st) && st >= 0 ? Math.Round(st, 3) : null,
                T(r.Machine, 1000), "", T(r.Tester, 1000), T(r.Material, 500), T(r.Note, 1000), "", "", jigs);
            if (row.No.Length + row.Process.Length + row.Machine.Length + row.Tester.Length + row.Material.Length + row.Note.Length
                + jigs.Count == 0
                && row.Workers is null && row.StdTime is null) continue;
            list.Add(row);
        }
        if (list.Count > MaxRows) throw new ArgumentException($"A Process List has at most {MaxRows} rows.");
        return list;
    }

    // ---------------------------------------------------------------- reading a JIG / Machine / Tester list workbook

    /// <summary>
    /// The processes of a JIG / Machine / Tester list workbook by process number: Machine List rows give the name and the
    /// machines, JIG LIST rows (with process ranges such as "1) 1 ~ 5 / 2) 7 ~ 9") the jigs, Tester List rows the testers.
    /// </summary>
    public static List<WorkingProcessRow> ReadWorkbook(Stream xlsx) => JoinByNumber(ReadItems(xlsx));

    /// <summary>The items of a list workbook joined into one row per process number (numbered names from the JIG LIST).</summary>
    public static List<WorkingProcessRow> JoinByNumber(IReadOnlyList<WorkingListItem> items)
    {
        var byNo = new SortedDictionary<int, Proc>();
        Proc At(int no) => byNo.TryGetValue(no, out Proc? p) ? p : byNo[no] = new Proc();
        foreach (WorkingListItem item in items)
        {
            if (item.Kind == WorkingListItem.Jig)
                // "1)  1. UPPER ARRAY ~ 5. SPK ASSY": names of the numbered processes
                foreach (Match m in NumberedName().Matches(item.Process))
                    At(int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture)).JigName ??= m.Groups[2].Value.Trim();
            foreach (int no in item.ProcessNos)
            {
                Proc p = At(no);
                if (item.Kind != WorkingListItem.Jig) p.Name ??= NullIfEmpty(item.Process);
                p.Items.Add(item);
                if (item.Note.Length > 0) p.Notes.Add(item.Note);
            }
        }
        // a tester without process numbers goes to the processes of its name ("AIR LEAK TEST")
        foreach (WorkingListItem t in items.Where(i => i.Kind == WorkingListItem.Tester && i.ProcessNos.Count == 0 && i.Process.Length > 0))
            foreach (Proc p in byNo.Values.Where(p => string.Equals(p.Name ?? p.JigName, t.Process, StringComparison.OrdinalIgnoreCase)))
                p.Items.Add(t);
        if (byNo.Count == 0)
            throw new ArgumentException("No Machine List, JIG LIST or Tester List sheet with process numbers was found in the workbook.");
        return byNo.Select(kv => Row(Guid.NewGuid().ToString("N")[..12], kv.Key.ToString("00", CultureInfo.InvariantCulture),
            kv.Value.Name ?? kv.Value.JigName ?? "", kv.Value.Items, string.Join(" / ", kv.Value.Notes.Distinct()))).ToList();
    }

    /// <summary>
    /// A row of a process and its items: machines and testers one per line, each jig with its drawing number, version and
    /// quantity; no people or standard time.
    /// </summary>
    public static WorkingProcessRow Row(string id, string no, string process, IEnumerable<WorkingListItem> items, string note)
    {
        List<WorkingListItem> its = items.ToList();
        string Of(string kind) => string.Join("\n", its.Where(i => i.Kind == kind).Select(i => i.Label).Distinct());
        List<WorkingJig> jigs = its.Where(i => i.Kind == WorkingListItem.Jig)
            .Select(i => new WorkingJig(Cap(i.Label, 200), Cap(i.DrawingNo, 100), Cap(i.Ver, 40), Cap(i.Qty, 40))).Distinct().Take(MaxJigs).ToList();
        return new WorkingProcessRow(id, no, Cap(process, 200), null, null, Cap(Of(WorkingListItem.Machine), 1000),
            "", Cap(Of(WorkingListItem.Tester), 1000), "", Cap(note, 1000), "", "", jigs);
    }

    private static string Cap(string s, int max) => s.Length > max ? s[..max] : s;

    /// <summary>
    /// Every machine (Machine List), jig of the main line (JIG LIST) and tester (Tester List) of a list workbook with the
    /// process it is written for and its process numbers; ids M1.., J1.., T1.. in the order of the sheets.
    /// </summary>
    public static List<WorkingListItem> ReadItems(Stream xlsx)
    {
        XLWorkbook book;
        try { book = new XLWorkbook(xlsx); }
        catch (Exception ex) when (ex is not ArgumentException)
        {
            throw new ArgumentException("The workbook could not be opened. If it is protected (DRM), upload a decrypted copy.");
        }
        using (book)
        {
            var items = new List<WorkingListItem>();
            void Add(string kind, string name, string label, string process, string nos, string note, string ver = "", string qty = "", string drawing = "") =>
                items.Add(new WorkingListItem(kind[..1].ToUpperInvariant() + (items.Count(i => i.Kind == kind) + 1),
                    kind, name, label, process, Numbers(nos), note, ver, qty, drawing));

            foreach (IXLWorksheet ws in book.Worksheets)
            {
                // a Machine List has a process column; a drawing transfer list ("JIG & 설비명") has none
                if (Header(ws, "설비명", "MACHINE NAME") is { } mh && Col(mh, "사용 공정", "사용공정", "USING THE PROCESS") > 0)
                {
                    int cNo = Col(mh, "사용공정 번호", "PROCESS NO"), cName = Col(mh, "설비명", "MACHINE NAME"),
                        cProc = Col(mh, "사용 공정", "USING THE PROCESS)"), cQty = Col(mh, "소요 수량", "REQUIRED QUANTITY"),
                        cNote = Col(mh, "비고", "NOTE");
                    foreach (IXLRow row in Rows(ws, mh))
                    {
                        string machine = Text(row, cName);
                        if (machine.Length == 0) continue;
                        string qty = Text(row, cQty);
                        Add(WorkingListItem.Machine, machine, qty is "" or "1" ? machine : $"{machine} ×{qty}",
                            OneLine(Text(row, cProc)), Text(row, cNo), OneLine(Text(row, cNote)));
                    }
                }
                else if (Header(ws, "JIG 및 소모성", "JIG & CONSUMABLE") is { } jh)
                {
                    int cLine = Col(jh, "LINE"), cProc = Col(jh, "공정", "(PROCESS)"), cJig = Col(jh, "JIG 및 소모성", "JIG & CONSUMABLE"),
                        cNo = Col(jh, "사용공정 번호", "PROCESS NO"), cVer = Col(jh, "MSM-", "VER");
                    int cMass = Col(jh, "MASS JIG");
                    // "도번(Drawing No)" stands over a "Ver." and a "Drawing No" column on the row below the header
                    int cDraw = ws.Row(jh.RowNumber() + 1).CellsUsed().FirstOrDefault(c => Upper(c).Contains("DRAWING NO", StringComparison.Ordinal))
                        ?.Address.ColumnNumber ?? Col(jh, "DRAWING NO", "도번");
                    foreach (IXLRow row in Rows(ws, jh))
                    {
                        string jig = OneLine(Text(row, cJig));
                        string line = Text(row, cLine).ToUpperInvariant();
                        if (jig.Length == 0 || (line.Length > 0 && line != "MAIN")) continue;
                        string ver = cVer > 0 ? Text(row, cVer + 1) : "";
                        string mass = cMass > 0 ? Text(row, cMass) : "";
                        if (ver.StartsWith("20", StringComparison.Ordinal)) ver = "";
                        if (mass is "0" or "-") mass = "";
                        string drawing = OneLine(Text(row, cDraw));
                        if (drawing is "-") drawing = "";
                        Add(WorkingListItem.Jig, jig, jig, OneLine(Text(row, cProc)), Text(row, cNo), "", ver, mass, drawing);
                    }
                }
                else if (Header(ws, "계측기명", "TESTER NAME") is { } th)
                {
                    int cName = Col(th, "계측기명", "TESTER NAME"), cModel = Col(th, "모델명", "MODEL NAME"),
                        cProc = Col(th, "사용 공정", "USING THE PROCESS)"), cNo = Col(th, "사용공정 번호", "PROCESS NO");
                    foreach (IXLRow row in Rows(ws, th))
                    {
                        string name = OneLine(Text(row, cName));
                        if (name.Length == 0) continue;
                        string model = OneLine(Text(row, cModel));
                        Add(WorkingListItem.Tester, name, model.Length > 0 ? $"{name} ({model})" : name,
                            OneLine(Text(row, cProc)), Text(row, cNo), "");
                    }
                }
            }
            return items;
        }
    }

    private sealed class Proc
    {
        public string? Name, JigName;
        public readonly List<WorkingListItem> Items = [];
        public readonly List<string> Notes = [];
    }

    [GeneratedRegex(@"(\d{1,3})\.\s*([A-Za-z][^~/\n\r]*?)\s*(?=~|/|\r|\n|$|\d{1,3}\))")]
    private static partial Regex NumberedName();

    /// <summary>Process numbers of a cell: "4", "4 ~ 5", "1) 1 ~ 5\n2) 7 ~ 9", "36, 37, 40" (list markers "1)" left out).</summary>
    public static List<int> Numbers(string cell)
    {
        var list = new List<int>();
        string s = Regex.Replace(cell, @"\d+\)", " ");
        foreach (Match m in Regex.Matches(s, @"(\d{1,3})(?:\s*~\s*(\d{1,3}))?"))
        {
            int lo = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            int hi = m.Groups[2].Success ? int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture) : lo;
            if (hi < lo || hi - lo > 200) hi = lo;
            for (int n = lo; n <= hi; n++) if (n > 0 && !list.Contains(n)) list.Add(n);
        }
        return list;
    }

    /// <summary>The first row (of the first 15) with a cell containing one of the texts: the header of a list sheet.</summary>
    private static IXLRow? Header(IXLWorksheet ws, params string[] texts)
    {
        foreach (IXLRow row in ws.RowsUsed().Take(15))
            if (row.CellsUsed().Any(c => texts.Any(t => Upper(c).Contains(t, StringComparison.Ordinal)))) return row;
        return null;
    }

    /// <summary>The column of the header cell containing one of the texts, or 0.</summary>
    private static int Col(IXLRow header, params string[] texts) =>
        header.CellsUsed().FirstOrDefault(c => texts.Any(t => Upper(c).Contains(t, StringComparison.Ordinal)))?.Address.ColumnNumber ?? 0;

    private static string Upper(IXLCell c) => c.GetFormattedString().Replace('\n', ' ').ToUpperInvariant();

    private static IEnumerable<IXLRow> Rows(IXLWorksheet ws, IXLRow header) =>
        ws.RowsUsed().Where(r => r.RowNumber() > header.RowNumber());

    private static string Text(IXLRow row, int col)
    {
        if (col <= 0) return "";
        IXLCell cell = row.Cell(col);
        return (cell.IsMerged() ? cell.MergedRange().FirstCell() : cell).GetFormattedString().Trim();
    }

    private static string OneLine(string s) => Regex.Replace(s, @"\s*[\r\n]+\s*", " / ").Trim(' ', '/');

    private static string? NullIfEmpty(string s) => s.Length == 0 ? null : s;
}
