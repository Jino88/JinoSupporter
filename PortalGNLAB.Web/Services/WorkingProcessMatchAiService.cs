using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace PortalGNLAB.Web.Services;

/// <summary>A process of the order the items are matched to: its place (from 1), code, name, standard time, people and tools.</summary>
public sealed record MatchProcess(int Seq, string Code, string Name, double? StdTime, int? Workers, List<string> Tools);

/// <summary>A process of the AI answer: a process of the order (Seq) or one it adds after the one before (Seq 0), and its item ids.</summary>
public sealed record MatchAnswerRow(int Seq, string Process, List<string> Items, string Note);

public sealed record MatchAnswer(List<MatchAnswerRow> Rows, string Note);

/// <summary>
/// The rows made by AI, in process order and not saved: Basis is where the order came from ("instructions" for the
/// Work Instructions, "process-list" for the current Process List, "workbook" for the process numbers of the workbook),
/// Unmatched the items put on no process.
/// </summary>
public sealed record ProcessMatchResult(List<WorkingProcessRow> Rows, string Note, List<string> Unmatched, string Basis, string Model);

/// <summary>
/// Working > Process List > AI match: the machines, jigs and testers of a JIG / Machine / Tester list workbook are put
/// on the processes of the model in process order by the Codex API. The order is that of the Work Instructions when
/// the model has them, else the current Process List, else the process numbers of the workbook. Nothing is saved here;
/// the client shows the rows to check and keeps them with Save.
/// </summary>
public sealed class WorkingProcessMatchAiService(
    CodexApiService codex, WorkingInstructionService instructions, WorkingProcessListService lists)
{
    public const string Effort = "medium";
    public const int MaxItems = 1000;

    private static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public async Task<ProcessMatchResult> MatchAsync(string modelName, IReadOnlyList<WorkingListItem> items, CancellationToken ct = default)
    {
        if (items.Count == 0)
            throw new ArgumentException("No Machine List, JIG LIST or Tester List rows were found in the workbook.");
        if (items.Count > MaxItems) throw new ArgumentException($"The workbook has more than {MaxItems} machines, jigs and testers.");
        if (!codex.IsConfigured) throw new ArgumentException("Codex API is not configured. Set the API key in Setting.");

        (string basis, List<MatchProcess> processes) = BasisOf(instructions.Get(modelName), lists.Get(modelName), items);
        JsonArray content = BuildContent(modelName, basis, processes, items);

        string model = codex.Model;
        CodexTextResult result = await codex.GenerateAsync(Instructions, content, model, Effort, 32000, ct);
        codex.RecordUsage(result.Usage);

        MatchAnswer answer = ParseAnswer(result.Text);
        (List<WorkingProcessRow> rows, List<string> unmatched) = Build(answer, processes, items);
        return new ProcessMatchResult(rows, Limit(answer.Note, 1000), unmatched, basis, $"{model} · {Effort}");
    }

    /// <summary>The process order: the Work Instructions sheets, else the Process List rows, else the workbook's process numbers.</summary>
    public static (string Basis, List<MatchProcess> Processes) BasisOf(
        InstructionDoc? doc, WorkingProcessList? list, IReadOnlyList<WorkingListItem> items)
    {
        if (doc is { Processes.Count: > 0 })
            return ("instructions", doc.Processes.Select((p, i) => new MatchProcess(i + 1, p.Code,
                p.Name.Length > 0 ? p.Name : p.Sheet.Trim(), p.StandardTime, null,
                p.Tools.Select(t => t.Name).Where(n => n.Length > 0).ToList())).ToList());
        if (list is { Rows.Count: > 0 })
            return ("process-list", list.Rows.Select((r, i) => new MatchProcess(i + 1, r.No, r.Process, r.StdTime, r.Workers, [])).ToList());
        return ("workbook", WorkingProcessListService.JoinByNumber(items)
            .Select((r, i) => new MatchProcess(i + 1, r.No, r.Process, null, null, [])).ToList());
    }

    /// <summary>The AI input: the process order, the workbook's own numbered processes and the items, as JSON.</summary>
    public static JsonArray BuildContent(string modelName, string basis, IReadOnlyList<MatchProcess> processes, IReadOnlyList<WorkingListItem> items)
    {
        List<object> numbered;
        try
        {
            numbered = WorkingProcessListService.JoinByNumber(items).Select(r => (object)new { no = r.No, name = r.Process }).ToList();
        }
        catch (ArgumentException) { numbered = []; }

        var text = new System.Text.StringBuilder();
        text.AppendLine($"Model name: {modelName}");
        text.AppendLine($"The process order comes from: {basis switch
        {
            "instructions" => "the Work Instructions (one sheet per process, in line order)",
            "process-list" => "the current Process List of the model",
            _ => "the process numbers of the workbook",
        }}");
        text.AppendLine();
        text.AppendLine("=== PROCESSES IN ORDER (JSON) ===");
        text.AppendLine(JsonSerializer.Serialize(processes.Select(p => new { seq = p.Seq, code = p.Code, name = p.Name, tools = p.Tools }), Json));
        text.AppendLine();
        text.AppendLine("=== PROCESS NUMBERS OF THE WORKBOOK (JSON; numbers of the line layout, names as the workbook writes them) ===");
        text.AppendLine(JsonSerializer.Serialize(numbered, Json));
        text.AppendLine();
        text.AppendLine("=== ITEMS OF THE WORKBOOK (JSON) ===");
        text.AppendLine(JsonSerializer.Serialize(items.Select(i => new { id = i.Id, kind = i.Kind, name = i.Name, process = i.Process, processNos = i.ProcessNos }), Json));
        return [new JsonObject { ["type"] = "input_text", ["text"] = text.ToString() }];
    }

    /// <summary>Reads the AI answer leniently: the outermost {...}, rows with a seq, a process name and item ids.</summary>
    public static MatchAnswer ParseAnswer(string? text)
    {
        const string fail = "The AI answer could not be read. Try again.";
        string s = text ?? "";
        int a = s.IndexOf('{'), b = s.LastIndexOf('}');
        if (a < 0 || b <= a) throw new ArgumentException(fail);
        try
        {
            using JsonDocument doc = JsonDocument.Parse(s[a..(b + 1)]);
            JsonElement root = doc.RootElement;
            if (root.ValueKind != JsonValueKind.Object) throw new ArgumentException(fail);
            var rows = new List<MatchAnswerRow>();
            if (Prop(root, "rows") is { ValueKind: JsonValueKind.Array } arr)
                foreach (JsonElement r in arr.EnumerateArray())
                {
                    if (r.ValueKind != JsonValueKind.Object) continue;
                    int seq = int.TryParse(Str(Prop(r, "seq")), NumberStyles.Integer, CultureInfo.InvariantCulture, out int n) ? n : 0;
                    List<string> ids = Prop(r, "items") is { ValueKind: JsonValueKind.Array } ia
                        ? ia.EnumerateArray().Select(Str).Where(x => x.Length > 0).ToList() : [];
                    rows.Add(new MatchAnswerRow(seq, Str(Prop(r, "process")), ids, Str(Prop(r, "note"))));
                }
            return new MatchAnswer(rows, Str(Prop(root, "note")));
        }
        catch (JsonException) { throw new ArgumentException(fail); }
    }

    /// <summary>
    /// The Process List rows of an answer: every process of the order once and in order (one the answer leaves out is
    /// kept without items), a process the answer adds right after the one before it, numbered 01.. in that order; the
    /// cells are made from the items given (unknown ids left out): one per line, a jig's version and quantity in their own cells; no people or standard time.
    /// </summary>
    public static (List<WorkingProcessRow> Rows, List<string> Unmatched) Build(
        MatchAnswer answer, IReadOnlyList<MatchProcess> processes, IReadOnlyList<WorkingListItem> items)
    {
        var byId = items.ToDictionary(i => i.Id, StringComparer.OrdinalIgnoreCase);
        var bySeq = processes.ToDictionary(p => p.Seq);
        var main = new Dictionary<int, (string Process, List<string> Items, List<string> Notes)>();
        // added processes by the process they follow (0: before the first)
        var added = new Dictionary<int, List<(string Process, List<string> Items, List<string> Notes)>>();
        int after = 0;
        foreach (MatchAnswerRow r in answer.Rows)
        {
            List<string> ids = r.Items.Where(byId.ContainsKey).Select(id => byId[id].Id).ToList();
            if (bySeq.ContainsKey(r.Seq))
            {
                after = r.Seq;
                if (main.TryGetValue(r.Seq, out var had)) { had.Items.AddRange(ids); if (r.Note.Length > 0) had.Notes.Add(r.Note); }
                else main[r.Seq] = (r.Process, ids, r.Note.Length > 0 ? [r.Note] : []);
            }
            else if (r.Process.Trim().Length > 0)
            {
                if (!added.TryGetValue(after, out var list)) added[after] = list = [];
                list.Add((r.Process, ids, r.Note.Length > 0 ? [r.Note] : []));
            }
        }

        // an item the answer left out goes where the workbook puts it: the processes of its numbers (a process whose
        // code is that number) or the process of its name
        var placed = new HashSet<string>(main.Values.Concat(added.Values.SelectMany(x => x)).SelectMany(x => x.Items), StringComparer.Ordinal);
        foreach (WorkingListItem item in items.Where(i => !placed.Contains(i.Id)))
            foreach (MatchProcess p in processes)
            {
                bool byNo = int.TryParse(p.Code.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out int no) && item.ProcessNos.Contains(no);
                bool byName = item.Process.Trim().Length > 0 && string.Equals(item.Process.Trim(), p.Name.Trim(), StringComparison.OrdinalIgnoreCase);
                if (!byNo && !byName) continue;
                if (main.TryGetValue(p.Seq, out var had)) had.Items.Add(item.Id);
                else main[p.Seq] = ("", [item.Id], []);
            }

        var rows = new List<WorkingProcessRow>();
        var used = new HashSet<string>(StringComparer.Ordinal);
        void Add(string process, List<string> ids, List<string> notes)
        {
            List<WorkingListItem> its = ids.Distinct().Select(id => byId[id]).ToList();
            used.UnionWith(its.Select(i => i.Id));
            string note = string.Join(" / ", its.Select(i => i.Note).Concat(notes).Where(n => n.Trim().Length > 0).Distinct());
            rows.Add(WorkingProcessListService.Row("", (rows.Count + 1).ToString("00", CultureInfo.InvariantCulture), process.Trim(), its, note));
        }
        void AddAfter(int seq)
        {
            if (added.TryGetValue(seq, out var list))
                foreach (var x in list) Add(x.Process, x.Items, x.Notes);
        }
        AddAfter(0);
        foreach (MatchProcess p in processes.OrderBy(p => p.Seq))
        {
            if (main.TryGetValue(p.Seq, out var m))
                // the name of the order, so the rows tie to the Process table made from the same documents
                Add(p.Name.Trim().Length > 0 ? p.Name : m.Process, m.Items, m.Notes);
            else Add(p.Name, [], []);
            AddAfter(p.Seq);
        }
        if (rows.Count > WorkingProcessListService.MaxRows) rows = rows[..WorkingProcessListService.MaxRows];

        // what is left has no process here: why, so it does not read as an error
        List<string> unmatched = items.Where(i => !used.Contains(i.Id)).Select(i =>
            $"{(i.Kind == WorkingListItem.Jig ? "JIG" : i.Kind == WorkingListItem.Machine ? "Machine" : "Tester")}: {i.Label}{(i.Ver.Length > 0 ? $" ({i.Ver})" : "")}"
            + (i.Process.Length > 0 ? $" — {i.Process}" : "")
            + (i.ProcessNos.Count == 0 ? " (no process number; no process of that name in the list)"
                : $" (workbook process {string.Join(", ", i.ProcessNos)}; no such process in the list)")).ToList();
        return (rows, unmatched);
    }

    private static JsonElement? Prop(JsonElement e, string name)
    {
        foreach (JsonProperty p in e.EnumerateObject())
            if (string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase)) return p.Value;
        return null;
    }

    private static string Str(JsonElement? e) => e?.ValueKind switch
    {
        JsonValueKind.String => e.Value.GetString()?.Trim() ?? "",
        JsonValueKind.Number => e.Value.GetRawText(),
        _ => "",
    };

    private static string Str(JsonElement e) => Str((JsonElement?)e);

    private static string Limit(string? value, int max)
    {
        string v = (value ?? "").Trim();
        return v.Length > max ? v[..max] : v;
    }

    internal const string Instructions = """
        You put the machines, jigs and testers of a JIG / Machine / Tester list workbook on the processes of one
        manufacturing line. The input is the model name, the processes of the line IN ORDER (seq from 1, with a code, a
        name and the tools a work instruction names), the process numbers of the workbook (the numbers of the line layout
        with the names the workbook writes for them) and the items of the workbook (id, kind machine / jig / tester, name,
        the process it is written for and its process numbers).

        Match each item to the process or processes it is used in:
        - The process names may be in another language than the workbook (Vietnamese, Korean or English): match them by
          meaning ("Xếp UPPER" is UPPER ARRAY, "Bôi keo" is bonding / gluing / dispensing, "Kiểm tra AOI" is AOI TEST,
          "Tổ hợp" is assembly, "Chuyển JIG" / "지그 전환" is JIG SWAP, "Kiểm tra" is a test / inspection).
        - The process numbers of an item are numbers of the workbook's layout, NOT the seq of the order. Find the processes
          those numbers name in the workbook and put the item on the processes of the order that are the same work.
          The two numberings drift apart (the layout counts steps like DRY that have no work instruction).
        - Use the item's process text and the tools of a process as further hints. An item used over a range of
          processes (a carrier jig, "1 ~ 9") goes on each process of the range that is in the order.
        - A step of the workbook that is not in the order but has items of its own (for example DRY with a conveyor)
          is added as a row with seq 0, right after the process it follows on the line.
        - An item you cannot place with good confidence is left out of every row; never guess.

        Answer ONLY a JSON object, no other text:
        {"rows":[{"seq":1,"process":"<name>","items":["M1","J2"],"note":""}],"note":"..."}
        - One row per process of the order, in the order of seq, plus the added rows (seq 0) where they belong.
        - process: the name of the order; when the workbook's English name for it is different, add it in parentheses,
          e.g. "Xếp UPPER (UPPER ARRAY)". For an added row, the workbook's name.
        - items: the ids of the items on the process, machines first, then jigs, then testers.
        - note of a row: "" or a few words when the match is unsure.
        - note: 1 to 3 short English sentences on how the numberings were matched and what was left out.
        """;
}
