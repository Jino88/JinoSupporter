using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace PortalGNLAB.Web.Services;

public sealed class PlaceAiTask
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string Start { get; set; } = "";
    public string End { get; set; } = "";
    public string Owner { get; set; } = "";
    public bool Done { get; set; }
}

public sealed class PlaceAiEvent
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string Start { get; set; } = "";
    public string End { get; set; } = "";
    public List<PlaceAiTask> Tasks { get; set; } = [];
}

public sealed class PlaceAiEntryContext
{
    public string Id { get; set; } = "";
    public string Kind { get; set; } = "";
    public string Stage { get; set; } = "";
    public int Seq { get; set; }
    public string Date { get; set; } = "";
    public string Title { get; set; } = "";
    public string Status { get; set; } = "";
}

public sealed class PlaceAiInfo
{
    public string Customer { get; set; } = "";
    public string Build { get; set; } = "";
    public string Phase { get; set; } = "";
    public string Owner { get; set; } = "";
    public string Line { get; set; } = "";
    public string SopTarget { get; set; } = "";
    public string Note { get; set; } = "";
}

public sealed class PlaceAiEntry
{
    public string Kind { get; set; } = "";
    public string Stage { get; set; } = "";
    public string Date { get; set; } = "";
    public string Title { get; set; } = "";
    public string Body { get; set; } = "";
    public string Owner { get; set; } = "";
    public string Due { get; set; } = "";
    public string Severity { get; set; } = "";
    public string Result { get; set; } = "";
    public string Phase { get; set; } = "";
    public string Category { get; set; } = "";
    public string Process { get; set; } = "";
    public string Cause { get; set; } = "";
}

public sealed class PlaceAiRequest
{
    public string? Text { get; set; }
    public string? Today { get; set; }
    public PlaceAiInfo? Info { get; set; }
    public List<PlaceAiEvent>? Events { get; set; }
    public List<PlaceAiEntryContext>? Entries { get; set; }
}

/// <summary>The placement plan: only the Model Info fields to change, the Dev Plan events/tasks and the entries to add.</summary>
public sealed class PlaceAiPlan
{
    public Dictionary<string, string> Info { get; set; } = [];
    public List<PlaceAiEvent> Events { get; set; } = [];
    public List<PlaceAiEntry> Entries { get; set; } = [];
    public string Note { get; set; } = "";
}

public sealed record PlaceAiResult(
    Dictionary<string, string> Info, List<PlaceAiEvent> Events, List<PlaceAiEntry> Entries, string Note, string Model);

/// <summary>
/// Working > AI tab: the user types or pastes free text and the Codex API decides where each piece belongs
/// (Model Info, Dev Plan events/tasks, Dev History / Dev Issue / Review / MP Issue entries). Nothing is saved
/// here; the client previews the plan and saves it through the existing APIs.
/// </summary>
public sealed class WorkingPlaceAiService(CodexApiService codex)
{
    public const string Effort = "medium";
    private static readonly string[] InfoKeys = ["customer", "build", "phase", "owner", "line", "sopTarget", "note"];
    private static readonly string[] Phases = ["ER", "DR", "DV", "PV", "MP"];
    private static readonly string[] Kinds = ["dev", "issue", "review"];
    private static readonly string[] Severities = ["Critical", "Major", "Minor"];
    private static readonly string[] Results = ["Pass", "Conditional", "Fail"];
    private static readonly string[] DateFormats =
        ["yyyy-MM-dd", "yyyy/MM/dd", "yyyy.MM.dd", "yyyy-M-d", "yyyy/M/d", "yyyy.M.d"];

    private static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public async Task<PlaceAiResult> PlaceAsync(string modelName, PlaceAiRequest body, CancellationToken ct = default)
    {
        string request = body.Text?.Trim() ?? "";
        if (request.Length == 0) throw new ArgumentException("Type or paste what to put in.");
        if (!codex.IsConfigured) throw new ArgumentException("Codex API is not configured. Set the API key in Setting.");

        string today = TodayOf(body);
        var context = new
        {
            today,
            info = body.Info ?? new PlaceAiInfo(),
            events = body.Events ?? [],
            entries = body.Entries ?? [],
        };
        var text = new System.Text.StringBuilder();
        text.AppendLine($"Model name: {modelName}");
        text.AppendLine();
        text.AppendLine("=== CURRENT STATE (JSON) ===");
        text.AppendLine(JsonSerializer.Serialize(context, Json));
        text.AppendLine();
        text.AppendLine("=== INPUT ===");
        text.AppendLine(request);
        var content = new JsonArray { new JsonObject { ["type"] = "input_text", ["text"] = text.ToString() } };

        string model = codex.Model;
        CodexTextResult result = await codex.GenerateAsync(Instructions, content, model, Effort, 16000, ct);
        codex.RecordUsage(result.Usage);

        PlaceAiPlan plan = ParseAnswer(result.Text, body);
        return new PlaceAiResult(plan.Info, plan.Events, plan.Entries, plan.Note, $"{model} · {Effort}");
    }

    /// <summary>Reads the AI answer leniently: strips code fences, takes the outermost {...}, then normalizes.</summary>
    public static PlaceAiPlan ParseAnswer(string? text, PlaceAiRequest? request)
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

            var plan = new PlaceAiPlan();
            if (root.TryGetProperty("info", out JsonElement info) && info.ValueKind == JsonValueKind.Object)
                foreach (JsonProperty p in info.EnumerateObject())
                    plan.Info[p.Name] = p.Value.ValueKind switch
                    {
                        JsonValueKind.String => p.Value.GetString() ?? "",
                        JsonValueKind.Number or JsonValueKind.True or JsonValueKind.False => p.Value.GetRawText(),
                        _ => "",
                    };
            if (root.TryGetProperty("events", out JsonElement ev) && ev.ValueKind == JsonValueKind.Array)
                plan.Events = ev.Deserialize<List<PlaceAiEvent>>(Json) ?? [];
            if (root.TryGetProperty("entries", out JsonElement en) && en.ValueKind == JsonValueKind.Array)
                plan.Entries = en.Deserialize<List<PlaceAiEntry>>(Json) ?? [];
            if (root.TryGetProperty("note", out JsonElement n) && n.ValueKind == JsonValueKind.String)
                plan.Note = n.GetString() ?? "";
            return Normalize(plan, request);
        }
        catch (JsonException) { throw new ArgumentException(fail); }
    }

    /// <summary>Cleans a plan against the request: allowed info keys and values, known ids, valid dates, entry fields.</summary>
    public static PlaceAiPlan Normalize(PlaceAiPlan? plan, PlaceAiRequest? request)
    {
        string today = TodayOf(request);
        PlaceAiInfo current = request?.Info ?? new PlaceAiInfo();
        var known = new Dictionary<string, PlaceAiEvent>(StringComparer.Ordinal);
        foreach (PlaceAiEvent? e in request?.Events ?? [])
            if (e is not null && !string.IsNullOrWhiteSpace(e.Id)) known[e.Id.Trim()] = e;

        return new PlaceAiPlan
        {
            Info = NormalizeInfo(plan?.Info, current),
            Events = NormalizeEvents(plan?.Events, known),
            Entries = NormalizeEntries(plan?.Entries, today),
            Note = (plan?.Note ?? "").Trim(),
        };
    }

    private static Dictionary<string, string> NormalizeInfo(Dictionary<string, string>? info, PlaceAiInfo current)
    {
        var result = new Dictionary<string, string>();
        foreach ((string rawKey, string? rawValue) in info ?? [])
        {
            string? key = InfoKeys.FirstOrDefault(k => string.Equals(k, (rawKey ?? "").Trim(), StringComparison.OrdinalIgnoreCase));
            if (key is null) continue;
            string value = (rawValue ?? "").Trim();
            if (key == "phase") value = PhaseOf(value, allowMp: true);
            else if (key == "sopTarget") value = DateOf(value);
            if (value.Length == 0) continue;
            string now = key switch
            {
                "customer" => current.Customer,
                "build" => current.Build,
                "phase" => current.Phase,
                "owner" => current.Owner,
                "line" => current.Line,
                "sopTarget" => current.SopTarget,
                _ => current.Note,
            };
            if (string.Equals(value, (now ?? "").Trim(), StringComparison.Ordinal)) continue;
            result[key] = value;
        }
        return result;
    }

    private static List<PlaceAiEvent> NormalizeEvents(List<PlaceAiEvent>? events, Dictionary<string, PlaceAiEvent> known)
    {
        var result = new List<PlaceAiEvent>();
        foreach (PlaceAiEvent? e in events ?? [])
        {
            if (e is null) continue;
            string id = (e.Id ?? "").Trim();
            known.TryGetValue(id, out PlaceAiEvent? existing);
            if (existing is null) id = "";

            string name = (e.Name ?? "").Trim(), start = DateOf(e.Start), end = DateOf(e.End);
            if (existing is not null)
            {
                // blank means "leave as it is"
                if (name.Length == 0) name = (existing.Name ?? "").Trim();
                if (start.Length == 0) start = DateOf(existing.Start);
                if (end.Length == 0) end = DateOf(existing.End);
            }
            else if (name.Length == 0) continue;

            var taskIds = new HashSet<string>(
                (existing?.Tasks ?? []).Where(t => t is not null).Select(t => (t.Id ?? "").Trim()).Where(i => i.Length > 0),
                StringComparer.Ordinal);
            var tasks = new List<PlaceAiTask>();
            foreach (PlaceAiTask? t in e.Tasks ?? [])
            {
                string tName = (t?.Name ?? "").Trim();
                if (t is null || tName.Length == 0) continue;
                string tId = (t.Id ?? "").Trim();
                tasks.Add(new PlaceAiTask
                {
                    Id = taskIds.Contains(tId) ? tId : "",
                    Name = tName,
                    Start = DateOf(t.Start),
                    End = DateOf(t.End),
                    Owner = (t.Owner ?? "").Trim(),
                    Done = t.Done,
                });
            }

            if (existing is not null && tasks.Count == 0
                && name == (existing.Name ?? "").Trim()
                && start == DateOf(existing.Start)
                && end == DateOf(existing.End)) continue;
            result.Add(new PlaceAiEvent { Id = id, Name = name, Start = start, End = end, Tasks = tasks });
        }
        return result;
    }

    private static List<PlaceAiEntry> NormalizeEntries(List<PlaceAiEntry>? entries, string today)
    {
        var result = new List<PlaceAiEntry>();
        foreach (PlaceAiEntry? e in entries ?? [])
        {
            if (e is null) continue;
            string kind = (e.Kind ?? "").Trim().ToLowerInvariant();
            if (!Kinds.Contains(kind)) continue;
            string title = (e.Title ?? "").Trim();
            if (title.Length == 0) continue;
            if (title.Length > 200) title = title[..200].TrimEnd();
            string date = DateOf(e.Date);

            var entry = new PlaceAiEntry
            {
                Kind = kind,
                Date = date.Length > 0 ? date : today,
                Title = title,
                Body = (e.Body ?? "").Trim(),
                Owner = (e.Owner ?? "").Trim(),
                Result = "",
            };
            if (kind == "issue")
            {
                string phase = PhaseOf(e.Phase, allowMp: true);
                bool mp = string.Equals((e.Stage ?? "").Trim(), "mp", StringComparison.OrdinalIgnoreCase) || phase == "MP";
                entry.Stage = mp ? "mp" : "dev";
                entry.Phase = mp ? "MP" : phase;
                entry.Due = DateOf(e.Due);
                entry.Severity = Severities.FirstOrDefault(v => string.Equals(v, (e.Severity ?? "").Trim(), StringComparison.OrdinalIgnoreCase)) ?? "";
                entry.Category = (e.Category ?? "").Trim();
                entry.Process = (e.Process ?? "").Trim();
                entry.Cause = (e.Cause ?? "").Trim();
            }
            else if (kind == "review")
            {
                entry.Result = Results.FirstOrDefault(v => string.Equals(v, (e.Result ?? "").Trim(), StringComparison.OrdinalIgnoreCase)) ?? "Pass";
            }
            result.Add(entry);
        }
        return result;
    }

    private static string TodayOf(PlaceAiRequest? request)
    {
        string d = DateOf(request?.Today);
        return d.Length > 0 ? d : DateTime.Today.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
    }

    /// <summary>yyyy-MM-dd (also yyyy/MM/dd, yyyy.MM.dd) as yyyy-MM-dd; anything else "".</summary>
    private static string DateOf(string? value) =>
        DateTime.TryParseExact((value ?? "").Trim(), DateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime d)
            ? d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : "";

    private static string PhaseOf(string? value, bool allowMp)
    {
        string v = (value ?? "").Trim().ToUpperInvariant();
        v = v switch { "PROTO" => "ER", "EVT" => "DR", "DVT" => "DV", "PVT" => "PV", _ => v };
        return Phases.Contains(v) && (allowMp || v != "MP") ? v : "";
    }

    private const string Instructions = """
        You place free text into the Working page of one manufacturing model. The input is the model name, the
        current state as JSON ({"today":"yyyy-MM-dd","info":{"customer","build","phase","owner","line","sopTarget",
        "note"},"events":[{"id","name","start","end","tasks":[{"id","name","start","end","owner","done"}]}],
        "entries":[{"id","kind","stage","seq","date","title","status"}]}) and the user's text. The text may contain
        cells copied from Excel (tab separated). Read it and put each piece of information in the most fitting place:
        - Model Info: customer, build, phase (ER, DR, DV, PV or MP), owner (the PM in charge), line, sopTarget
          (yyyy-MM-dd) and a free note.
        - Dev Plan: development events that have a schedule (builds, tooling, reliability tests ...) with their
          detailed tasks (start, end, owner, done). If an event is the same as an existing one, use its id and give
          only the tasks to add or change; if a task is the same as an existing one, use its id to change it.
          A new event has id "".
        - entries, kind "dev": a development history record or a decision.
        - entries, kind "issue": a problem. In development (stage "dev", phase ER/DR/DV/PV or "") a defect or design
          problem; in mass production (stage "mp", phase "MP") a mass-production problem.
          severity is Critical, Major or Minor; category, process and cause are short texts; due is a date.
        - entries, kind "review": the result of a review or meeting; result is Pass, Conditional or Fail.

        Answer ONLY one JSON object, no other text:
        {"info":{only the fields to change},"events":[{"id","name","start","end","tasks":[{"id","name","start","end","owner","done"}]}],
        "entries":[{"kind","stage","date","title","body","owner","due","severity","result","phase","category","process","cause"}],
        "note":"..."}
        - note: one to three English sentences saying what was put where, and why anything could not be placed.
        - Do not create entries that already exist in the current entries, and do not repeat unchanged info.
        - Dates are yyyy-MM-dd; a date without a year uses the year of today.
        - Keep the original language of the text; do not translate.
        - Create only what the text supports; do not invent anything.
        """;
}
