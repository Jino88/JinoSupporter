using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace PortalGNLAB.Web.Services;

/// <summary>A Setup Plan row: in the request the current rows, in the answer the rows to add (Id "") or to change (Id of an existing row).</summary>
public sealed class SetupAiRow
{
    public string Id { get; set; } = "";
    public int No { get; set; }
    public string LayoutNo { get; set; } = "";
    public string Line { get; set; } = "";
    public string ProcessName { get; set; } = "";
    public string MachinePlan { get; set; } = "";
    public string MachineState { get; set; } = "";
    public string MachineNote { get; set; } = "";
    public string SetupPlan { get; set; } = "";
    public string SetupFinish { get; set; } = "";
    public string SetupStatus { get; set; } = "";
}

public sealed class SetupAiRequest
{
    public string? Text { get; set; }
    public string? Today { get; set; }
    public List<SetupAiRow>? Rows { get; set; }
    public List<PlaceAiImage>? Images { get; set; }
}

public sealed class SetupAiAnswer
{
    public List<SetupAiRow> Rows { get; set; } = [];
    public string Note { get; set; } = "";
}

public sealed record SetupAiResult(List<SetupAiRow> Rows, string Note, string Model);

/// <summary>
/// Working > Mass Production > Plan > AI input: the user pastes notes, mails, Excel cells or pictures and the Codex API
/// turns them into Setup Plan rows. Nothing is saved here; the client previews the rows and saves the ones picked.
/// </summary>
public sealed class WorkingSetupAiService(CodexApiService codex)
{
    public const string Effort = "medium";
    private static readonly string[] States = ["O", "X", "Δ"];
    private static readonly string[] DateFormats =
        ["yyyy-MM-dd", "yyyy/MM/dd", "yyyy.MM.dd", "yyyy-M-d", "yyyy/M/d", "yyyy.M.d"];

    private static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public async Task<SetupAiResult> SetupAsync(string modelName, SetupAiRequest body, CancellationToken ct = default)
    {
        string request = body.Text?.Trim() ?? "";
        List<PlaceAiPicture> pictures = WorkingPlaceAiService.ReadImages(body.Images);
        if (request.Length == 0 && pictures.Count == 0) throw new ArgumentException("Type or paste what to put in.");
        if (!codex.IsConfigured) throw new ArgumentException("Codex API is not configured. Set the API key in Setting.");

        JsonArray content = BuildContent(modelName, body, pictures);

        string model = codex.Model;
        CodexTextResult result = await codex.GenerateAsync(Instructions, content, model, Effort, 16000, ct);
        codex.RecordUsage(result.Usage);

        SetupAiAnswer answer = ParseAnswer(result.Text, body);
        return new SetupAiResult(answer.Rows, answer.Note, $"{model} · {Effort}");
    }

    /// <summary>The AI input: one text block (model, current rows, picture list, user text) then each picture as input_image.</summary>
    public static JsonArray BuildContent(string modelName, SetupAiRequest body, IReadOnlyList<PlaceAiPicture> pictures)
    {
        var context = new { today = TodayOf(body), rows = body.Rows ?? [] };
        var text = new System.Text.StringBuilder();
        text.AppendLine($"Model name: {modelName}");
        text.AppendLine();
        text.AppendLine("=== CURRENT SETUP ROWS (JSON) ===");
        text.AppendLine(JsonSerializer.Serialize(context, Json));
        text.AppendLine();
        if (pictures.Count > 0)
        {
            text.AppendLine("=== PICTURES (numbered from 0, in the order they follow) ===");
            for (int i = 0; i < pictures.Count; i++) text.AppendLine($"Picture {i}: {pictures[i].Name}");
            text.AppendLine();
        }
        text.AppendLine("=== INPUT ===");
        text.AppendLine((body.Text ?? "").Trim());
        var content = new JsonArray { new JsonObject { ["type"] = "input_text", ["text"] = text.ToString() } };
        foreach (PlaceAiPicture p in pictures)
            content.Add(new JsonObject { ["type"] = "input_image", ["image_url"] = $"data:{p.MediaType};base64,{p.Base64}" });
        return content;
    }

    /// <summary>Reads the AI answer leniently: strips code fences, takes the outermost {...}, then normalizes.</summary>
    public static SetupAiAnswer ParseAnswer(string? text, SetupAiRequest? request)
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

            var answer = new SetupAiAnswer();
            if (root.TryGetProperty("rows", out JsonElement rows) && rows.ValueKind == JsonValueKind.Array)
                foreach (JsonElement r in rows.EnumerateArray())
                {
                    if (r.ValueKind != JsonValueKind.Object) throw new ArgumentException(fail);
                    answer.Rows.Add(ReadRow(r));
                }
            if (root.TryGetProperty("note", out JsonElement n) && n.ValueKind == JsonValueKind.String)
                answer.Note = n.GetString() ?? "";
            return Normalize(answer, request);
        }
        catch (JsonException) { throw new ArgumentException(fail); }
    }

    /// <summary>Cleans an answer against the request: known ids, valid dates and states, row length limits, no duplicate rows.</summary>
    public static SetupAiAnswer Normalize(SetupAiAnswer? answer, SetupAiRequest? request)
    {
        var known = new Dictionary<string, SetupAiRow>(StringComparer.Ordinal);
        foreach (SetupAiRow? r in request?.Rows ?? [])
            if (r is not null && !string.IsNullOrWhiteSpace(r.Id)) known[r.Id.Trim()] = r;

        var rows = new List<SetupAiRow>();
        var changed = new HashSet<string>(StringComparer.Ordinal);
        foreach (SetupAiRow? raw in answer?.Rows ?? [])
        {
            if (raw is null) continue;
            string id = (raw.Id ?? "").Trim();
            known.TryGetValue(id, out SetupAiRow? existing);
            // an id that is not a current row is a new row; a row already changed once is not changed twice
            if (existing is null) id = "";
            else if (!changed.Add(id)) continue;

            string process = Limit(raw.ProcessName, 200);
            if (process.Length == 0 && existing is not null) process = Limit(existing.ProcessName, 200);
            if (process.Length == 0) continue;
            var row = new SetupAiRow
            {
                Id = id,
                No = raw.No is > 0 and <= 100000 ? raw.No : 0,
                LayoutNo = Limit(raw.LayoutNo, 20),
                Line = Limit(raw.Line, 50),
                ProcessName = process,
                MachinePlan = DateOf(raw.MachinePlan),
                MachineState = StateOf(raw.MachineState),
                MachineNote = Limit(raw.MachineNote, 100),
                SetupPlan = DateOf(raw.SetupPlan),
                SetupFinish = DateOf(raw.SetupFinish),
                SetupStatus = StateOf(raw.SetupStatus),
            };
            // a new row that repeats a current row (same line and process) is not made again
            if (existing is null && known.Values.Any(k => Same(k, row))) continue;
            if (existing is null && rows.Any(k => k.Id.Length == 0 && Same(k, row))) continue;
            rows.Add(row);
        }
        return new SetupAiAnswer { Rows = rows, Note = (answer?.Note ?? "").Trim() };
    }

    private static bool Same(SetupAiRow a, SetupAiRow b) =>
        string.Equals(a.Line.Trim(), b.Line.Trim(), StringComparison.OrdinalIgnoreCase)
        && string.Equals(a.ProcessName.Trim(), b.ProcessName.Trim(), StringComparison.OrdinalIgnoreCase);

    private static SetupAiRow ReadRow(JsonElement e) => new()
    {
        Id = Text(e, "id"),
        No = int.TryParse(Text(e, "no"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int no) ? no : 0,
        LayoutNo = Text(e, "layoutNo"),
        Line = Text(e, "line"),
        ProcessName = Text(e, "processName"),
        MachinePlan = Text(e, "machinePlan"),
        MachineState = Text(e, "machineState"),
        MachineNote = Text(e, "machineNote"),
        SetupPlan = Text(e, "setupPlan"),
        SetupFinish = Text(e, "setupFinish"),
        SetupStatus = Text(e, "setupStatus"),
    };

    /// <summary>A property as text: strings as they are, numbers and booleans as written, anything else "".</summary>
    private static string Text(JsonElement e, string name)
    {
        foreach (JsonProperty p in e.EnumerateObject())
            if (string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase))
                return p.Value.ValueKind switch
                {
                    JsonValueKind.String => p.Value.GetString() ?? "",
                    JsonValueKind.Number or JsonValueKind.True or JsonValueKind.False => p.Value.GetRawText(),
                    _ => "",
                };
        return "";
    }

    private static string Limit(string? value, int max)
    {
        string v = (value ?? "").Trim();
        return v.Length > max ? v[..max] : v;
    }

    private static string StateOf(string? value)
    {
        string v = (value ?? "").Trim().ToUpperInvariant();
        v = v switch { "0" or "OK" or "○" or "◯" or "V" or "✓" => "O", "×" or "✕" or "NG" => "X", "△" or "▵" or "D" => "Δ", _ => v };
        return States.Contains(v) ? v : "";
    }

    private static string TodayOf(SetupAiRequest? request)
    {
        string d = DateOf(request?.Today);
        return d.Length > 0 ? d : DateTime.Today.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
    }

    /// <summary>yyyy-MM-dd (also yyyy/MM/dd, yyyy.MM.dd) as yyyy-MM-dd; anything else "".</summary>
    private static string DateOf(string? value) =>
        DateTime.TryParseExact((value ?? "").Trim(), DateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime d)
            ? d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : "";

    internal const string Instructions = """
        You turn free text into Setup Plan rows of one manufacturing model (Mass Production > Plan). The input is the model
        name, the current rows as JSON ({"today":"yyyy-MM-dd","rows":[{"id","no","layoutNo","line","processName",
        "machinePlan","machineState","machineNote","setupPlan","setupFinish","setupStatus"}]}) and the user's text. The text
        may be meeting notes, a mail, cells copied from Excel (tab separated) or pictures of a schedule. Read it and answer
        the rows that should be added or changed.

        A row is one process on one line:
        - no: row number (integer, 0 when unknown; a new row with 0 goes after the last row)
        - layoutNo: the layout number of the process on the floor plan (text)
        - line: production line name
        - processName: the process (required; a row without it is dropped)
        - machinePlan: date (yyyy-MM-dd) the machine is planned to arrive or be ready, or ""
        - machineState: "" (not known), "O" (machine ready), "X" (not yet) or "Δ" (in progress)
        - machineNote: short note about the machine (up to 100 characters)
        - setupPlan: date (yyyy-MM-dd) the setup is planned, or ""
        - setupFinish: date (yyyy-MM-dd) the setup finished or is expected to finish, or ""
        - setupStatus: "" (not known), "O" (setup done), "X" (not yet set up) or "Δ" (setting up now)

        Rules:
        - Answer ONLY a JSON object: {"rows":[{"id":"","no":0,"layoutNo":"","line":"","processName":"","machinePlan":"",
          "machineState":"","machineNote":"","setupPlan":"","setupFinish":"","setupStatus":""}],"note":"..."} with no other text.
        - id "" makes a new row. To change a current row give its id and ALL its fields (keep the values that do not change).
        - If a row in the input is already a current row (same line and process), change it with its id; never add it again.
        - Dates are yyyy-MM-dd. A date without a year takes the year of today. Use "" when a date is not given; never guess.
        - States are only "", "O", "X" or "Δ". Words like done/OK mean "O", not yet/pending mean "X", in progress/now mean "Δ".
        - Do not invent rows or values that are not in the input. Leave out what does not fit a setup row.
        - note: 1 to 3 short English sentences saying what you made and anything left out or unclear.
        """;
}
