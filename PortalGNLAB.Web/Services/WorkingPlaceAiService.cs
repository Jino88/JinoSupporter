using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
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
    /// <summary>Quantity table rows: in the request the current rows, in the answer the rows to add or change.</summary>
    public List<PlaceAiRow> Rows { get; set; } = [];
    /// <summary>Answer only: ids of existing tasks that moved to the quantity table.</summary>
    [JsonConverter(typeof(PlaceAiStringListConverter))]
    public List<string> RemoveTasks { get; set; } = [];
}

/// <summary>A Dev Plan quantity row: group "material" (Material Plan) or "making" (Making Plan), quantity per day.</summary>
public sealed class PlaceAiRow
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string Group { get; set; } = "";
    [JsonConverter(typeof(PlaceAiQtyConverter))]
    public Dictionary<string, double> Qty { get; set; } = [];
}

/// <summary>Reads a JSON array of strings; numbers become text and everything else is ignored.</summary>
public sealed class PlaceAiStringListConverter : JsonConverter<List<string>>
{
    public override List<string> Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        var list = new List<string>();
        if (reader.TokenType != JsonTokenType.StartArray) { reader.Skip(); return list; }
        while (reader.Read() && reader.TokenType != JsonTokenType.EndArray)
        {
            if (reader.TokenType == JsonTokenType.String) list.Add(reader.GetString() ?? "");
            else if (reader.TokenType == JsonTokenType.Number) list.Add(Encoding.UTF8.GetString(reader.ValueSpan));
            else reader.Skip();
        }
        return list;
    }

    public override void Write(Utf8JsonWriter writer, List<string> value, JsonSerializerOptions options)
    {
        writer.WriteStartArray();
        foreach (string v in value) writer.WriteStringValue(v);
        writer.WriteEndArray();
    }
}

/// <summary>Reads a JSON object of date → quantity; a quantity may be a number or text like "1,200", anything else is ignored.</summary>
public sealed class PlaceAiQtyConverter : JsonConverter<Dictionary<string, double>>
{
    public override Dictionary<string, double> Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        var map = new Dictionary<string, double>(StringComparer.Ordinal);
        if (reader.TokenType != JsonTokenType.StartObject) { reader.Skip(); return map; }
        while (reader.Read() && reader.TokenType != JsonTokenType.EndObject)
        {
            string key = reader.GetString() ?? "";
            reader.Read();
            double n = double.NaN;
            if (reader.TokenType == JsonTokenType.Number) n = reader.GetDouble();
            else if (reader.TokenType == JsonTokenType.String)
                n = double.TryParse((reader.GetString() ?? "").Replace(",", "").Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                    ? parsed : double.NaN;
            else reader.Skip();
            if (double.IsFinite(n)) map[key] = n;
        }
        return map;
    }

    public override void Write(Utf8JsonWriter writer, Dictionary<string, double> value, JsonSerializerOptions options)
    {
        writer.WriteStartObject();
        foreach ((string k, double v) in value) writer.WriteNumber(k, v);
        writer.WriteEndObject();
    }
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
    /// <summary>0-based numbers of the request pictures that are evidence for this entry.</summary>
    [JsonConverter(typeof(PlaceAiIntListConverter))]
    public List<int> Images { get; set; } = [];
}

/// <summary>A pasted picture: file name and a data URL (data:image/png;base64,...).</summary>
public sealed class PlaceAiImage
{
    public string Name { get; set; } = "";
    public string DataUrl { get; set; } = "";
}

/// <summary>Reads a JSON array of numbers and ignores everything in it that is not an integer.</summary>
public sealed class PlaceAiIntListConverter : JsonConverter<List<int>>
{
    public override List<int> Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        var list = new List<int>();
        if (reader.TokenType != JsonTokenType.StartArray) { reader.Skip(); return list; }
        while (reader.Read() && reader.TokenType != JsonTokenType.EndArray)
        {
            if (reader.TokenType == JsonTokenType.Number && reader.TryGetInt32(out int n)) list.Add(n);
            else reader.Skip();
        }
        return list;
    }

    public override void Write(Utf8JsonWriter writer, List<int> value, JsonSerializerOptions options)
    {
        writer.WriteStartArray();
        foreach (int n in value) writer.WriteNumberValue(n);
        writer.WriteEndArray();
    }
}

public sealed class PlaceAiRequest
{
    public string? Text { get; set; }
    public string? Today { get; set; }
    public PlaceAiInfo? Info { get; set; }
    public List<PlaceAiEvent>? Events { get; set; }
    public List<PlaceAiEntryContext>? Entries { get; set; }
    public List<PlaceAiImage>? Images { get; set; }
    /// <summary>
    /// The tab the AI input was opened on: plan (Dev Plan), devissue, dev (Dev History), review, issue (MP Issues) or
    /// entries (Timeline); only that kind of data is answered. Empty or unknown: anywhere (AI tab).
    /// </summary>
    public string? Target { get; set; }
}

/// <summary>A checked picture: its name, media type and base64 data.</summary>
public sealed record PlaceAiPicture(string Name, string MediaType, string Base64);

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
    public const int MaxImages = 6;
    public const int MaxImageBytes = 5 * 1024 * 1024;
    private static readonly string[] ImageTypes = ["image/png", "image/jpeg", "image/gif", "image/webp"];
    private static readonly Regex DataUrl = new(@"^data:([^;,]+);base64,(.*)$", RegexOptions.Singleline | RegexOptions.IgnoreCase);
    private static readonly string[] InfoKeys = ["customer", "build", "phase", "owner", "line", "sopTarget", "note"];
    private static readonly string[] Phases = ["ER", "DR", "DV", "PV", "MP"];
    private static readonly string[] Kinds = ["dev", "issue", "review"];
    private static readonly string[] Severities = ["Critical", "Major", "Minor"];
    private static readonly string[] Results = ["Pass", "Conditional", "Fail"];
    private static readonly string[] DateFormats =
        ["yyyy-MM-dd", "yyyy/MM/dd", "yyyy.MM.dd", "yyyy-M-d", "yyyy/M/d", "yyyy.M.d"];

    /// <summary>What the AI is told for each target (PlaceAiRequest.Target).</summary>
    private static readonly Dictionary<string, string> Targets = new(StringComparer.Ordinal)
    {
        ["plan"] = "Dev Plan. Answer only events (with their tasks, rows and removeTasks); info is {} and entries is [].",
        ["devissue"] = "Dev Issue. Answer only entries of kind \"issue\" with stage \"dev\" (phase ER, DR, DV, PV or \"\"); info is {} and events is [].",
        ["dev"] = "Dev History. Answer only entries of kind \"dev\"; info is {} and events is [].",
        ["review"] = "Reviews. Answer only entries of kind \"review\"; info is {} and events is [].",
        ["issue"] = "MP Issues. Answer only entries of kind \"issue\" with stage \"mp\" and phase \"MP\"; info is {} and events is [].",
        ["entries"] = "Timeline. Answer only entries (kind dev, issue or review); info is {} and events is [].",
    };

    private static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public async Task<PlaceAiResult> PlaceAsync(string modelName, PlaceAiRequest body, CancellationToken ct = default)
    {
        string request = body.Text?.Trim() ?? "";
        List<PlaceAiPicture> pictures = ReadImages(body.Images);
        if (request.Length == 0 && pictures.Count == 0) throw new ArgumentException("Type or paste what to put in.");
        if (!codex.IsConfigured) throw new ArgumentException("Codex API is not configured. Set the API key in Setting.");

        JsonArray content = BuildContent(modelName, body, pictures);

        string model = codex.Model;
        CodexTextResult result = await codex.GenerateAsync(Instructions, content, model, Effort, 16000, ct);
        codex.RecordUsage(result.Usage);

        PlaceAiPlan plan = ParseAnswer(result.Text, body);
        return new PlaceAiResult(plan.Info, plan.Events, plan.Entries, plan.Note, $"{model} · {Effort}");
    }

    /// <summary>Checks the request pictures (count, type, base64, size) and returns them ready to send.</summary>
    public static List<PlaceAiPicture> ReadImages(List<PlaceAiImage>? images)
    {
        var result = new List<PlaceAiPicture>();
        if (images is null || images.Count == 0) return result;
        if (images.Count > MaxImages) throw new ArgumentException("Up to 6 pictures can be sent.");
        foreach (PlaceAiImage? img in images)
        {
            Match m = DataUrl.Match((img?.DataUrl ?? "").Trim());
            string type = m.Success ? m.Groups[1].Value.Trim().ToLowerInvariant() : "";
            if (!ImageTypes.Contains(type)) throw new ArgumentException("Only PNG, JPEG, GIF or WebP pictures can be read.");
            string b64 = m.Groups[2].Value.Trim();
            // the base64 length gives the decoded size, so a huge string is rejected before it is decoded
            if (b64.Length / 4 * 3 > MaxImageBytes + 3) throw new ArgumentException("A picture is larger than 5 MB.");
            byte[] bytes = new byte[b64.Length / 4 * 3 + 3];
            if (!Convert.TryFromBase64String(b64, bytes, out int written) || written == 0)
                throw new ArgumentException("A picture could not be read. Paste it again.");
            if (written > MaxImageBytes) throw new ArgumentException("A picture is larger than 5 MB.");
            string name = (img!.Name ?? "").Trim();
            result.Add(new PlaceAiPicture(name.Length > 0 ? name : $"picture-{result.Count + 1}", type, b64));
        }
        return result;
    }

    /// <summary>The AI input: one text block (model, current state, picture list, user text) then each picture as input_image.</summary>
    public static JsonArray BuildContent(string modelName, PlaceAiRequest body, IReadOnlyList<PlaceAiPicture> pictures)
    {
        var context = new
        {
            today = TodayOf(body),
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
        if (pictures.Count > 0)
        {
            text.AppendLine("=== PICTURES (numbered from 0, in the order they follow) ===");
            for (int i = 0; i < pictures.Count; i++) text.AppendLine($"Picture {i}: {pictures[i].Name}");
            text.AppendLine();
        }
        if (Targets.TryGetValue(TargetOf(body), out string? target))
        {
            text.AppendLine("=== TARGET ===");
            text.AppendLine($"The user opened AI input on {target} Make everything in the input into this form; leave out what does not fit and say so in the note.");
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

        // a target keeps only its own kind of data; entries are made into the kind of the tab
        string target = TargetOf(request);
        List<PlaceAiEntry>? entries = plan?.Entries;
        if (target is "devissue" or "dev" or "review" or "issue") entries = entries?.Select(e => Retarget(e, target)).ToList();
        return new PlaceAiPlan
        {
            Info = target.Length == 0 ? NormalizeInfo(plan?.Info, current) : [],
            Events = target is "" or "plan" ? NormalizeEvents(plan?.Events, known, today) : [],
            Entries = target == "plan" ? [] : NormalizeEntries(entries, today, request?.Images?.Count ?? 0),
            Note = (plan?.Note ?? "").Trim(),
        };
    }

    /// <summary>The target of a request when it is a known one, else "".</summary>
    private static string TargetOf(PlaceAiRequest? request)
    {
        string t = (request?.Target ?? "").Trim().ToLowerInvariant();
        return Targets.ContainsKey(t) ? t : "";
    }

    /// <summary>An entry made into the kind of a Dev Issue, Dev History, Reviews or MP Issues tab.</summary>
    private static PlaceAiEntry? Retarget(PlaceAiEntry? e, string target)
    {
        if (e is null) return null;
        e.Kind = target switch { "dev" => "dev", "review" => "review", _ => "issue" };
        if (target == "issue") { e.Stage = "mp"; e.Phase = "MP"; }
        else if (target == "devissue") { e.Stage = "dev"; e.Phase = PhaseOf(e.Phase, allowMp: false); }
        return e;
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

    private static List<PlaceAiEvent> NormalizeEvents(List<PlaceAiEvent>? events, Dictionary<string, PlaceAiEvent> known, string today)
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

            var rows = NormalizeRows(e.Rows, existing, today);
            var removeTasks = new List<string>();
            if (existing is not null)
                foreach (string? r in e.RemoveTasks ?? [])
                {
                    string rId = (r ?? "").Trim();
                    if (taskIds.Contains(rId) && !removeTasks.Contains(rId)) removeTasks.Add(rId);
                }

            if (existing is not null && tasks.Count == 0 && rows.Count == 0 && removeTasks.Count == 0
                && name == (existing.Name ?? "").Trim()
                && start == DateOf(existing.Start)
                && end == DateOf(existing.End)) continue;
            result.Add(new PlaceAiEvent { Id = id, Name = name, Start = start, End = end, Tasks = tasks, Rows = rows, RemoveTasks = removeTasks });
        }
        return result;
    }

    private static List<PlaceAiRow> NormalizeRows(List<PlaceAiRow>? rows, PlaceAiEvent? existing, string today)
    {
        var known = new Dictionary<string, PlaceAiRow>(StringComparer.Ordinal);
        foreach (PlaceAiRow? r in existing?.Rows ?? [])
            if (r is not null && !string.IsNullOrWhiteSpace(r.Id)) known[r.Id.Trim()] = r;

        var result = new List<PlaceAiRow>();
        foreach (PlaceAiRow? r in rows ?? [])
        {
            if (r is null) continue;
            var qty = new Dictionary<string, double>(StringComparer.Ordinal);
            foreach ((string rawDate, double v) in r.Qty ?? [])
            {
                string date = DateOf(rawDate, today);
                if (date.Length > 0 && double.IsFinite(v) && v >= 0) qty[date] = v;
            }
            known.TryGetValue((r.Id ?? "").Trim(), out PlaceAiRow? old);
            if (old is not null)
            {
                var before = DatesOf(old.Qty, today);
                if (qty.All(q => before.GetValueOrDefault(q.Key) == q.Value)) continue;
                result.Add(new PlaceAiRow { Id = old.Id.Trim(), Name = (old.Name ?? "").Trim(), Group = old.Group ?? "", Qty = qty });
                continue;
            }
            string name = (r.Name ?? "").Trim();
            if (name.Length == 0) continue;
            string group = string.Equals((r.Group ?? "").Trim(), "material", StringComparison.OrdinalIgnoreCase) ? "material" : "making";
            result.Add(new PlaceAiRow { Id = "", Name = name, Group = group, Qty = qty });
        }
        return result;
    }

    private static Dictionary<string, double> DatesOf(Dictionary<string, double>? qty, string today)
    {
        var map = new Dictionary<string, double>(StringComparer.Ordinal);
        foreach ((string k, double v) in qty ?? []) { string d = DateOf(k, today); if (d.Length > 0) map[d] = v; }
        return map;
    }

    private static List<PlaceAiEntry> NormalizeEntries(List<PlaceAiEntry>? entries, string today, int imageCount)
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
                Images = (e.Images ?? []).Where(n => n >= 0 && n < imageCount).Distinct().ToList(),
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

    /// <summary>Like DateOf, and also M-d / M/d / M.d (no year), which take the year of today.</summary>
    private static string DateOf(string? value, string today)
    {
        string full = DateOf(value);
        if (full.Length > 0) return full;
        string v = (value ?? "").Trim();
        if (v.Length == 0 || today.Length < 4) return "";
        return v.Length <= 5 && DateTime.TryParseExact($"{today[..4]}-{v.Replace('/', '-').Replace('.', '-')}", "yyyy-M-d",
            CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime d)
            ? d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : "";
    }

    private static string PhaseOf(string? value, bool allowMp)
    {
        string v = (value ?? "").Trim().ToUpperInvariant();
        v = v switch { "PROTO" => "ER", "EVT" => "DR", "DVT" => "DV", "PVT" => "PV", _ => v };
        return Phases.Contains(v) && (allowMp || v != "MP") ? v : "";
    }

    internal const string Instructions = """
        You place free text into the Working page of one manufacturing model. The input is the model name, the
        current state as JSON ({"today":"yyyy-MM-dd","info":{"customer","build","phase","owner","line","sopTarget",
        "note"},"events":[{"id","name","start","end","tasks":[{"id","name","start","end","owner","done"}],
        "rows":[{"id","name","group","qty":{"yyyy-MM-dd":number}}]}],
        "entries":[{"id","kind","stage","seq","date","title","status"}]}) and the user's text. The text may contain
        cells copied from Excel (tab separated). Read it and put each piece of information in the most fitting place:
        - Model Info: customer, build, phase (ER, DR, DV, PV or MP), owner (the PM in charge), line, sopTarget
          (yyyy-MM-dd) and a free note.
        - Dev Plan: development events that have a schedule (builds, tooling, reliability tests ...) with their
          detailed tasks (start, end, owner, done). If an event is the same as an existing one, use its id and give
          only the tasks to add or change; if a task is the same as an existing one, use its id to change it.
          A new event has id "".
          Each event also has a Material Plan (the materials that come in, with the quantity per day) and a Making
          Plan (the making, inspection and shipping steps, with the quantity per day): the "rows" quantity table,
          group "material" or "making", qty = {"yyyy-MM-dd": number}. A material arrival that has a date and a
          quantity goes to a Material Plan row: use the existing row whose name fits (e.g. "F-PCB (Free) 720ea
          arrives" -> the row F-PCB-G778B, that date, 720); the same material arriving several times on one day is
          summed; if no row fits, add a new row (id ""). If a date cell only has an arrival mark (입고, GN입고,
          도착, IN, ✓ ...) and no quantity, put the quantity of that row on that date: use the Total (합계) column
          of the same row, else the planned quantity written in the row (S-code ...). When one row has several
          dates but only one quantity, the quantity written in a date cell comes first and the remaining quantity
          goes to the dates without one. Never drop an arrival date: only when the quantity is really unknown, put
          1 and say in the note which material and date, e.g. "quantity unknown, 1 put in". Check every material
          row of the table and answer all rows that have an arrival date; a row without a date (e.g. "잔여분
          사용") is not put in. An Excel material name (e.g. "G778B UPPER (PV2)") that fits an existing row name
          (e.g. ENCLOSURE-G778B-UPPER) uses that row id; read the category and name columns together
          (ENCLOSURE + UPPER). Making, inspection or shipping with a quantity goes to a
          Making Plan row (e.g. "SPK Assy 500 11-17" -> the row SPK Assy, 11-17, 500). A task is only an activity
          without a quantity (set-up, meeting, business trip ...). If the user asks to move or split existing tasks
          into the quantity table, put the quantities in rows and the ids of the moved tasks in "removeTasks"
          (do not keep a copy as a task). The date of a one-day task is the date of its quantity.
        - entries, kind "dev": a development history record or a decision.
        - entries, kind "issue": a problem. In development (stage "dev", phase ER/DR/DV/PV or "") a defect or design
          problem; in mass production (stage "mp", phase "MP") a mass-production problem.
          severity is Critical, Major or Minor; category, process and cause are short texts; due is a date.
        - entries, kind "review": the result of a review or meeting; result is Pass, Conditional or Fail.

        Answer ONLY one JSON object, no other text:
        {"info":{only the fields to change},"events":[{"id","name","start","end","tasks":[{"id","name","start","end","owner","done"}],
        "rows":[{"id","name","group","qty":{"yyyy-MM-dd":number}}],"removeTasks":["task id"]}],
        "entries":[{"kind","stage","date","title","body","owner","due","severity","result","phase","category","process","cause","images":[0]}],
        "note":"..."}
        - rows: the row id of an existing row (or "" for a new row), and in qty only the dates to set now (0 clears
          that day). "removeTasks" is [] when no task moves.
        - note: one to three English sentences saying what was put where, and why anything could not be placed.
        - Pictures may be screenshots of a table, a schedule or a mail, or photos of a defect. Read the text and
          tables in a picture like the typed text and place them the same way. When a picture is evidence for an
          issue, a history record or a review, put its number (0 based, as listed under PICTURES) in that
          entry's "images" array. "images" is [] when there is none.
        - Do not create entries that already exist in the current entries, and do not repeat unchanged info.
        - Dates are yyyy-MM-dd; a date without a year uses the year of today.
        - Keep the original language of the text; do not translate.
        - Create only what the text supports; do not invent anything.
        """;
}
