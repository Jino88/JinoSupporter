using System.Text.Json;
using System.Text.Json.Nodes;

namespace PortalGNLAB.Web.Services;

public sealed class ManualAiMaterial
{
    public string Material { get; set; } = "";
    public double QtyPerUnit { get; set; } = 1;
}

public sealed class ManualAiProcess
{
    public string Id { get; set; } = "";
    public string Type { get; set; } = "SUB";
    public string Name { get; set; } = "";
    public string LocalName { get; set; } = "";
    public List<string> NgItems { get; set; } = [];
    public List<ManualAiMaterial> Materials { get; set; } = [];
}

public sealed class ManualAiGroup
{
    public List<string> Materials { get; set; } = [];
    public List<ManualAiProcess> Processes { get; set; } = [];
}

public sealed record ManualAiRequest(string? Request, ManualAiGroup? Group);

public sealed record ManualAiResult(ManualAiGroup Group, string Note, string Model);

/// <summary>
/// Working > Model Info, hand-entered process table: the user types a free-form request (plus pasted Excel
/// cells) and the Codex API answers the changed table. Nothing is saved here; the client shows the result
/// and the user keeps it with Save.
/// </summary>
public sealed class WorkingManualAiService(CodexApiService codex)
{
    public const string Effort = "low";
    private static readonly string[] Types = ["SUB", "MAIN", "FUNCTION", "VISUAL"];

    private static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public async Task<ManualAiResult> EditAsync(string modelName, ManualAiRequest body, CancellationToken ct = default)
    {
        string request = body.Request?.Trim() ?? "";
        if (request.Length == 0) throw new ArgumentException("Type what to change.");
        if (!codex.IsConfigured) throw new ArgumentException("Codex API is not configured. Set the API key in Setting.");

        ManualAiGroup current = Normalize(body.Group);
        var text = new System.Text.StringBuilder();
        text.AppendLine($"Model name: {modelName}");
        text.AppendLine();
        text.AppendLine("=== CURRENT GROUP (JSON) ===");
        text.AppendLine(JsonSerializer.Serialize(current, Json));
        text.AppendLine();
        text.AppendLine("=== REQUEST ===");
        text.AppendLine(request);
        var content = new JsonArray { new JsonObject { ["type"] = "input_text", ["text"] = text.ToString() } };

        string model = codex.Model;
        CodexTextResult result = await codex.GenerateAsync(Instructions, content, model, Effort, 16000, ct);
        codex.RecordUsage(result.Usage);

        (ManualAiGroup group, string note) = ParseAnswer(result.Text);
        return new ManualAiResult(group, note, $"{model} · {Effort}");
    }

    /// <summary>Reads the AI answer leniently: strips code fences, takes the outermost {...}, then normalizes.</summary>
    public static (ManualAiGroup Group, string Note) ParseAnswer(string? text)
    {
        const string fail = "The AI answer could not be read. Try again.";
        string s = text ?? "";
        int a = s.IndexOf('{'), b = s.LastIndexOf('}');
        if (a < 0 || b <= a) throw new ArgumentException(fail);
        try
        {
            using JsonDocument doc = JsonDocument.Parse(s[a..(b + 1)]);
            JsonElement root = doc.RootElement;
            if (root.ValueKind != JsonValueKind.Object
                || !root.TryGetProperty("group", out JsonElement g)
                || g.ValueKind != JsonValueKind.Object)
                throw new ArgumentException(fail);
            ManualAiGroup group = g.Deserialize<ManualAiGroup>(Json) ?? throw new ArgumentException(fail);
            string note = root.TryGetProperty("note", out JsonElement n) && n.ValueKind == JsonValueKind.String
                ? n.GetString()?.Trim() ?? "" : "";
            return (Normalize(group), note);
        }
        catch (JsonException) { throw new ArgumentException(fail); }
    }

    /// <summary>Cleans a group: type, trimmed strings, no empty/duplicate items, qty &gt; 0, process materials listed.</summary>
    public static ManualAiGroup Normalize(ManualAiGroup? group)
    {
        var materials = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        void AddMaterial(string? m)
        {
            string t = (m ?? "").Trim();
            if (t.Length > 0 && seen.Add(t)) materials.Add(t);
        }
        foreach (string? m in group?.Materials ?? []) AddMaterial(m);

        var processes = new List<ManualAiProcess>();
        foreach (ManualAiProcess? p in group?.Processes ?? [])
        {
            if (p is null) continue;
            string type = (p.Type ?? "").Trim().ToUpperInvariant();
            var ng = new List<string>();
            var ngSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string? n in p.NgItems ?? [])
            {
                string t = (n ?? "").Trim();
                if (t.Length > 0 && ngSeen.Add(t)) ng.Add(t);
            }
            var mats = new List<ManualAiMaterial>();
            var matSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ManualAiMaterial? m in p.Materials ?? [])
            {
                string name = (m?.Material ?? "").Trim();
                if (name.Length == 0 || !matSeen.Add(name)) continue;
                double q = m!.QtyPerUnit;
                mats.Add(new ManualAiMaterial { Material = name, QtyPerUnit = double.IsFinite(q) && q > 0 ? q : 1 });
                AddMaterial(name);
            }
            processes.Add(new ManualAiProcess
            {
                Id = (p.Id ?? "").Trim(),
                Type = Types.Contains(type) ? type : "SUB",
                Name = (p.Name ?? "").Trim(),
                LocalName = (p.LocalName ?? "").Trim(),
                NgItems = ng,
                Materials = mats,
            });
        }
        return new ManualAiGroup { Materials = materials, Processes = processes };
    }

    private const string Instructions = """
        You edit the process table of one manufacturing model. The input is the model name, the current
        group as JSON ({"materials":[...],"processes":[{"id","type","name","localName","ngItems":[...],
        "materials":[{"material","qtyPerUnit"}]}]}) and the user's request. The request may contain cells
        copied from Excel (tab separated). For example a Daily NG Rate header: one row with the processes
        ("N. NAME" then a line break and the local name) and one row of "Input / NG items / NG rate (%)".
        Read it as: the process number is dropped, the name is NAME, the second line is localName, and the
        NG items of a process are the cells after its Input up to the next Input, without "NG rate (%)"
        and without blank cells.

        Answer ONLY one JSON object, no other text: {"group":{...same shape as the input...},"note":"..."}.
        - note: one or two English sentences saying what was changed, or why nothing could be changed.
        - Change only what the request asks for; keep everything else exactly as it is.
        - Keep the id of every existing process. New processes get id "".
        - Every material used by a process must be in the group's materials list; if the request names
          new materials, add them to the list.
        - qtyPerUnit is a number greater than 0.
        - type is one of SUB, MAIN, FUNCTION, VISUAL.
        - Keep the original language of all names; do not translate.
        """;
}
