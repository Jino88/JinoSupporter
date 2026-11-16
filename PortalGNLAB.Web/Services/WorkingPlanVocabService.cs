using System.Text.Json;
using System.Text.Json.Serialization;

namespace PortalGNLAB.Web.Services;

public sealed record PlanVocabItem(string Code, string Name, List<string> Aliases);

public sealed record PlanVocab(List<PlanVocabItem> Stages, List<PlanVocabItem> Tasks, Dictionary<string, List<string>> Templates);

/// <summary>
/// The Dev Plan standard names (stages, tasks, their other names and the tasks each stage starts with), one list for
/// every model. Kept as JSON in the NgRateSettingsService setting <see cref="SettingKey"/>; null until saved.
/// </summary>
public sealed class WorkingPlanVocabService
{
    public const string SettingKey = "Working:PlanVocab";
    public const int MaxBytes = 256 * 1024;
    public const int MaxItems = 200;
    public const int MaxAliases = 30;
    public const int MaxLength = 100;

    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web)
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.Never,
    };

    private readonly Func<string?> _getSetting;
    private readonly Action<string> _setSetting;

    public WorkingPlanVocabService(NgRateSettingsService settings)
        : this(() => settings.GetSetting(SettingKey), v => settings.SetSetting(SettingKey, v)) { }

    public WorkingPlanVocabService(Func<string?> getSetting, Action<string> setSetting)
    {
        _getSetting = getSetting;
        _setSetting = setSetting;
    }

    /// <summary>The saved list, or null until one is saved (or when the stored text is unreadable).</summary>
    public PlanVocab? Get()
    {
        string? text = _getSetting();
        if (string.IsNullOrWhiteSpace(text)) return null;
        try { return JsonSerializer.Deserialize<PlanVocab>(text, Json); }
        catch (JsonException) { return null; }
    }

    /// <summary>Cleans up and saves the list in <paramref name="body"/> ({stages, tasks, templates}); throws ArgumentException when it is unusable.</summary>
    public PlanVocab Save(string body)
    {
        if (System.Text.Encoding.UTF8.GetByteCount(body ?? "") > MaxBytes)
            throw new ArgumentException($"The list is larger than {MaxBytes / 1024} KB.");
        PlanVocab clean = Clean(Parse(body ?? ""));
        _setSetting(JsonSerializer.Serialize(clean, Json));
        return clean;
    }

    private static PlanVocab Parse(string body)
    {
        try
        {
            using JsonDocument doc = JsonDocument.Parse(body);
            JsonElement root = doc.RootElement;
            if (root.ValueKind != JsonValueKind.Object
                || !root.TryGetProperty("stages", out JsonElement stages) || stages.ValueKind != JsonValueKind.Array
                || !root.TryGetProperty("tasks", out JsonElement tasks) || tasks.ValueKind != JsonValueKind.Array)
                throw new ArgumentException("stages and tasks lists are required.");
            var templates = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            if (root.TryGetProperty("templates", out JsonElement t) && t.ValueKind == JsonValueKind.Object)
                foreach (JsonProperty p in t.EnumerateObject())
                    templates[p.Name] = p.Value.ValueKind == JsonValueKind.Array ? Strings(p.Value) : [];
            return new(Items(stages), Items(tasks), templates);
        }
        catch (JsonException) { throw new ArgumentException("The list is not valid JSON."); }
    }

    private static List<string> Strings(JsonElement array) =>
        array.EnumerateArray().Select(e => e.ValueKind == JsonValueKind.String ? e.GetString() ?? "" : "").ToList();

    private static List<PlanVocabItem> Items(JsonElement array) =>
        array.EnumerateArray().Select(e =>
        {
            if (e.ValueKind != JsonValueKind.Object) return new PlanVocabItem("", "", []);
            string Str(string n) => e.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() ?? "" : "";
            List<string> aliases = e.TryGetProperty("aliases", out var a) && a.ValueKind == JsonValueKind.Array ? Strings(a) : [];
            return new PlanVocabItem(Str("code"), Str("name"), aliases);
        }).ToList();

    private static string Cut(string s)
    {
        s = (s ?? "").Trim();
        return s.Length > MaxLength ? s[..MaxLength].TrimEnd() : s;
    }

    private static List<PlanVocabItem> CleanItems(List<PlanVocabItem> items)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<PlanVocabItem>();
        foreach (PlanVocabItem i in items)
        {
            string code = Cut(i.Code), name = Cut(i.Name);
            if (code.Length == 0 || name.Length == 0 || !seen.Add(code)) continue;
            var aliases = (i.Aliases ?? []).Select(Cut).Where(a => a.Length > 0)
                .Distinct(StringComparer.Ordinal).Take(MaxAliases).ToList();
            result.Add(new(code, name, aliases));
            if (result.Count >= MaxItems) break;
        }
        return result;
    }

    public static PlanVocab Clean(PlanVocab v)
    {
        var stages = CleanItems(v.Stages ?? []);
        var tasks = CleanItems(v.Tasks ?? []);
        var taskCodes = tasks.Select(t => t.Code).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var raw = v.Templates ?? [];
        var templates = new Dictionary<string, List<string>>();
        foreach (PlanVocabItem s in stages)
        {
            var list = raw.FirstOrDefault(kv => string.Equals(kv.Key, s.Code, StringComparison.OrdinalIgnoreCase)).Value ?? [];
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            templates[s.Code] = list.Select(Cut).Where(c => taskCodes.Contains(c) && seen.Add(c))
                .Select(c => tasks.First(t => string.Equals(t.Code, c, StringComparison.OrdinalIgnoreCase)).Code).ToList();
        }
        return new(stages, tasks, templates);
    }
}
