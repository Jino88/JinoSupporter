using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Data.Sqlite;

namespace PortalGNLAB.Web.Services;

/// <summary>Display language of the team memo pages; <see cref="Original"/> shows the text as written.</summary>
public static class TeamMemoLanguage
{
    public const string Original = "";
    public const string English = "en";
    public const string Korean = "ko";
    public const string Vietnamese = "vi";

    public static readonly (string Code, string Label)[] All =
        [(Original, "Original"), (English, "EN"), (Korean, "KO"), (Vietnamese, "VI")];

    /// <summary>The languages every text is translated into, all in the same call.</summary>
    public static readonly string[] Targets = [English, Korean, Vietnamese];

    public static string Name(string code) => code switch
    {
        English => "English",
        Korean => "Korean",
        Vietnamese => "Vietnamese",
        _ => "",
    };
}

/// <summary>
/// Translates team memo text (titles, history, comments, AI summaries) for display with Codex Luna at
/// low effort. Each call returns English, Korean and Vietnamese together and all three are cached in the
/// memo DB under a hash of the source text — so a text is sent to the API once, switching language is
/// free, and an edited text (new hash) is simply translated again.
/// </summary>
public sealed class TeamMemoTranslationService(TeamMemoService memos, CodexApiService codex, ILogger<TeamMemoTranslationService> log)
{
    public const string Model = CodexApiService.LunaModel;
    public const string Effort = "low";
    // Three translations come back per item, so batches are smaller than a single-language call.
    private const int MaxItemsPerCall = 25;
    private const int MaxCharsPerCall = 6_000;

    private static readonly SemaphoreSlim CallGate = new(3, 3);
    private static readonly SemaphoreSlim AllGate = new(1, 1);
    private static readonly object SchemaGate = new();
    private static string? _schemaFor;

    /// <summary>State of the "translate every memo" job, shared by all users and the background pass.</summary>
    public static bool AllRunning => AllGate.CurrentCount == 0;
    public static int AllDone { get; private set; }
    public static int AllTotal { get; private set; }
    public static string? AllMessage { get; private set; }

    /// <summary>Raised when new translations were stored, so open pages can repaint.</summary>
    public static event Action? Translated;

    public bool IsConfigured => codex.IsConfigured;

    /// <summary>Cached translations into <paramref name="lang"/> only (no API call).</summary>
    public Dictionary<string, string> Cached(IEnumerable<string> texts, string lang)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        if (string.IsNullOrEmpty(lang)) return result;
        foreach (var (text, byLang) in Lookup(Distinct(texts)))
            if (byLang.TryGetValue(lang, out var t)) result[text] = t;
        return result;
    }

    /// <summary>Translations into <paramref name="lang"/>, translating (into all three languages) what is not cached yet.</summary>
    public async Task<Dictionary<string, string>> TranslateAsync(IEnumerable<string> texts, string lang, CancellationToken ct = default)
    {
        var all = Distinct(texts);
        await TranslateMissingAsync(all, null, ct);
        return Cached(all, lang);
    }

    /// <summary>
    /// Translates every text of every memo (folder names, titles, summaries, history, comments, AI summaries)
    /// that is not yet cached in all three languages. One run at a time; progress is in <see cref="AllDone"/>.
    /// </summary>
    public async Task<int> TranslateAllMemosAsync(CancellationToken ct = default)
    {
        if (!await AllGate.WaitAsync(0, ct)) return 0;
        try
        {
            AllDone = 0;
            AllTotal = 0;
            AllMessage = "Collecting memo text…";
            var texts = Distinct(CollectAllTexts());
            int sent = await TranslateMissingAsync(texts, (done, total) => { AllDone = done; AllTotal = total; AllMessage = null; }, ct);
            AllMessage = sent == 0
                ? $"All {texts.Count} texts are already translated."
                : $"Translated {sent} new texts ({texts.Count} in total) at {DateTime.Now:HH:mm}.";
            log.LogInformation("Team memo translate-all: {Sent} new of {Total}", sent, texts.Count);
            return sent;
        }
        catch (Exception ex)
        {
            AllMessage = "Translation failed: " + ex.Message;
            log.LogWarning(ex, "Team memo translate-all failed");
            return 0;
        }
        finally
        {
            AllGate.Release();
        }
    }

    private IEnumerable<string?> CollectAllTexts()
    {
        foreach (var f in memos.GetFolders()) yield return f.Name;
        foreach (var s in memos.ListMemos())
        {
            TeamMemo? m = memos.GetMemo(s.Id);
            if (m is null) continue;
            yield return m.Title;
            yield return m.Description;
            yield return m.AiSummary;
            foreach (var c in m.Comments) yield return c.Body;
            foreach (var e in m.Entries)
            {
                yield return e.Title;
                yield return e.Body;
                yield return e.AiSummary;
                foreach (var c in e.Comments) yield return c.Body;
            }
        }
    }

    /// <summary>Sends the texts missing any of the three languages; returns how many were sent.</summary>
    private async Task<int> TranslateMissingAsync(List<string> texts, Action<int, int>? progress, CancellationToken ct)
    {
        var have = Lookup(texts);
        var missing = texts.Where(t => !have.TryGetValue(t, out var byLang) || TeamMemoLanguage.Targets.Any(l => !byLang.ContainsKey(l))).ToList();
        progress?.Invoke(0, missing.Count);
        if (missing.Count == 0) return 0;
        if (!codex.IsConfigured) throw new InvalidOperationException("Codex API is not configured. Set the API key in Setting to translate.");

        var batches = new List<List<string>>();
        var current = new List<string>();
        int chars = 0;
        foreach (string t in missing)
        {
            if (current.Count > 0 && (current.Count >= MaxItemsPerCall || chars + t.Length > MaxCharsPerCall))
            {
                batches.Add(current);
                current = new List<string>();
                chars = 0;
            }
            current.Add(t);
            chars += t.Length;
        }
        if (current.Count > 0) batches.Add(current);

        int done = 0;
        await Task.WhenAll(batches.Select(async batch =>
        {
            var rows = await TranslateBatchAsync(batch, ct);
            Store(rows);
            progress?.Invoke(Interlocked.Add(ref done, batch.Count), missing.Count);
            try { Translated?.Invoke(); } catch { }
        }));
        return missing.Count;
    }

    private async Task<List<(string Source, Dictionary<string, string> ByLang)>> TranslateBatchAsync(List<string> items, CancellationToken ct)
    {
        const string instructions = """
            You translate text from a shared engineering work memo (manufacturing, quality, process review).
            The input is a JSON array of strings. Translate EVERY string into English, Korean and Vietnamese.
            Return ONLY a JSON object {"items":[{"en":"...","ko":"...","vi":"..."}, ...]} with exactly one
            object per input string, in the same order.

            Rules:
            - Natural, concise wording used by engineers on a factory floor.
            - If a string is already in a target language, return it unchanged for that language.
            - Keep product codes, model names, part numbers, process names in Latin letters (e.g. MSM-X526B, TIU-C11-20,
              SPK, NG, PV2, JIG), numbers, units, dates, file names and @mentions exactly as written.
            - Keep markdown (headings, bullets, tables, **bold**, `code`) and line breaks exactly; translate only the words.
            - Never add explanations or notes.
            """;
        var content = new JsonArray
        {
            new JsonObject { ["type"] = "input_text", ["text"] = JsonSerializer.Serialize(items) },
        };

        await CallGate.WaitAsync(ct);
        try
        {
            CodexTextResult r = await codex.GenerateAsync(instructions, content, Model, Effort,
                Math.Min(32000, 2000 + items.Sum(i => i.Length) * 8), ct);
            codex.RecordUsage(r.Usage);
            var output = ParseItems(r.Text);
            if (output is null || output.Count != items.Count)
            {
                log.LogWarning("Team memo translation returned {Got} items for {Asked}; retrying one by one", output?.Count, items.Count);
                if (items.Count == 1)
                    return output is { Count: 1 } ? [(items[0], output[0])] : [];
                var single = new List<(string, Dictionary<string, string>)>();
                foreach (string item in items) single.AddRange(await TranslateBatchAsync([item], ct));
                return single;
            }
            return items.Zip(output, (s, byLang) => (s, byLang)).ToList();
        }
        finally
        {
            CallGate.Release();
        }
    }

    private static List<Dictionary<string, string>>? ParseItems(string text)
    {
        text = text.Trim();
        if (text.StartsWith("```", StringComparison.Ordinal))
        {
            int nl = text.IndexOf('\n');
            int fence = text.LastIndexOf("```", StringComparison.Ordinal);
            if (nl >= 0 && fence > nl) text = text[(nl + 1)..fence];
        }
        try
        {
            using var doc = JsonDocument.Parse(text);
            JsonElement arr = doc.RootElement.ValueKind == JsonValueKind.Array
                ? doc.RootElement
                : doc.RootElement.TryGetProperty("items", out var items) ? items : default;
            if (arr.ValueKind != JsonValueKind.Array) return null;
            var list = new List<Dictionary<string, string>>();
            foreach (JsonElement e in arr.EnumerateArray())
            {
                var byLang = new Dictionary<string, string>(StringComparer.Ordinal);
                if (e.ValueKind == JsonValueKind.Object)
                    foreach (string lang in TeamMemoLanguage.Targets)
                        if (e.TryGetProperty(lang, out var v) && v.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(v.GetString()))
                            byLang[lang] = v.GetString()!;
                list.Add(byLang);
            }
            return list;
        }
        catch (JsonException)
        {
            return null;
        }
    }

    // ── Cache ──────────────────────────────────────────────────────────────────

    private Dictionary<string, Dictionary<string, string>> Lookup(List<string> texts)
    {
        var result = new Dictionary<string, Dictionary<string, string>>(StringComparer.Ordinal);
        if (texts.Count == 0) return result;
        using SqliteConnection conn = OpenCache();
        foreach (var chunk in texts.Chunk(300))
        {
            var byHash = chunk.GroupBy(Hash).ToDictionary(g => g.Key, g => g.First(), StringComparer.Ordinal);
            using var cmd = conn.CreateCommand();
            var names = new List<string>();
            int i = 0;
            foreach (string h in byHash.Keys)
            {
                names.Add($"@h{i}");
                cmd.Parameters.AddWithValue($"@h{i++}", h);
            }
            cmd.CommandText = $"SELECT Hash, Lang, Text FROM WebTranslations WHERE Hash IN ({string.Join(",", names)});";
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                string text = byHash[r.GetString(0)];
                if (!result.TryGetValue(text, out var byLang)) result[text] = byLang = new(StringComparer.Ordinal);
                byLang[r.GetString(1)] = r.GetString(2);
            }
        }
        return result;
    }

    private void Store(List<(string Source, Dictionary<string, string> ByLang)> rows)
    {
        if (rows.Count == 0) return;
        using SqliteConnection conn = OpenCache();
        using SqliteTransaction tx = conn.BeginTransaction();
        string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        foreach (var (source, byLang) in rows)
        {
            foreach (var (lang, text) in byLang)
            {
                using var cmd = conn.CreateCommand();
                cmd.Transaction = tx;
                cmd.CommandText = "INSERT OR REPLACE INTO WebTranslations (Hash, Lang, Text, Model, CreatedAt) VALUES (@h, @l, @t, @m, @now);";
                cmd.Parameters.AddWithValue("@h", Hash(source));
                cmd.Parameters.AddWithValue("@l", lang);
                cmd.Parameters.AddWithValue("@t", text);
                cmd.Parameters.AddWithValue("@m", $"{Model} · {Effort}");
                cmd.Parameters.AddWithValue("@now", now);
                cmd.ExecuteNonQuery();
            }
        }
        tx.Commit();
    }

    private SqliteConnection OpenCache()
    {
        SqliteConnection conn = memos.Open();
        lock (SchemaGate)
        {
            if (_schemaFor != memos.DbPath)
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = """
                    CREATE TABLE IF NOT EXISTS WebTranslations (
                        Hash      TEXT NOT NULL,
                        Lang      TEXT NOT NULL,
                        Text      TEXT NOT NULL,
                        Model     TEXT NOT NULL DEFAULT '',
                        CreatedAt TEXT NOT NULL,
                        PRIMARY KEY (Hash, Lang)
                    );
                    """;
                cmd.ExecuteNonQuery();
                _schemaFor = memos.DbPath;
            }
        }
        return conn;
    }

    private static List<string> Distinct(IEnumerable<string?> texts) =>
        texts.Where(t => !string.IsNullOrWhiteSpace(t) && t.Any(char.IsLetter)).Select(t => t!).Distinct(StringComparer.Ordinal).ToList();

    private static string Hash(string text) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text)));
}
