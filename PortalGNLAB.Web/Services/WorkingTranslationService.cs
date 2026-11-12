using System.Text.Json;

namespace PortalGNLAB.Web.Services;

/// <summary>A request for translations of Working texts: the language and the trimmed, distinct, non-empty texts.</summary>
public sealed record WorkingTranslationRequest(string Lang, List<string> Texts);

/// <summary>
/// Translations of the Working page (English / Korean / Vietnamese). Everything goes through
/// <see cref="TeamMemoTranslationService"/> and its hash cache (all three languages are stored by one call), so no
/// second store exists. A request answers from the cache at once; what is missing is translated in the background,
/// one job at a time (a request that arrives meanwhile joins the queue). Singleton: it owns the queue and the status.
/// </summary>
public sealed class WorkingTranslationService(WorkingService working, IServiceScopeFactory scopes, ILogger<WorkingTranslationService> log)
{
    public const int MaxTexts = 3000;
    private const int BatchSize = 50;

    public static readonly string[] Languages = TeamMemoLanguage.Targets;

    private readonly object _gate = new();
    private readonly HashSet<string> _pending = new(StringComparer.Ordinal);
    private bool _running;
    private int _done;
    private int _total;
    private string? _message;

    // ── Pure helpers (unit tested) ─────────────────────────────────────────────

    /// <summary>True when the text has a letter; numbers, dates and symbols alone are never translated.</summary>
    public static bool HasLetters(string? text) => !string.IsNullOrWhiteSpace(text) && text.Any(char.IsLetter);

    /// <summary>Trims, drops empty and duplicate texts, keeps the first <see cref="MaxTexts"/>; rejects a language other than en / ko / vi.</summary>
    public static WorkingTranslationRequest Normalize(string? lang, IEnumerable<string?>? texts)
    {
        string l = (lang ?? "").Trim().ToLowerInvariant();
        if (!Languages.Contains(l)) throw new ArgumentException("lang must be en, ko or vi.");
        var list = (texts ?? []).Select(t => (t ?? "").Trim()).Where(t => t.Length > 0)
            .Distinct(StringComparer.Ordinal).Take(MaxTexts).ToList();
        return new(l, list);
    }

    /// <summary>
    /// Every text a person wrote in the Working page: entry title / body / final result / category / process / cause, progress
    /// notes, setup row process name and machine note, model info note, and the Dev Plan's event, task and Making Plan row names
    /// and task notes. Owners, customers, lines and codes are left out.
    /// </summary>
    public static List<string> CollectTexts(WorkingService working)
    {
        var texts = new List<string?>();
        foreach (string model in working.ListDataModels())
        {
            JsonElement info = working.GetModel(model).Info;
            if (info.ValueKind == JsonValueKind.Object)
            {
                if (info.TryGetProperty("note", out JsonElement note) && note.ValueKind == JsonValueKind.String) texts.Add(note.GetString());
                if (info.TryGetProperty("devPlan", out JsonElement plan)) CollectDevPlan(plan, texts);
            }
            foreach (WorkingEntry e in working.ListEntries(model))
            {
                texts.AddRange([e.Title, e.Body, e.FinalResult, e.Category, e.Process, e.Cause]);
                foreach (WorkingEntryLog l in e.Logs) texts.Add(l.Text);
            }
            foreach (WorkingSetupRow r in working.ListSetupRows(model)) texts.AddRange([r.ProcessName, r.MachineNote]);
        }
        return Clean(texts);
    }

    /// <summary>The names and notes of a Dev Plan (a JSON object or a string holding one); material rows are not texts.</summary>
    public static List<string> CollectDevPlanTexts(JsonElement plan)
    {
        var texts = new List<string?>();
        CollectDevPlan(plan, texts);
        return Clean(texts);
    }

    private static void CollectDevPlan(JsonElement plan, List<string?> texts)
    {
        try
        {
            using JsonDocument? parsed = plan.ValueKind == JsonValueKind.String ? JsonDocument.Parse(plan.GetString() ?? "") : null;
            JsonElement root = parsed?.RootElement ?? plan;
            if (root.ValueKind != JsonValueKind.Object || !root.TryGetProperty("events", out JsonElement events) || events.ValueKind != JsonValueKind.Array) return;
            foreach (JsonElement ev in events.EnumerateArray())
            {
                if (ev.ValueKind != JsonValueKind.Object) continue;
                texts.Add(Str(ev, "name"));
                if (ev.TryGetProperty("tasks", out JsonElement tasks) && tasks.ValueKind == JsonValueKind.Array)
                    foreach (JsonElement t in tasks.EnumerateArray())
                    {
                        if (t.ValueKind != JsonValueKind.Object) continue;
                        texts.Add(Str(t, "name"));
                        texts.Add(Str(t, "note"));
                    }
                if (ev.TryGetProperty("rows", out JsonElement rows) && rows.ValueKind == JsonValueKind.Array)
                    foreach (JsonElement r in rows.EnumerateArray())
                        if (r.ValueKind == JsonValueKind.Object && string.Equals(Str(r, "group")?.Trim(), "making", StringComparison.OrdinalIgnoreCase))
                            texts.Add(Str(r, "name"));
            }
        }
        catch (JsonException) { }
    }

    private static string? Str(JsonElement o, string name) =>
        o.TryGetProperty(name, out JsonElement v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;

    private static List<string> Clean(IEnumerable<string?> texts) =>
        texts.Select(t => (t ?? "").Trim()).Where(HasLetters).Distinct(StringComparer.Ordinal).ToList();

    // ── Requests ───────────────────────────────────────────────────────────────

    public bool IsConfigured
    {
        get
        {
            using IServiceScope scope = scopes.CreateScope();
            return scope.ServiceProvider.GetRequiredService<TeamMemoTranslationService>().IsConfigured;
        }
    }

    /// <summary>The cached translations at once; what is missing starts (or joins) the background translation.</summary>
    public object Translations(string? lang, IEnumerable<string?>? texts)
    {
        WorkingTranslationRequest req = Normalize(lang, texts);
        var translations = new Dictionary<string, string>(StringComparer.Ordinal);
        var wanted = new List<string>();
        foreach (string t in req.Texts)
            if (HasLetters(t)) wanted.Add(t);
            else translations[t] = t;

        using (IServiceScope scope = scopes.CreateScope())
            foreach (var (text, tr) in scope.ServiceProvider.GetRequiredService<TeamMemoTranslationService>().Cached(wanted, req.Lang))
                translations[text] = tr;

        var missing = wanted.Where(t => !translations.ContainsKey(t)).ToList();
        if (missing.Count > 0) Enqueue(missing);
        return new { lang = req.Lang, translations, missing = missing.Count, running = Status().Running };
    }

    public (bool Running, int Done, int Total, string? Message) Status()
    {
        lock (_gate) return (_running, _done, _total, _message);
    }

    /// <summary>Queues the texts for translation (Codex must be configured); starts the job unless one is already running.</summary>
    public void Enqueue(IEnumerable<string> texts)
    {
        if (!IsConfigured) return;
        lock (_gate)
        {
            int before = _pending.Count;
            foreach (string t in texts) _pending.Add(t);
            _total += _pending.Count - before;
            if (_pending.Count == 0 || _running) return;
            _running = true;
            _message = null;
        }
        _ = Task.Run(RunAsync);
    }

    /// <summary>Queues everything in Working that is not translated into all three languages yet; returns how many.</summary>
    public int EnqueueAllMissing()
    {
        if (!IsConfigured) return 0;
        List<string> missing = MissingOf(CollectTexts(working));
        Enqueue(missing);
        return missing.Count;
    }

    private List<string> MissingOf(List<string> texts)
    {
        using IServiceScope scope = scopes.CreateScope();
        var cache = scope.ServiceProvider.GetRequiredService<TeamMemoTranslationService>();
        var have = Languages.Select(l => cache.Cached(texts, l)).ToList();
        return texts.Where(t => have.Any(h => !h.ContainsKey(t))).ToList();
    }

    private async Task RunAsync()
    {
        int translated = 0;
        while (true)
        {
            List<string> batch;
            lock (_gate)
            {
                if (_pending.Count == 0)
                {
                    _running = false;
                    _message = translated == 0 ? null : $"Translated {translated} texts at {DateTime.Now:HH:mm}.";
                    return;
                }
                batch = _pending.Take(BatchSize).ToList();
                _pending.ExceptWith(batch);
            }
            try
            {
                using IServiceScope scope = scopes.CreateScope();
                await scope.ServiceProvider.GetRequiredService<TeamMemoTranslationService>().TranslateAsync(batch, Languages[0]);
                translated += batch.Count;
            }
            catch (Exception ex)
            {
                log.LogWarning(ex, "Working translation of {Count} texts failed", batch.Count);
                lock (_gate) _message = "Translation failed: " + ex.Message;
            }
            lock (_gate) _done += batch.Count;
        }
    }
}

/// <summary>
/// Keeps the Working page pre-translated: once a minute after start and again once Working data has been quiet for 20 seconds
/// after a change, every text that is not translated yet is queued. Off when Setting "Working:AutoTranslate" is 0; skipped
/// quietly when Codex is not configured; failures are only logged.
/// </summary>
public sealed class WorkingTranslationBackgroundService(
    WorkingService working, WorkingTranslationService translations, WebRepository repo, ILogger<WorkingTranslationBackgroundService> log)
    : BackgroundService
{
    public const string AutoTranslateKey = "Working:AutoTranslate";
    private static readonly TimeSpan StartDelay = TimeSpan.FromMinutes(1);
    private static readonly TimeSpan QuietPeriod = TimeSpan.FromSeconds(20);
    private static readonly TimeSpan PollInterval = TimeSpan.FromSeconds(5);

    private DateTime? _dirtyAt;

    private void MarkDirty() => _dirtyAt = DateTime.Now;

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        try { await Task.Delay(StartDelay, stoppingToken); } catch (OperationCanceledException) { return; }
        working.Changed += MarkDirty;
        try
        {
            _dirtyAt = DateTime.Now - QuietPeriod - TimeSpan.FromSeconds(1); // the pass after start
            while (!stoppingToken.IsCancellationRequested)
            {
                if (_dirtyAt is { } dirty && DateTime.Now - dirty > QuietPeriod)
                {
                    _dirtyAt = null;
                    Pass();
                }
                try { await Task.Delay(PollInterval, stoppingToken); } catch (OperationCanceledException) { return; }
            }
        }
        finally
        {
            working.Changed -= MarkDirty;
        }
    }

    private void Pass()
    {
        try
        {
            if (repo.GetSetting(AutoTranslateKey) == "0" || !translations.IsConfigured) return;
            int n = translations.EnqueueAllMissing();
            if (n > 0) log.LogInformation("Queued {Count} Working texts for translation", n);
        }
        catch (Exception ex)
        {
            log.LogWarning(ex, "Working pre-translation failed; will retry after the next change");
        }
    }
}
