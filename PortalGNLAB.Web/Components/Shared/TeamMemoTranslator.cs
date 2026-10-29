using PortalGNLAB.Web.Services;

namespace PortalGNLAB.Web.Components.Shared;

/// <summary>
/// Per-page translation map for the team memo. <see cref="Refresh"/> paints cached translations at once
/// and fetches the rest in the background; <see cref="T"/> returns the translation or the original.
/// </summary>
public sealed class TeamMemoTranslator(TeamMemoTranslationService service, Func<Task> redraw)
{
    private Dictionary<string, string> _map = new(StringComparer.Ordinal);
    private int _generation;

    public string Lang { get; private set; } = TeamMemoLanguage.Original;
    public bool Busy { get; private set; }
    public int Pending { get; private set; }
    public string? Error { get; private set; }

    public bool Active => Lang.Length > 0;

    public string T(string? text)
    {
        if (string.IsNullOrEmpty(text)) return text ?? "";
        return Active && _map.TryGetValue(text, out var t) ? t : text;
    }

    public void Refresh(string lang, IEnumerable<string?> texts)
    {
        int gen = ++_generation;
        Lang = lang;
        Error = null;
        if (!Active)
        {
            _map = new(StringComparer.Ordinal);
            Busy = false;
            return;
        }

        var list = texts.Where(t => !string.IsNullOrWhiteSpace(t)).Select(t => t!).Distinct(StringComparer.Ordinal).ToList();
        Dictionary<string, string> cached;
        try { cached = service.Cached(list, lang); }
        catch (Exception ex) { Error = ex.Message; return; }

        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var kv in cached) map[kv.Key] = kv.Value;
        _map = map;

        var missing = list.Where(t => !map.ContainsKey(t) && t.Any(char.IsLetter)).ToList();
        Pending = missing.Count;
        Busy = missing.Count > 0;
        if (!Busy) return;

        _ = Task.Run(async () =>
        {
            try
            {
                var fetched = await service.TranslateAsync(missing, lang);
                if (gen != _generation) return;
                var merged = new Dictionary<string, string>(_map, StringComparer.Ordinal);
                foreach (var kv in fetched) merged[kv.Key] = kv.Value;
                _map = merged;
            }
            catch (Exception ex)
            {
                if (gen == _generation) Error = "Translation failed: " + ex.Message;
            }
            finally
            {
                if (gen == _generation)
                {
                    Busy = false;
                    Pending = 0;
                }
            }
            await redraw();
        });
    }
}
