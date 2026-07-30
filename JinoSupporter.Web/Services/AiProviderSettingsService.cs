using System.Text.Json;
using System.Text.Json.Serialization;

namespace JinoSupporter.Web.Services;

public sealed class AiProviderSettingsSnapshot
{
    public bool ClaudeCliEnabled { get; set; } = false;
    public bool ClaudeApiEnabled { get; set; } = true;
    public bool CodexCliEnabled  { get; set; } = true;
    public bool CodexApiEnabled  { get; set; } = true;
}

public sealed class AiProviderSettingsService
{
    private readonly object _gate = new();
    private readonly JsonSerializerOptions _jsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.Never,
    };

    private AiProviderSettingsSnapshot? _cached;
    private DateTime _lastWriteUtc = DateTime.MinValue;

    public AiProviderSettingsService()
    {
        SettingsPath = Path.Combine(FindRepoRoot(), "ai-provider-settings.json");
        EnsureFile();
    }

    public string SettingsPath { get; }

    public bool ClaudeCliEnabled => GetSnapshot().ClaudeCliEnabled;
    public bool ClaudeApiEnabled => GetSnapshot().ClaudeApiEnabled;
    public bool CodexCliEnabled  => GetSnapshot().CodexCliEnabled;
    public bool CodexApiEnabled  => GetSnapshot().CodexApiEnabled;

    public AiProviderSettingsSnapshot GetSnapshot()
    {
        lock (_gate)
        {
            DateTime currentWriteUtc = GetLastWriteUtcNoLock();
            if (_cached is null || currentWriteUtc != _lastWriteUtc)
            {
                _cached = LoadNoLock();
                _lastWriteUtc = currentWriteUtc;
            }
            return Clone(_cached);
        }
    }

    public void Save(AiProviderSettingsSnapshot snapshot)
    {
        lock (_gate)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath) ?? ".");
            string json = JsonSerializer.Serialize(snapshot, _jsonOptions);
            File.WriteAllText(SettingsPath, json, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            _cached = Clone(snapshot);
            _lastWriteUtc = GetLastWriteUtcNoLock();
        }
    }

    public string ProviderDisabledMessage(string provider)
        => $"{provider} is disabled in ai-provider-settings.json.";

    private void EnsureFile()
    {
        lock (_gate)
        {
            if (File.Exists(SettingsPath))
            {
                _cached = LoadNoLock();
                _lastWriteUtc = GetLastWriteUtcNoLock();
                return;
            }

            _cached = new AiProviderSettingsSnapshot();
            Save(_cached);
        }
    }

    private AiProviderSettingsSnapshot LoadNoLock()
    {
        try
        {
            if (!File.Exists(SettingsPath)) return new AiProviderSettingsSnapshot();

            string json = File.ReadAllText(SettingsPath);
            return JsonSerializer.Deserialize<AiProviderSettingsSnapshot>(
                       json,
                       new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
                   ?? new AiProviderSettingsSnapshot();
        }
        catch
        {
            return new AiProviderSettingsSnapshot();
        }
    }

    private static AiProviderSettingsSnapshot Clone(AiProviderSettingsSnapshot source)
        => new()
        {
            ClaudeCliEnabled = source.ClaudeCliEnabled,
            ClaudeApiEnabled = source.ClaudeApiEnabled,
            CodexCliEnabled  = source.CodexCliEnabled,
            CodexApiEnabled  = source.CodexApiEnabled,
        };

    private static string FindRepoRoot()
    {
        string dir = AppContext.BaseDirectory;
        for (int i = 0; i < 10; i++)
        {
            if (File.Exists(Path.Combine(dir, "JinoSupporter.sln"))) return dir;
            string? parent = Path.GetDirectoryName(dir.TrimEnd('\\', '/'));
            if (string.IsNullOrWhiteSpace(parent) || parent == dir) break;
            dir = parent;
        }

        string current = Directory.GetCurrentDirectory();
        dir = current;
        for (int i = 0; i < 10; i++)
        {
            if (File.Exists(Path.Combine(dir, "JinoSupporter.sln"))) return dir;
            string? parent = Path.GetDirectoryName(dir.TrimEnd('\\', '/'));
            if (string.IsNullOrWhiteSpace(parent) || parent == dir) break;
            dir = parent;
        }

        return current;
    }

    private DateTime GetLastWriteUtcNoLock()
        => File.Exists(SettingsPath) ? File.GetLastWriteTimeUtc(SettingsPath) : DateTime.MinValue;
}
