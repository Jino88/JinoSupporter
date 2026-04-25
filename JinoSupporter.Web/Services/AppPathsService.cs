using System.Text.Json;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Central, persistent registry of every file/folder path the JinoSupporter.Web app
/// reads or writes. Backed by <c>%LOCALAPPDATA%\JinoWorkHost\webapp-paths.json</c>.
/// Edited via the Admin Paths page; consumed by services + Program.cs at startup.
/// </summary>
public sealed class AppPathsConfig
{
    public string MainDbPath                { get; set; } = string.Empty;
    public string ScheduleDbPath            { get; set; } = string.Empty;
    public string NgRateDbSaveDirectory     { get; set; } = string.Empty;
    public string NgRateRoutingFilePath     { get; set; } = string.Empty;
    public string NgRateReasonFilePath      { get; set; } = string.Empty;
    public string NgRateSettingsDbDirectory { get; set; } = string.Empty;
    public string ModelBmesJsonFolder       { get; set; } = string.Empty;
}

public sealed class AppPathsService
{
    private static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "JinoWorkHost");

    private static readonly string ConfigFile = Path.Combine(ConfigDir, "webapp-paths.json");

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
    };

    private AppPathsConfig _current;

    public AppPathsService()
    {
        _current = LoadOrCreate();
    }

    /// <summary>The currently effective path configuration. Always non-null;
    /// any blank fields are filled with platform defaults.</summary>
    public AppPathsConfig Current => _current;

    public string ConfigFilePath => ConfigFile;

    /// <summary>Baseline used for initial population and for any field the user leaves blank.
    /// Honors the WPF app's configured paths (workhost-settings.json) when present, so
    /// the Web app keeps reading the same DBs the WPF app already writes.</summary>
    public static AppPathsConfig Defaults()
    {
        string localApp = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        string ngBase   = @"D:\000. MyWorks\000. 일일업무\04. DB\01. NGRATE";
        string ngBmes   = Path.Combine(ngBase, "ModelBmes");

        // Honor existing WPF settings so the Web app stays in sync with the WPF app's DB.
        string mainDb = WpfSettingsReader.TryGetDatabasePath()
                        ?? Path.Combine(localApp, "JinoWorkHost", "process-review.db");
        string schedDb = WpfSettingsReader.TryGetScheduleDatabasePath()
                        ?? Path.Combine(localApp, "JinoWorkHost", "schedule.db");

        return new AppPathsConfig
        {
            MainDbPath                = mainDb,
            ScheduleDbPath            = schedDb,
            NgRateDbSaveDirectory     = ngBase,
            NgRateRoutingFilePath     = Path.Combine(ngBase, "Routing.txt"),
            NgRateReasonFilePath      = Path.Combine(ngBase, "reason.txt"),
            NgRateSettingsDbDirectory = ngBmes,
            ModelBmesJsonFolder       = ngBmes,
        };
    }

    private static AppPathsConfig LoadOrCreate()
    {
        AppPathsConfig def = Defaults();
        if (!File.Exists(ConfigFile)) return def;
        try
        {
            string json = File.ReadAllText(ConfigFile);
            var loaded = JsonSerializer.Deserialize<AppPathsConfig>(json, JsonOpts);
            if (loaded is null) return def;
            return Merge(def, loaded);
        }
        catch
        {
            // Corrupt or unreadable file → fall back to defaults silently.
            return def;
        }
    }

    public void Save(AppPathsConfig cfg)
    {
        var merged = Merge(Defaults(), cfg ?? new AppPathsConfig());
        Directory.CreateDirectory(ConfigDir);
        string json = JsonSerializer.Serialize(merged, JsonOpts);
        File.WriteAllText(ConfigFile, json);
        _current = merged;
    }

    /// <summary>Returns a copy where each blank field in <paramref name="user"/> is replaced
    /// with the corresponding value from <paramref name="defaults"/>.</summary>
    private static AppPathsConfig Merge(AppPathsConfig defaults, AppPathsConfig user) => new()
    {
        MainDbPath                = Pick(user.MainDbPath,                defaults.MainDbPath),
        ScheduleDbPath            = Pick(user.ScheduleDbPath,            defaults.ScheduleDbPath),
        NgRateDbSaveDirectory     = Pick(user.NgRateDbSaveDirectory,     defaults.NgRateDbSaveDirectory),
        NgRateRoutingFilePath     = Pick(user.NgRateRoutingFilePath,     defaults.NgRateRoutingFilePath),
        NgRateReasonFilePath      = Pick(user.NgRateReasonFilePath,      defaults.NgRateReasonFilePath),
        NgRateSettingsDbDirectory = Pick(user.NgRateSettingsDbDirectory, defaults.NgRateSettingsDbDirectory),
        ModelBmesJsonFolder       = Pick(user.ModelBmesJsonFolder,       defaults.ModelBmesJsonFolder),
    };

    private static string Pick(string? user, string fallback) =>
        string.IsNullOrWhiteSpace(user) ? fallback : user.Trim();
}
