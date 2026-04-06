using System.Text.Json;
using System.IO;
using CustomKeyboardCSharp.Models;

namespace CustomKeyboardCSharp.Services;

public sealed class SettingsService
{
    public const int MinTimeoutSeconds = 5;
    public const int MaxTimeoutSeconds = 180;
    private readonly string _settingsPath;

    public SettingsService()
    {
        var baseDirectory = CustomKeyboardPathResolver.GetAppDataDirectory();
        Directory.CreateDirectory(baseDirectory);
        _settingsPath = Path.Combine(baseDirectory, "settings.json");
    }

    public AppSettings Load()
    {
        if (!File.Exists(_settingsPath))
        {
            var defaults = new AppSettings();
            Save(defaults);
            return defaults;
        }

        try
        {
            var json = File.ReadAllText(_settingsPath);
            var settings = JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
            settings.TranslationTimeoutSeconds = ClampTimeout(settings.TranslationTimeoutSeconds);
            return settings;
        }
        catch
        {
            return new AppSettings();
        }
    }

    public void Save(AppSettings settings)
    {
        settings.TranslationTimeoutSeconds = ClampTimeout(settings.TranslationTimeoutSeconds);
        var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(_settingsPath, json);
    }

    public static int ClampTimeout(int timeoutSeconds) =>
        Math.Clamp(timeoutSeconds, MinTimeoutSeconds, MaxTimeoutSeconds);
}
