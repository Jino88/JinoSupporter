using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Reads settings from the WPF app's settings file.
/// The actual settings file path is resolved through settings-bootstrap.json,
/// which the WPF app writes to %LOCALAPPDATA%\JinoWorkHost\.
/// </summary>
public static class WpfSettingsReader
{
    private static readonly string BootstrapPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "JinoWorkHost", "settings-bootstrap.json");

    private static readonly string FallbackSettingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "JinoWorkHost", "workhost-settings.json");

    /// <summary>
    /// Resolves the actual settings file path by reading settings-bootstrap.json.
    /// Falls back to the default path if the bootstrap file doesn't exist.
    /// </summary>
    private static string ResolveSettingsPath()
    {
        try
        {
            if (!File.Exists(BootstrapPath)) return FallbackSettingsPath;
            using JsonDocument boot = JsonDocument.Parse(File.ReadAllText(BootstrapPath));
            if (!boot.RootElement.TryGetProperty("SettingsFilePath", out JsonElement pathEl)) return FallbackSettingsPath;
            string? path = pathEl.GetString();
            return string.IsNullOrWhiteSpace(path) ? FallbackSettingsPath : path;
        }
        catch
        {
            return FallbackSettingsPath;
        }
    }

    private static JsonDocument? TryOpenSettings()
    {
        string path = ResolveSettingsPath();
        if (!File.Exists(path)) return null;
        return JsonDocument.Parse(File.ReadAllText(path));
    }

    /// <summary>
    /// Returns the custom DB path set by the WPF app, or null if not configured.
    /// Path is stored in DataInference.DatabasePath inside workhost-settings.json.
    /// </summary>
    public static string? TryGetDatabasePath()
    {
        try
        {
            using JsonDocument? doc = TryOpenSettings();
            if (doc is null) return null;

            if (!doc.RootElement.TryGetProperty("DataInference", out JsonElement diEl)) return null;
            if (!diEl.TryGetProperty("DatabasePath", out JsonElement pathEl))           return null;

            string? path = pathEl.GetString();
            return string.IsNullOrWhiteSpace(path) ? null : path;
        }
        catch { return null; }
    }

    /// <summary>
    /// Returns the custom Schedule DB path set by the WPF app, or null if not configured.
    /// Path is stored in Schedule.DatabasePath inside workhost-settings.json.
    /// </summary>
    public static string? TryGetScheduleDatabasePath()
    {
        try
        {
            using JsonDocument? doc = TryOpenSettings();
            if (doc is null) return null;

            if (!doc.RootElement.TryGetProperty("Schedule", out JsonElement schedEl)) return null;
            if (!schedEl.TryGetProperty("DatabasePath", out JsonElement pathEl))      return null;

            string? path = pathEl.GetString();
            return string.IsNullOrWhiteSpace(path) ? null : path;
        }
        catch { return null; }
    }

    /// <summary>Returns the decrypted Claude API key, or null if unavailable.</summary>
    public static string? TryGetClaudeApiKey()
    {
        try
        {
            using JsonDocument? doc = TryOpenSettings();
            if (doc is null) return null;

            if (!doc.RootElement.TryGetProperty("Claude", out JsonElement claudeEl)) return null;
            if (!claudeEl.TryGetProperty("EncryptedApiKey", out JsonElement keyEl))  return null;

            string? encrypted = keyEl.GetString();
            if (string.IsNullOrWhiteSpace(encrypted)) return null;

            byte[] protectedBytes = Convert.FromBase64String(encrypted);
            byte[] plaintextBytes = ProtectedData.Unprotect(protectedBytes, null, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(plaintextBytes);
        }
        catch { return null; }
    }
}
