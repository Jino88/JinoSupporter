using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text;
using System.Windows.Input;

namespace WorkbenchHost.Infrastructure;

public static class WorkbenchSettingsStore
{
    public sealed class WorkbenchSettings
    {
        public string StorageRootDirectory { get; set; } = DefaultStorageRootDirectory;
        public Dictionary<string, CaptureHotkeySetting> CaptureHotkeys { get; set; } = new();
        public VideoConverterSetting VideoConverter { get; set; } = new();
        public MemoSetting Memo { get; set; } = new();
        public BmesSetting Bmes { get; set; } = new();
        public CodexSetting Codex { get; set; } = new();
        public ClaudeSetting Claude { get; set; } = new();
        public OpenAiSetting OpenAi { get; set; } = new();
        public DataInferenceSetting DataInference { get; set; } = new();
        public ScheduleSetting      Schedule      { get; set; } = new();
    }

    public sealed class CaptureHotkeySetting
    {
        public ModifierKeys Modifiers { get; set; }
        public string Key { get; set; } = string.Empty;
    }

    public sealed class VideoConverterSetting
    {
        public string FfmpegFolder { get; set; } = string.Empty;
        public string OutputFolder { get; set; } = string.Empty;
        public bool HevcOnlyMode { get; set; }
    }

    public sealed class MemoSetting
    {
        public string DatabasePath { get; set; } = string.Empty;
    }

    public sealed class BmesSetting
    {
        public string LoginId { get; set; } = string.Empty;
        public string Password { get; set; } = string.Empty;
        public string LastImportExportFilePath { get; set; } = string.Empty;
    }

    public sealed class CodexSetting
    {
        public bool AutoLoginEnabled { get; set; } = true;
        public string EncryptedCookieHeader { get; set; } = string.Empty;
    }

    public sealed class ClaudeSetting
    {
        public string EncryptedApiKey { get; set; } = string.Empty;
    }

    public sealed class OpenAiSetting
    {
        public string EncryptedApiKey { get; set; } = string.Empty;
    }

    public sealed class DataInferenceSetting
    {
        public string DatabasePath        { get; set; } = string.Empty;
        public string AnthropicSessionKey { get; set; } = string.Empty;
    }

    public sealed class ScheduleSetting
    {
        public string DatabasePath { get; set; } = string.Empty;
    }

    private sealed class SettingsBootstrap
    {
        public string SettingsFilePath { get; set; } = string.Empty;
    }

    private static readonly object SyncRoot = new();
    private static readonly string BootstrapDirectory =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "JinoWorkHost");

    private static readonly string BootstrapFilePathValue = Path.Combine(BootstrapDirectory, "settings-bootstrap.json");
    public static string DefaultStorageRootDirectory =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "JinoWorkHost");

    public static string DefaultSettingsFilePath =>
        Path.Combine(DefaultStorageRootDirectory, "workhost-settings.json");

    private static string? _settingsFilePath;
    private static WorkbenchSettings? _cachedSettings;

    public static string BootstrapFilePath => BootstrapFilePathValue;

    public static string SettingsFilePath
    {
        get
        {
            lock (SyncRoot)
            {
                return _settingsFilePath ??= LoadSettingsFilePath();
            }
        }
    }

    public static WorkbenchSettings GetSettings()
    {
        lock (SyncRoot)
        {
            if (_cachedSettings is not null)
            {
                return Clone(_cachedSettings);
            }

            _cachedSettings = LoadSettingsInternal();
            return Clone(_cachedSettings);
        }
    }

    public static void UpdateSettings(Action<WorkbenchSettings> updateAction)
    {
        lock (SyncRoot)
        {
            WorkbenchSettings settings = _cachedSettings is null
                ? LoadSettingsInternal()
                : Clone(_cachedSettings);

            updateAction(settings);
            EnsureDefaults(settings);
            SaveSettingsInternal(settings);
            _cachedSettings = Clone(settings);
        }
    }

    public static void SetSettingsFilePath(string settingsFilePath)
    {
        lock (SyncRoot)
        {
            string normalizedPath = NormalizeFilePath(settingsFilePath, DefaultSettingsFilePath);
            WorkbenchSettings settings = _cachedSettings ?? LoadSettingsInternal();

            Directory.CreateDirectory(Path.GetDirectoryName(normalizedPath) ?? DefaultStorageRootDirectory);
            Directory.CreateDirectory(BootstrapDirectory);

            File.WriteAllText(
                BootstrapFilePathValue,
                JsonSerializer.Serialize(new SettingsBootstrap { SettingsFilePath = normalizedPath }, GetJsonOptions()));

            _settingsFilePath = normalizedPath;
            SaveSettingsInternal(settings);
            _cachedSettings = Clone(settings);
        }
    }

    private static WorkbenchSettings LoadSettingsInternal()
    {
        string path = _settingsFilePath ??= LoadSettingsFilePath();

        try
        {
            if (File.Exists(path))
            {
                WorkbenchSettings? settings = JsonSerializer.Deserialize<WorkbenchSettings>(File.ReadAllText(path));
                if (settings is not null)
                {
                    EnsureDefaults(settings);
                    return settings;
                }
            }
        }
        catch
        {
        }

        return CreateDefaultSettings();
    }

    private static void SaveSettingsInternal(WorkbenchSettings settings)
    {
        string path = _settingsFilePath ??= LoadSettingsFilePath();
        Directory.CreateDirectory(Path.GetDirectoryName(path) ?? DefaultStorageRootDirectory);
        File.WriteAllText(path, JsonSerializer.Serialize(settings, GetJsonOptions()));
    }

    private static string LoadSettingsFilePath()
    {
        try
        {
            if (File.Exists(BootstrapFilePathValue))
            {
                SettingsBootstrap? bootstrap = JsonSerializer.Deserialize<SettingsBootstrap>(File.ReadAllText(BootstrapFilePathValue));
                if (!string.IsNullOrWhiteSpace(bootstrap?.SettingsFilePath))
                {
                    return NormalizeFilePath(bootstrap.SettingsFilePath, DefaultSettingsFilePath);
                }
            }
        }
        catch
        {
        }

        return DefaultSettingsFilePath;
    }

    private static WorkbenchSettings CreateDefaultSettings()
    {
        return new WorkbenchSettings
        {
            StorageRootDirectory = DefaultStorageRootDirectory,
            Memo = new MemoSetting
            {
                DatabasePath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "JinoWorkHost",
                    "memo.db")
            },
            Bmes = new BmesSetting
            {
                LastImportExportFilePath = Path.Combine(DefaultStorageRootDirectory, "BMESIDINFO.json")
            }
        };
    }

    private static void EnsureDefaults(WorkbenchSettings settings)
    {
        settings.StorageRootDirectory = NormalizeDirectoryPath(settings.StorageRootDirectory);
        settings.CaptureHotkeys ??= new Dictionary<string, CaptureHotkeySetting>();
        settings.VideoConverter ??= new VideoConverterSetting();
        settings.Memo ??= new MemoSetting();
        settings.Bmes ??= new BmesSetting();
        settings.Codex ??= new CodexSetting();
        settings.Claude ??= new ClaudeSetting();
        settings.DataInference ??= new DataInferenceSetting();
        settings.Schedule      ??= new ScheduleSetting();

        if (string.IsNullOrWhiteSpace(settings.Bmes.LastImportExportFilePath))
        {
            settings.Bmes.LastImportExportFilePath = Path.Combine(settings.StorageRootDirectory, "BMESIDINFO.json");
        }
    }

    private static WorkbenchSettings Clone(WorkbenchSettings settings)
    {
        string json = JsonSerializer.Serialize(settings, GetJsonOptions());
        return JsonSerializer.Deserialize<WorkbenchSettings>(json, GetJsonOptions()) ?? CreateDefaultSettings();
    }

    private static JsonSerializerOptions GetJsonOptions()
    {
        return new JsonSerializerOptions { WriteIndented = true };
    }

    private static string NormalizeFilePath(string path, string fallbackPath)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return fallbackPath;
        }

        return Path.GetFullPath(path.Trim());
    }

    private static string NormalizeDirectoryPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return DefaultStorageRootDirectory;
        }

        return Path.GetFullPath(path.Trim());
    }

    public static void SaveCodexCookieHeader(string cookieHeader)
    {
        if (string.IsNullOrWhiteSpace(cookieHeader))
        {
            return;
        }

        UpdateSettings(settings =>
        {
            settings.Codex.AutoLoginEnabled = true;
            settings.Codex.EncryptedCookieHeader = Protect(cookieHeader.Trim());
        });
    }

    public static bool IsCodexAutoLoginEnabled()
    {
        try
        {
            return GetSettings().Codex.AutoLoginEnabled;
        }
        catch
        {
            return true;
        }
    }

    public static void SetCodexAutoLoginEnabled(bool enabled)
    {
        UpdateSettings(settings => settings.Codex.AutoLoginEnabled = enabled);
    }

    public static string? TryGetCodexCookieHeader()
    {
        try
        {
            WorkbenchSettings settings = GetSettings();
            if (!settings.Codex.AutoLoginEnabled || string.IsNullOrWhiteSpace(settings.Codex.EncryptedCookieHeader))
            {
                return null;
            }

            return Unprotect(settings.Codex.EncryptedCookieHeader);
        }
        catch
        {
            return null;
        }
    }

    public static void ClearCodexCookieHeader()
    {
        UpdateSettings(settings =>
        {
            settings.Codex.EncryptedCookieHeader = string.Empty;
        });
    }

    public static void SaveAnthropicSessionKey(string key)
    {
        UpdateSettings(settings =>
        {
            settings.DataInference ??= new DataInferenceSetting();
            settings.DataInference.AnthropicSessionKey = key.Trim();
        });
    }

    public static string GetAnthropicSessionKey()
    {
        try { return GetSettings().DataInference?.AnthropicSessionKey ?? string.Empty; }
        catch { return string.Empty; }
    }

    public static void SaveDataInferenceDatabasePath(string path)
    {
        UpdateSettings(settings =>
        {
            settings.DataInference ??= new DataInferenceSetting();
            settings.DataInference.DatabasePath = path;
        });
    }

    public static string GetDataInferenceDatabasePath()
    {
        try { return GetSettings().DataInference?.DatabasePath ?? string.Empty; }
        catch { return string.Empty; }
    }

    public static void SaveScheduleDatabasePath(string path)
    {
        UpdateSettings(settings =>
        {
            settings.Schedule ??= new ScheduleSetting();
            settings.Schedule.DatabasePath = path;
        });
    }

    public static string GetScheduleDatabasePath()
    {
        try { return GetSettings().Schedule?.DatabasePath ?? string.Empty; }
        catch { return string.Empty; }
    }

    public static void SaveClaudeApiKey(string apiKey)
    {
        UpdateSettings(settings =>
        {
            settings.Claude.EncryptedApiKey = string.IsNullOrWhiteSpace(apiKey)
                ? string.Empty
                : Protect(apiKey.Trim());
        });
    }

    public static string? TryGetClaudeApiKey()
    {
        try
        {
            WorkbenchSettings settings = GetSettings();
            if (string.IsNullOrWhiteSpace(settings.Claude.EncryptedApiKey))
            {
                return null;
            }

            return Unprotect(settings.Claude.EncryptedApiKey);
        }
        catch
        {
            return null;
        }
    }

    public static void SaveOpenAiApiKey(string apiKey)
    {
        UpdateSettings(settings =>
        {
            settings.OpenAi.EncryptedApiKey = string.IsNullOrWhiteSpace(apiKey)
                ? string.Empty
                : Protect(apiKey.Trim());
        });
    }

    public static string? TryGetOpenAiApiKey()
    {
        try
        {
            WorkbenchSettings settings = GetSettings();
            if (string.IsNullOrWhiteSpace(settings.OpenAi.EncryptedApiKey))
            {
                return null;
            }

            return Unprotect(settings.OpenAi.EncryptedApiKey);
        }
        catch
        {
            return null;
        }
    }

    private static string Protect(string plainText)
    {
        byte[] plaintextBytes = Encoding.UTF8.GetBytes(plainText);
        byte[] protectedBytes = ProtectedData.Protect(plaintextBytes, null, DataProtectionScope.CurrentUser);
        return Convert.ToBase64String(protectedBytes);
    }

    private static string Unprotect(string encryptedText)
    {
        byte[] protectedBytes = Convert.FromBase64String(encryptedText);
        byte[] plaintextBytes = ProtectedData.Unprotect(protectedBytes, null, DataProtectionScope.CurrentUser);
        return Encoding.UTF8.GetString(plaintextBytes);
    }
}
