using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
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
}
