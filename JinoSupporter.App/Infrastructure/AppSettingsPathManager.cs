using System.IO;

namespace WorkbenchHost.Infrastructure;

public static class AppSettingsPathManager
{
    public static string BootstrapFilePath => WorkbenchSettingsStore.BootstrapFilePath;

    public static string SettingsFilePath => WorkbenchSettingsStore.SettingsFilePath;

    public static string DefaultStorageRootDirectory => WorkbenchSettingsStore.DefaultStorageRootDirectory;

    public static string StorageRootDirectory =>
        WorkbenchSettingsStore.GetSettings().StorageRootDirectory;

    public static string GetModuleDirectory(string moduleName)
    {
        return Path.Combine(StorageRootDirectory, moduleName);
    }

    public static string GetModuleFilePath(string moduleName, string fileName)
    {
        return Path.Combine(GetModuleDirectory(moduleName), fileName);
    }

    public static void SetStorageRootDirectory(string directoryPath)
    {
        WorkbenchSettingsStore.UpdateSettings(settings => settings.StorageRootDirectory = directoryPath);
    }
}
