using System.IO;

namespace CustomKeyboardCSharp.Services;

internal static class CustomKeyboardPathResolver
{
    private const string AppDataOverrideEnvironmentVariable = "CUSTOMKEYBOARD_APPDATA_DIR";
    private const string DefaultAppDirectoryName = "CustomKeyboardCSharp";

    public static string GetAppDataDirectory()
    {
        string? overrideDirectory = Environment.GetEnvironmentVariable(AppDataOverrideEnvironmentVariable);
        string baseDirectory = string.IsNullOrWhiteSpace(overrideDirectory)
            ? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                DefaultAppDirectoryName)
            : overrideDirectory;

        Directory.CreateDirectory(baseDirectory);
        return baseDirectory;
    }

    public static string GetAppDataPath(params string[] relativeSegments)
    {
        string currentPath = GetAppDataDirectory();
        foreach (string segment in relativeSegments)
        {
            currentPath = Path.Combine(currentPath, segment);
        }

        return currentPath;
    }
}
