namespace BmesNgRateStandalone.Services;

public static class AppStoragePaths
{
    private const string DataDirEnvironmentVariable = "JINOSUPPORTER_DATA_DIR";
    private const string DefaultRootDirectory = @"D:\000. MyWorks\002. DB";

    public static string RootDirectory
    {
        get
        {
            string? configured = Environment.GetEnvironmentVariable(DataDirEnvironmentVariable);
            string root = string.IsNullOrWhiteSpace(configured)
                ? DefaultRootDirectory
                : configured.Trim();
            return Path.GetFullPath(root);
        }
    }

    public static string Combine(params string[] parts)
    {
        string path = RootDirectory;
        foreach (string part in parts)
            path = Path.Combine(path, part);
        return path;
    }
}
