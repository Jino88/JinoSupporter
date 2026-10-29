namespace PortalGNLAB.Web.Services;

public static class AppStoragePaths
{
    private const string DataDirEnvironmentVariable = "PORTALGNLAB_DATA_DIR";
    // Name used before the app was renamed from JinoSupporter; still honoured so existing setups keep working.
    private const string LegacyDataDirEnvironmentVariable = "JINO" + "SUPPORTER_DATA_DIR";
    private const string LegacyAppFolderName = "Jino" + "Supporter";
    private const string DefaultRootDirectory = @"D:\000. MyWorks\002. DB";

    public static string RootDirectory
    {
        get
        {
            string? configured = Environment.GetEnvironmentVariable(DataDirEnvironmentVariable);
            if (string.IsNullOrWhiteSpace(configured))
                configured = Environment.GetEnvironmentVariable(LegacyDataDirEnvironmentVariable);
            string root = string.IsNullOrWhiteSpace(configured)
                ? DefaultRootDirectory
                : configured.Trim();
            return Path.GetFullPath(root);
        }
    }

    /// <summary>Folder under &lt;root&gt;\PortalGNLAB\&lt;name&gt;. If only the folder from before the rename
    /// (&lt;root&gt;\JinoSupporter\&lt;name&gt;) exists, it is moved over so saved logins are kept.</summary>
    public static string AppFolder(string name)
    {
        string path = Combine("PortalGNLAB", name);
        string legacy = Combine(LegacyAppFolderName, name);
        if (!Directory.Exists(path) && Directory.Exists(legacy))
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            Directory.Move(legacy, path);
        }
        return path;
    }

    public static string Combine(params string[] parts)
    {
        string path = RootDirectory;
        foreach (string part in parts)
            path = Path.Combine(path, part);
        return path;
    }
}
