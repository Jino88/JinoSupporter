namespace JinoSupporter.Web.Services;

public static class AiPromptRegistry
{
    public const string PromptFolderName = "AI_PROMPTS";

    public static string Read(string relativePath)
    {
        string path = ResolvePath(relativePath);
        if (!File.Exists(path))
            throw new FileNotFoundException($"AI prompt template not found: {path}", path);

        return NormalizeNewlines(File.ReadAllText(path)).TrimEnd();
    }

    public static string Render(string relativePath, params (string Key, string? Value)[] values)
    {
        string text = Read(relativePath);
        foreach ((string key, string? value) in values)
        {
            text = text.Replace("{{" + key + "}}", value ?? string.Empty, StringComparison.Ordinal);
        }

        return NormalizeNewlines(text).Replace("\n", Environment.NewLine);
    }

    public static string TextOrNone(string? value)
        => string.IsNullOrWhiteSpace(value) ? "(none)" : value.Trim();

    public static string FindRepositoryRoot()
    {
        foreach (string start in StartDirectories())
        {
            string? root = FindRepositoryRootFrom(start, requireSourceMarker: true);
            if (!string.IsNullOrWhiteSpace(root)) return root;
        }

        foreach (string start in StartDirectories())
        {
            string? root = FindRepositoryRootFrom(start, requireSourceMarker: false);
            if (!string.IsNullOrWhiteSpace(root)) return root;
        }

        return Directory.GetCurrentDirectory();
    }

    public static string ResolvePath(string relativePath)
    {
        string normalized = (relativePath ?? string.Empty)
            .Replace('\\', Path.DirectorySeparatorChar)
            .Replace('/', Path.DirectorySeparatorChar)
            .TrimStart(Path.DirectorySeparatorChar);

        return Path.Combine(FindPromptRoot(), normalized);
    }

    private static string FindPromptRoot()
    {
        string repoRoot = FindRepositoryRoot();
        string promptRoot = Path.Combine(repoRoot, PromptFolderName);
        if (Directory.Exists(promptRoot)) return promptRoot;

        return Path.Combine(Directory.GetCurrentDirectory(), PromptFolderName);
    }

    private static string? FindRepositoryRootFrom(string start, bool requireSourceMarker)
    {
        string dir = Path.GetFullPath(start);
        for (int i = 0; i < 12; i++)
        {
            bool hasPrompts = Directory.Exists(Path.Combine(dir, PromptFolderName));
            if (hasPrompts && (!requireSourceMarker || HasSourceMarker(dir)))
                return dir;

            string? parent = Path.GetDirectoryName(dir.TrimEnd('\\', '/'));
            if (string.IsNullOrWhiteSpace(parent) || string.Equals(parent, dir, StringComparison.OrdinalIgnoreCase))
                break;
            dir = parent;
        }

        return null;
    }

    private static bool HasSourceMarker(string dir)
        => File.Exists(Path.Combine(dir, "JinoSupporter.sln"))
           || Directory.Exists(Path.Combine(dir, ".git"))
           || File.Exists(Path.Combine(dir, "JinoSupporter.Web", "JinoSupporter.Web.csproj"));

    private static IEnumerable<string> StartDirectories()
    {
        yield return AppContext.BaseDirectory;
        yield return Directory.GetCurrentDirectory();
    }

    private static string NormalizeNewlines(string text)
        => (text ?? string.Empty)
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace("\r", "\n", StringComparison.Ordinal);
}
