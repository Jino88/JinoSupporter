using System.Text.Json;

namespace JinoSupporter.Web.Services;

/// <summary>
/// What the PC Download page offers: the standalone desktop build that
/// <c>BmesNgRateStandalone/tools/PublishStandaloneUpdate.ps1</c> dropped into
/// <c>standalone-updates/</c>. <see cref="Setup"/> is the Inno Setup installer (preferred
/// for a fresh PC); <see cref="Package"/> is the self-contained zip that the in-app
/// updater already pulls. Either may be missing — the installer only exists when the
/// publish machine had Inno Setup.
/// </summary>
public sealed record StandaloneRelease(
    string Version,
    string Notes,
    DateTimeOffset? PublishedAt,
    StandaloneAsset? Setup,
    StandaloneAsset? Package);

/// <summary>One downloadable file, sized for display and linked through
/// <c>/standalone/download/{fileName}</c>.</summary>
public sealed record StandaloneAsset(string FileName, long SizeBytes)
{
    public string Url => "/standalone/download/" + Uri.EscapeDataString(FileName);

    public string SizeText => SizeBytes >= 1024L * 1024 * 1024
        ? (SizeBytes / (1024d * 1024 * 1024)).ToString("N2") + " GB"
        : SizeBytes >= 1024 * 1024
            ? (SizeBytes / (1024d * 1024)).ToString("N1") + " MB"
            : (SizeBytes / 1024d).ToString("N0") + " KB";
}

public static class StandaloneDownloadCatalog
{
    public const string DirectoryName = "standalone-updates";

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    /// <summary>Reads the published manifest and resolves the files it names.
    /// Returns null when nothing has been published yet.</summary>
    public static StandaloneRelease? Read(string contentRootPath)
    {
        string dir = Path.Combine(contentRootPath, DirectoryName);
        string manifestPath = Path.Combine(dir, "update.json");
        if (!File.Exists(manifestPath)) return null;

        Manifest? m;
        try
        {
            m = JsonSerializer.Deserialize<Manifest>(File.ReadAllText(manifestPath), JsonOpts);
        }
        catch
        {
            return null;
        }
        if (m is null || string.IsNullOrWhiteSpace(m.Version)) return null;

        string version = m.Version.Trim();

        // The manifest names the zip by URL; the installer is published beside it under a
        // fixed naming convention, so fall back to that when the field is absent.
        StandaloneAsset? setup = Asset(dir, FileNameOf(m.SetupUrl) is { Length: > 0 } s
            ? s
            : $"BmesNgRateStandalone_Setup-{version}.exe");
        StandaloneAsset? package = Asset(dir, FileNameOf(m.Url) is { Length: > 0 } p
            ? p
            : $"BmesNgRateStandalone-{version}.zip");

        DateTimeOffset? published = DateTimeOffset.TryParse(m.PublishedAt, out DateTimeOffset dt)
            ? dt
            : null;

        return new StandaloneRelease(version, (m.Notes ?? "").Trim(), published, setup, package);
    }

    private static StandaloneAsset? Asset(string dir, string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName)) return null;
        string path = Path.Combine(dir, Path.GetFileName(fileName));
        if (!File.Exists(path)) return null;
        return new StandaloneAsset(Path.GetFileName(path), new FileInfo(path).Length);
    }

    /// <summary>The last path segment of an absolute or relative URL.</summary>
    private static string FileNameOf(string? url)
    {
        if (string.IsNullOrWhiteSpace(url)) return "";
        string trimmed = url.Trim().TrimEnd('/');
        int slash = trimmed.LastIndexOf('/');
        return slash >= 0 ? trimmed[(slash + 1)..] : trimmed;
    }

    private sealed class Manifest
    {
        public string Version { get; set; } = "";
        public string Url { get; set; } = "";
        public string SetupUrl { get; set; } = "";
        public string Notes { get; set; } = "";
        public string PublishedAt { get; set; } = "";
    }
}
