using System.Diagnostics;
using System.IO.Compression;
using System.Net.Http;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using System.Windows;

namespace BmesNgRateStandalone;

public sealed class StandaloneUpdateService
{
    private static readonly Uri ManifestUri = new("http://10.6.4.54:5050/standalone/update.json");
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    public async Task CheckAndPromptAsync(Window owner, bool notifyWhenCurrent = false)
    {
        try
        {
            UpdateManifest? manifest = await FetchManifestAsync();
            if (manifest is null)
            {
                if (notifyWhenCurrent)
                {
                    MessageBox.Show(
                        owner,
                        "No standalone update manifest was found on the server.",
                        "Standalone Update",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }

                return;
            }

            if (!IsNewer(manifest.Version))
            {
                if (notifyWhenCurrent)
                {
                    string currentVersion = GetCurrentVersion().ToString();
                    MessageBox.Show(
                        owner,
                        $"Already up to date.\n\nCurrent: {currentVersion}\nServer: {manifest.Version}",
                        "Standalone Update",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }

                return;
            }

            string notes = string.IsNullOrWhiteSpace(manifest.Notes) ? "" : "\n\n" + manifest.Notes.Trim();
            MessageBoxResult result = MessageBox.Show(
                owner,
                $"BMES NG Rate Standalone {manifest.Version} update is available.{notes}\n\nUpdate now?",
                "Standalone Update",
                MessageBoxButton.YesNo,
                MessageBoxImage.Information);

            if (result != MessageBoxResult.Yes) return;

            string packagePath = await DownloadAndValidateAsync(manifest);
            StartUpdaterAndExit(packagePath);
        }
        catch (Exception ex)
        {
            if (!notifyWhenCurrent) return;

            MessageBox.Show(
                owner,
                "Update check failed. The current version will continue to run.\n\n" + ex.Message,
                "Standalone Update",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private static async Task<UpdateManifest?> FetchManifestAsync()
    {
        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(8) };
        using var response = await http.GetAsync(ManifestUri);
        if (!response.IsSuccessStatusCode) return null;

        await using Stream stream = await response.Content.ReadAsStreamAsync();
        return await JsonSerializer.DeserializeAsync<UpdateManifest>(stream, JsonOptions);
    }

    private static async Task<string> DownloadAndValidateAsync(UpdateManifest manifest)
    {
        if (string.IsNullOrWhiteSpace(manifest.Url))
            throw new InvalidOperationException("Update manifest does not contain a package URL.");

        Uri packageUri = Uri.TryCreate(manifest.Url, UriKind.Absolute, out Uri? absolute)
            ? absolute
            : new Uri(ManifestUri, manifest.Url);

        string updateDir = Path.Combine(Path.GetTempPath(), "BmesNgRateStandalone", "updates");
        Directory.CreateDirectory(updateDir);
        string packagePath = Path.Combine(updateDir, "BmesNgRateStandalone-" + manifest.Version + ".zip");

        using var http = new HttpClient { Timeout = TimeSpan.FromMinutes(10) };
        await using (Stream input = await http.GetStreamAsync(packageUri))
        await using (FileStream output = File.Create(packagePath))
        {
            await input.CopyToAsync(output);
        }

        if (!string.IsNullOrWhiteSpace(manifest.Sha256))
        {
            string actual = await ComputeSha256Async(packagePath);
            if (!actual.Equals(manifest.Sha256.Trim(), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Downloaded update package hash does not match the manifest.");
        }

        return packagePath;
    }

    private static void StartUpdaterAndExit(string packagePath)
    {
        string appDir = AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        string exePath = Environment.ProcessPath
            ?? Path.Combine(appDir, "BmesNgRateStandalone.exe");
        string stagingDir = Path.Combine(Path.GetTempPath(), "BmesNgRateStandalone", "staging-" + Guid.NewGuid().ToString("N"));
        string updaterPath = Path.Combine(Path.GetTempPath(), "BmesNgRateStandalone", "ApplyStandaloneUpdate.ps1");

        Directory.CreateDirectory(Path.GetDirectoryName(updaterPath)!);
        File.WriteAllText(updaterPath, BuildUpdaterScript());

        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            UseShellExecute = true,
            WindowStyle = ProcessWindowStyle.Hidden,
        };
        psi.ArgumentList.Add("-NoProfile");
        psi.ArgumentList.Add("-ExecutionPolicy");
        psi.ArgumentList.Add("Bypass");
        psi.ArgumentList.Add("-File");
        psi.ArgumentList.Add(updaterPath);
        psi.ArgumentList.Add("-ProcessId");
        psi.ArgumentList.Add(Environment.ProcessId.ToString());
        psi.ArgumentList.Add("-PackagePath");
        psi.ArgumentList.Add(packagePath);
        psi.ArgumentList.Add("-TargetDir");
        psi.ArgumentList.Add(appDir);
        psi.ArgumentList.Add("-ExePath");
        psi.ArgumentList.Add(exePath);
        psi.ArgumentList.Add("-StagingDir");
        psi.ArgumentList.Add(stagingDir);

        Process.Start(psi);
        Application.Current.Shutdown();
    }

    private static string BuildUpdaterScript() =>
        """
        param(
            [Parameter(Mandatory=$true)][int]$ProcessId,
            [Parameter(Mandatory=$true)][string]$PackagePath,
            [Parameter(Mandatory=$true)][string]$TargetDir,
            [Parameter(Mandatory=$true)][string]$ExePath,
            [Parameter(Mandatory=$true)][string]$StagingDir
        )
        $ErrorActionPreference = 'Stop'
        Wait-Process -Id $ProcessId -ErrorAction SilentlyContinue
        if (Test-Path -LiteralPath $StagingDir) {
            Remove-Item -LiteralPath $StagingDir -Recurse -Force
        }
        New-Item -ItemType Directory -Path $StagingDir | Out-Null
        Expand-Archive -LiteralPath $PackagePath -DestinationPath $StagingDir -Force
        Copy-Item -Path (Join-Path $StagingDir '*') -Destination $TargetDir -Recurse -Force
        Start-Process -FilePath $ExePath -WorkingDirectory $TargetDir
        """;

    private static bool IsNewer(string? manifestVersion)
    {
        if (string.IsNullOrWhiteSpace(manifestVersion)) return false;
        Version current = GetCurrentVersion();
        return Version.TryParse(NormalizeVersion(manifestVersion), out Version? available)
               && available > current;
    }

    private static Version GetCurrentVersion() =>
        Assembly.GetExecutingAssembly().GetName().Version ?? new Version(0, 0, 0, 0);

    private static string NormalizeVersion(string version)
    {
        string cleaned = version.Trim().TrimStart('v', 'V');
        int dash = cleaned.IndexOfAny(['-', '+']);
        return dash >= 0 ? cleaned[..dash] : cleaned;
    }

    private static async Task<string> ComputeSha256Async(string path)
    {
        await using FileStream stream = File.OpenRead(path);
        byte[] hash = await SHA256.HashDataAsync(stream);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    private sealed class UpdateManifest
    {
        public string Version { get; set; } = "";
        public string Url { get; set; } = "";
        public string Sha256 { get; set; } = "";
        public string Notes { get; set; } = "";
    }
}
