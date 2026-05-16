using System.Diagnostics;
using System.IO.Compression;
using System.Net.Http;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
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
        bool updateRequested = false;

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

            updateRequested = true;
            string packagePath = await DownloadAndValidateAsync(manifest);
            StartUpdaterAndExit(packagePath);
        }
        catch (Exception ex)
        {
            if (!notifyWhenCurrent && !updateRequested) return;

            MessageBox.Show(
                owner,
                updateRequested
                    ? "Update could not be started. The current version will continue to run.\n\n" + ex.Message
                    : "Update check failed. The current version will continue to run.\n\n" + ex.Message,
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
        string updateRoot = Path.Combine(Path.GetTempPath(), "BmesNgRateStandalone");
        string updaterPath = Path.Combine(updateRoot, "ApplyStandaloneUpdate.ps1");
        string logPath = Path.Combine(updateRoot, "update-apply.log");

        Directory.CreateDirectory(Path.GetDirectoryName(updaterPath)!);
        File.WriteAllText(updaterPath, BuildUpdaterScript());

        bool needsElevation = !CanWriteToDirectory(appDir);
        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            UseShellExecute = true,
            WindowStyle = ProcessWindowStyle.Hidden,
            WorkingDirectory = appDir,
            Arguments = BuildUpdaterArguments(
                updaterPath,
                Environment.ProcessId,
                packagePath,
                appDir,
                exePath,
                stagingDir,
                logPath),
        };

        if (needsElevation)
        {
            psi.Verb = "runas";
        }

        if (Process.Start(psi) is null)
            throw new InvalidOperationException("Windows did not start the update process.");

        Application.Current.Shutdown();
    }

    private static string BuildUpdaterArguments(
        string updaterPath,
        int processId,
        string packagePath,
        string targetDir,
        string exePath,
        string stagingDir,
        string logPath)
    {
        string[] args =
        [
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            updaterPath,
            "-ProcessId",
            processId.ToString(),
            "-PackagePath",
            packagePath,
            "-TargetDir",
            targetDir,
            "-ExePath",
            exePath,
            "-StagingDir",
            stagingDir,
            "-LogPath",
            logPath,
        ];

        return string.Join(" ", args.Select(QuoteProcessArgument));
    }

    private static string QuoteProcessArgument(string value)
    {
        if (value.Length == 0) return "\"\"";
        if (!value.Any(c => char.IsWhiteSpace(c) || c == '"')) return value;

        var quoted = new StringBuilder();
        quoted.Append('"');
        int backslashCount = 0;

        foreach (char c in value)
        {
            if (c == '\\')
            {
                backslashCount++;
                continue;
            }

            if (c == '"')
            {
                quoted.Append('\\', backslashCount * 2 + 1);
                quoted.Append('"');
                backslashCount = 0;
                continue;
            }

            quoted.Append('\\', backslashCount);
            quoted.Append(c);
            backslashCount = 0;
        }

        quoted.Append('\\', backslashCount * 2);
        quoted.Append('"');
        return quoted.ToString();
    }

    private static bool CanWriteToDirectory(string directory)
    {
        try
        {
            Directory.CreateDirectory(directory);
            string probePath = Path.Combine(directory, ".update-write-test-" + Guid.NewGuid().ToString("N") + ".tmp");
            File.WriteAllText(probePath, "");
            File.Delete(probePath);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string BuildUpdaterScript() =>
        """
        param(
            [Parameter(Mandatory=$true)][int]$ProcessId,
            [Parameter(Mandatory=$true)][string]$PackagePath,
            [Parameter(Mandatory=$true)][string]$TargetDir,
            [Parameter(Mandatory=$true)][string]$ExePath,
            [Parameter(Mandatory=$true)][string]$StagingDir,
            [Parameter(Mandatory=$true)][string]$LogPath
        )
        $ErrorActionPreference = 'Stop'

        function Write-UpdateLog([string]$Message) {
            $logDir = Split-Path -Parent $LogPath
            if (-not [string]::IsNullOrWhiteSpace($logDir)) {
                New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            }
            $line = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Message
            Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
        }

        function Show-UpdateFailure([string]$Message) {
            try {
                Add-Type -AssemblyName PresentationFramework
                [System.Windows.MessageBox]::Show(
                    "Update failed. The current version was reopened if possible.`n`n$Message`n`nLog: $LogPath",
                    "Standalone Update",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                ) | Out-Null
            }
            catch {
            }
        }

        function Get-PackageRoot([string]$Root, [string]$ExeName) {
            if (Test-Path -LiteralPath (Join-Path $Root $ExeName)) {
                return $Root
            }

            $children = @(Get-ChildItem -LiteralPath $Root -Force)
            if ($children.Count -eq 1 -and $children[0].PSIsContainer) {
                $singleRoot = $children[0].FullName
                if (Test-Path -LiteralPath (Join-Path $singleRoot $ExeName)) {
                    return $singleRoot
                }
            }

            $foundExe = Get-ChildItem -LiteralPath $Root -Filter $ExeName -Recurse -File -Force | Select-Object -First 1
            if ($foundExe) {
                return $foundExe.DirectoryName
            }

            throw "The update package does not contain $ExeName."
        }

        try {
            Write-UpdateLog "Starting standalone update."
            Write-UpdateLog "PackagePath=$PackagePath"
            Write-UpdateLog "TargetDir=$TargetDir"
            Write-UpdateLog "ExePath=$ExePath"
            Write-UpdateLog "StagingDir=$StagingDir"

            Wait-Process -Id $ProcessId -ErrorAction SilentlyContinue

            if (-not (Test-Path -LiteralPath $PackagePath)) {
                throw "Downloaded package was not found: $PackagePath"
            }

            if (Test-Path -LiteralPath $StagingDir) {
                Remove-Item -LiteralPath $StagingDir -Recurse -Force
            }
            New-Item -ItemType Directory -Path $StagingDir -Force | Out-Null
            New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null

            Expand-Archive -LiteralPath $PackagePath -DestinationPath $StagingDir -Force
            $sourceRoot = Get-PackageRoot -Root $StagingDir -ExeName (Split-Path -Leaf $ExePath)
            Write-UpdateLog "SourceRoot=$sourceRoot"

            $sourceItems = @(Get-ChildItem -LiteralPath $sourceRoot -Force)
            if ($sourceItems.Count -eq 0) {
                throw "The update package is empty."
            }

            foreach ($item in $sourceItems) {
                Copy-Item -LiteralPath $item.FullName -Destination $TargetDir -Recurse -Force
            }

            Write-UpdateLog "Files copied. Restarting application."
            Start-Process -FilePath $ExePath -WorkingDirectory $TargetDir
            Write-UpdateLog "Update finished."
            exit 0
        }
        catch {
            $message = $_.Exception.Message
            try {
                Write-UpdateLog "Update failed: $message"
                Write-UpdateLog $_.ScriptStackTrace
            }
            catch {
            }

            try {
                if (Test-Path -LiteralPath $ExePath) {
                    Start-Process -FilePath $ExePath -WorkingDirectory $TargetDir
                }
            }
            catch {
            }

            Show-UpdateFailure $message
            exit 1
        }
        finally {
            try {
                if (Test-Path -LiteralPath $StagingDir) {
                    Remove-Item -LiteralPath $StagingDir -Recurse -Force
                }
            }
            catch {
            }
        }
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
