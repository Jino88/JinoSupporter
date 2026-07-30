param(
    [string]$Version = "",
    [string]$BaseUrl = "http://10.6.4.54:5050",
    [string]$Notes = ""
)

$ErrorActionPreference = 'Stop'

$standaloneRoot = Split-Path -Parent $PSScriptRoot
$repoRoot = Split-Path -Parent $standaloneRoot
$webUpdateDir = Join-Path $repoRoot 'JinoSupporter.Web\standalone-updates'
$webDebugUpdateDir = Join-Path $repoRoot 'JinoSupporter.Web\bin\Debug\net8.0\standalone-updates'
$projectPath = Join-Path $standaloneRoot 'BmesNgRateStandalone.csproj'
$publishDir = Join-Path $standaloneRoot 'bin\Release\net8.0-windows\win-x64\publish'

if ([string]::IsNullOrWhiteSpace($Version)) {
    [xml]$proj = Get-Content -LiteralPath $projectPath
    $Version = $proj.Project.PropertyGroup.Version
}
if ([string]::IsNullOrWhiteSpace($Version)) {
    throw 'Version was not provided and could not be read from the project file.'
}

$numericVersion = (($Version.Trim().TrimStart('v', 'V') -split '[-+]') | Select-Object -First 1)
if (-not [version]::TryParse($numericVersion, [ref]([version]$null))) {
    throw "Version '$Version' must start with a numeric version such as 1.0.1."
}

$versionParts = @($numericVersion.Split('.'))
while ($versionParts.Count -lt 4) {
    $versionParts += '0'
}
$assemblyVersion = ($versionParts[0..3] -join '.')

dotnet publish $projectPath -c Release -r win-x64 --self-contained true -o $publishDir `
    "-p:Version=$Version" `
    "-p:AssemblyVersion=$assemblyVersion" `
    "-p:FileVersion=$assemblyVersion"

New-Item -ItemType Directory -Path $webUpdateDir -Force | Out-Null
$zipName = "BmesNgRateStandalone-$Version.zip"
$zipPath = Join-Path $webUpdateDir $zipName
if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}

Compress-Archive -Path (Join-Path $publishDir '*') -DestinationPath $zipPath -Force
$sha256 = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()

# Installer for a fresh PC (the zip above is what the in-app updater pulls). Delegated to
# BuildStandaloneInstaller.ps1, which owns the ISCC lookup; -SkipPublish reuses the publish
# output produced above instead of building it a second time.
$setupName = $null
$setupSha256 = $null
try {
    & (Join-Path $PSScriptRoot 'BuildStandaloneInstaller.ps1') -Version $Version -SkipPublish

    $setupName = "BmesNgRateStandalone_Setup-$Version.exe"
    $setupSource = Join-Path $standaloneRoot (Join-Path 'dist' $setupName)
    if (-not (Test-Path -LiteralPath $setupSource)) {
        throw "Installer build finished but $setupSource was not found."
    }

    Copy-Item -LiteralPath $setupSource -Destination (Join-Path $webUpdateDir $setupName) -Force
    $setupSha256 = (Get-FileHash -LiteralPath $setupSource -Algorithm SHA256).Hash.ToLowerInvariant()
}
catch {
    # Zip-only publishing still works (that is what the in-app updater uses); the PC Download
    # page just falls back to offering the portable zip.
    $setupName = $null
    Write-Warning "Installer was not built — publishing the zip only. $($_.Exception.Message)"
}

$manifest = [ordered]@{
    version = $Version
    url = "$BaseUrl/standalone/download/$zipName"
    sha256 = $sha256
    notes = $Notes
    publishedAt = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
}

if ($setupName) {
    # Extra fields; the standalone updater ignores unknown properties and keeps using `url`.
    $manifest.setupUrl = "$BaseUrl/standalone/download/$setupName"
    $manifest.setupSha256 = $setupSha256
}

$manifestPath = Join-Path $webUpdateDir 'update.json'
$manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

New-Item -ItemType Directory -Path $webDebugUpdateDir -Force | Out-Null
Copy-Item -LiteralPath $zipPath -Destination (Join-Path $webDebugUpdateDir $zipName) -Force
Copy-Item -LiteralPath $manifestPath -Destination (Join-Path $webDebugUpdateDir 'update.json') -Force
if ($setupName) {
    Copy-Item -LiteralPath (Join-Path $webUpdateDir $setupName) -Destination (Join-Path $webDebugUpdateDir $setupName) -Force
}

Write-Host "Published standalone update manifest:"
Write-Host $manifestPath
Write-Host "Package:"
Write-Host $zipPath
if ($setupName) {
    Write-Host "Installer:"
    Write-Host (Join-Path $webUpdateDir $setupName)
}
Write-Host "Runtime update directory:"
Write-Host $webDebugUpdateDir
