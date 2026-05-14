param(
    [string]$Version = "",
    [string]$BaseUrl = "http://10.6.4.54:5050",
    [string]$Notes = ""
)

$ErrorActionPreference = 'Stop'

$standaloneRoot = Split-Path -Parent $PSScriptRoot
$repoRoot = Split-Path -Parent $standaloneRoot
$webUpdateDir = Join-Path $repoRoot 'JinoSupporter.Web\standalone-updates'
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

dotnet publish $projectPath -c Release -r win-x64 --self-contained false -o $publishDir `
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

$manifest = [ordered]@{
    version = $Version
    url = "$BaseUrl/standalone/download/$zipName"
    sha256 = $sha256
    notes = $Notes
    publishedAt = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
}

$manifestPath = Join-Path $webUpdateDir 'update.json'
$manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

Write-Host "Published standalone update manifest:"
Write-Host $manifestPath
Write-Host "Package:"
Write-Host $zipPath
