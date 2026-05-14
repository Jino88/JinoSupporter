param(
    [string]$Version = "",
    [string]$IsccPath = ""
)

$ErrorActionPreference = 'Stop'

$standaloneRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $standaloneRoot 'BmesNgRateStandalone.csproj'
$installerPath = Join-Path $standaloneRoot 'installer.iss'
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

if ([string]::IsNullOrWhiteSpace($IsccPath)) {
    $command = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if ($command) {
        $IsccPath = $command.Source
    }
}

if ([string]::IsNullOrWhiteSpace($IsccPath)) {
    $candidatePaths = @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles}\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles(x86)}\Inno Setup 5\ISCC.exe",
        "${env:ProgramFiles}\Inno Setup 5\ISCC.exe"
    )
    foreach ($candidate in $candidatePaths) {
        if (Test-Path -LiteralPath $candidate) {
            $IsccPath = $candidate
            break
        }
    }
}

if ([string]::IsNullOrWhiteSpace($IsccPath) -or -not (Test-Path -LiteralPath $IsccPath)) {
    throw 'Inno Setup compiler was not found. Install Inno Setup 6 or pass -IsccPath "C:\Path\To\ISCC.exe".'
}

Push-Location $standaloneRoot
try {
    & $IsccPath "/DMyAppVersion=$Version" $installerPath
}
finally {
    Pop-Location
}

$setupPath = Join-Path $standaloneRoot 'dist\BmesNgRateStandalone_Setup.exe'
if (-not (Test-Path -LiteralPath $setupPath)) {
    throw "Installer build finished but setup file was not found: $setupPath"
}

Write-Host "Installer:"
Write-Host $setupPath
