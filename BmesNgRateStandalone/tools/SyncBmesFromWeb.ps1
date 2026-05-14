$ErrorActionPreference = 'Stop'

$standaloneRoot = Split-Path -Parent $PSScriptRoot
$repoRoot = Split-Path -Parent $standaloneRoot
$webRoot = Join-Path $repoRoot 'JinoSupporter.Web'

if (-not (Test-Path -LiteralPath $webRoot)) {
    throw "JinoSupporter.Web not found at $webRoot"
}

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)

function Convert-Text([string] $text) {
    $text = $text -replace 'namespace JinoSupporter\.Web\.Services;', 'namespace BmesNgRateStandalone.Services;'
    $text = $text -replace 'using JinoSupporter\.Web\.Services;', 'using BmesNgRateStandalone.Services;'
    $text = $text -replace '@using\s+JinoSupporter\.Web\.Services', '@using BmesNgRateStandalone.Services'
    $text = $text -replace '@rendermode\s+InteractiveServer\r?\n', ''
    $text = $text -replace 'JinoSupporter\.Web process restarts', 'the standalone app restarts'
    return $text
}

function Copy-Converted([string] $sourceRelative, [string] $targetRelative) {
    $source = Join-Path $webRoot $sourceRelative
    $target = Join-Path $standaloneRoot $targetRelative
    if (-not (Test-Path -LiteralPath $source)) {
        throw "Source file not found: $source"
    }

    $targetDir = Split-Path -Parent $target
    if (-not (Test-Path -LiteralPath $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir | Out-Null
    }

    $text = [System.IO.File]::ReadAllText($source, [System.Text.Encoding]::UTF8)
    $text = Convert-Text $text
    [System.IO.File]::WriteAllText($target, $text, $utf8NoBom)
}

$serviceFiles = @(
    'Services\AppMenus.cs',
    'Services\AppPathsService.cs',
    'Services\AppRoles.cs',
    'Services\BmesMaterialService.cs',
    'Services\BmesRoutingScrapeService.cs',
    'Services\BmesSettingsSyncService.cs',
    'Services\FCostModels.cs',
    'Services\FCostReportService.cs',
    'Services\FCostService.cs',
    'Services\HierReportSupport.cs',
    'Services\ModelGroupPickerHelpers.cs',
    'Services\Models.cs',
    'Services\NgRateExcelExporter.cs',
    'Services\NgRateReportService.cs',
    'Services\NgRateService.cs',
    'Services\NgRateSettingsService.cs',
    'Services\WebRepository.cs',
    'Services\WorkerStatusExcelExporter.cs',
    'Services\WorkerStatusService.cs',
    'Services\WpfSettingsReader.cs'
)

$componentFiles = @(
    'Components\Shared\HierSubRows.razor',
    'Components\Pages\AdminPathsPage.razor',
    'Components\Pages\BmesFCostPage.razor',
    'Components\Pages\BmesMakeModelGroupPage.razor',
    'Components\Pages\BmesReasonTablePage.razor',
    'Components\Pages\BmesRoutingTablePage.razor',
    'Components\Pages\BmesSettingPage.razor',
    'Components\Pages\NgRateAllPage.razor',
    'Components\Pages\NgRateByGroupPage.razor',
    'Components\Pages\NgRateForWeeklyReportPage.razor',
    'Components\Pages\NgRatePage.razor',
    'Components\Pages\SubGroupNode.razor',
    'Components\Pages\WorkerStatusPage.razor'
)

foreach ($file in $serviceFiles) {
    Copy-Converted $file $file
}

foreach ($file in $componentFiles) {
    Copy-Converted $file $file
}
