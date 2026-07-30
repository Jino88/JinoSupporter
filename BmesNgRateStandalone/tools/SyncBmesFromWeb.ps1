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
    $text = $text -replace 'namespace JinoSupporter\.Web\.Services\.GraphMaker;', 'namespace BmesNgRateStandalone.Services.GraphMaker;'
    $text = $text -replace 'using JinoSupporter\.Web\.Services;', 'using BmesNgRateStandalone.Services;'
    $text = $text -replace 'using JinoSupporter\.Web\.Services\.GraphMaker;', 'using BmesNgRateStandalone.Services.GraphMaker;'
    $text = $text -replace '@using\s+JinoSupporter\.Web\.Services', '@using BmesNgRateStandalone.Services'
    $text = $text -replace '@using\s+JinoSupporter\.Web\.Services\.GraphMaker', '@using BmesNgRateStandalone.Services.GraphMaker'
    $text = $text -replace '@attribute\s+\[Authorize\]\r?\n', ''
    $text = $text -replace '@rendermode\s+InteractiveServer\r?\n', ''
    $text = [regex]::Replace($text, '<AuthorizeView Roles="Admin">\s*<Authorized>\s*', '')
    $text = [regex]::Replace($text, '\s*</Authorized>\s*</AuthorizeView>', '')
    $text = $text -replace 'JinoSupporter\.Web process restarts', 'the standalone app restarts'
    $text = $text -replace 'href="/admin/settings"', 'href="/bmes/setting"'
    $text = [regex]::Replace(
        $text,
        '<a href="/bmes/setting" class="small text-primary">[^<]*Admin Settings</a>',
        '<a href="/bmes/setting" class="small text-primary">BMES Setting</a>')
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
    'Services\AppStoragePaths.cs',
    'Services\AiProviderSettingsService.cs',
    'Services\BrowserDownload.cs',
    'Services\BmesFcostActualService.cs',
    'Services\BmesMaterialService.cs',
    'Services\BmesRoutingScrapeService.cs',
    'Services\BmesSettingsSyncService.cs',
    'Services\FCostModels.cs',
    'Services\FCostRawBreakdownExcelExporter.cs',
    'Services\FCostReportService.cs',
    'Services\FCostService.cs',
    'Services\HierReportSupport.cs',
    'Services\ModelGroupPickerHelpers.cs',
    'Services\Models.cs',
    'Services\CsvExportUtility.cs',
    'Services\NgRateCsvExporter.cs',
    'Services\NgRateExcelExporter.cs',
    'Services\NgRateModeSupport.cs',
    'Services\NgRateReportService.cs',
    'Services\NgRateService.cs',
    'Services\NgRateSettingsService.cs',
    'Services\WebRepository.cs',
    'Services\WorkerStatusExcelExporter.cs',
    'Services\WorkerStatusService.cs',
    'Services\WpfSettingsReader.cs',
    'Services\GraphMaker\GraphMakerCore.cs'
)

$componentFiles = @(
    'Components\Shared\HierSubRows.razor',
    'Components\Shared\NgRateModelGroupPicker.razor',
    'Components\Shared\NgRateReportStyles.razor',
    'Components\Shared\NgRateSetupPanel.razor',
    'Components\Shared\NgRateSimpleGroupPicker.razor',
    'Components\Shared\NgRateViewNav.razor',
    'Components\Pages\GraphMakerPage.razor',
    'Components\Pages\AdminPathsPage.razor',
    'Components\Pages\BmesFCostPage.razor',
    'Components\Pages\BmesMakeModelGroupPage.razor',
    'Components\Pages\BmesReasonTablePage.razor',
    'Components\Pages\BmesRoutingTablePage.razor',
    'Components\Pages\NgRateAllPage.razor',
    'Components\Pages\NgRateByGroupPage.razor',
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

Copy-Converted 'wwwroot\js\app.js' 'wwwroot\js\app.js'
