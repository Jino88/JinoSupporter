$ErrorActionPreference = 'Stop'

$standaloneRoot = Split-Path -Parent $PSScriptRoot
$repoRoot = Split-Path -Parent $standaloneRoot
$webRoot = Join-Path $repoRoot 'JinoSupporter.Web'

if (-not (Test-Path -LiteralPath $webRoot)) {
    throw "JinoSupporter.Web not found at $webRoot"
}

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)

function Convert-Text([string] $text) {
    # Nested namespaces are rewritten before the plain Services rules below. The @using rule
    # for Services is an unanchored prefix match, so a generic pass would rewrite the outer
    # segment of a nested name first and leave the specific rules with nothing to match.
    $text = $text -replace 'namespace JinoSupporter\.Web\.Services\.BmesReports\.Contracts;', 'namespace BmesNgRateStandalone.Services.BmesReports.Contracts;'
    $text = $text -replace 'namespace JinoSupporter\.Web\.Services\.BmesReports;', 'namespace BmesNgRateStandalone.Services.BmesReports;'
    $text = $text -replace 'using JinoSupporter\.Web\.Services\.BmesReports\.Contracts;', 'using BmesNgRateStandalone.Services.BmesReports.Contracts;'
    $text = $text -replace 'using JinoSupporter\.Web\.Services\.BmesReports;', 'using BmesNgRateStandalone.Services.BmesReports;'
    $text = $text -replace '@using\s+JinoSupporter\.Web\.Services\.BmesReports\.Contracts', '@using BmesNgRateStandalone.Services.BmesReports.Contracts'
    $text = $text -replace '@using\s+JinoSupporter\.Web\.Services\.BmesReports', '@using BmesNgRateStandalone.Services.BmesReports'
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

# Cuts one whole type declaration out of a synced file. Used to keep a source file inside
# the standalone's dependency boundary without hand-editing the generated copy. The throws
# make a boundary that no longer applies fail the build instead of silently removing
# nothing, so a rename or move on the web side surfaces here rather than as drift.
function Remove-TypeDeclaration([string] $text, [string] $declaration, [string] $reason) {
    $start = $text.IndexOf($declaration, [System.StringComparison]::Ordinal)
    if ($start -lt 0) {
        throw "Sync boundary guard: '$declaration' no longer exists in the web source. Re-check the sync boundary in $PSCommandPath."
    }

    $open = $text.IndexOf('{', $start)
    if ($open -lt 0) {
        throw "Sync boundary guard: '$declaration' has no body."
    }

    $depth = 0
    $end = -1
    for ($i = $open; $i -lt $text.Length; $i++) {
        if ($text[$i] -eq '{') {
            $depth++
        }
        elseif ($text[$i] -eq '}') {
            $depth--
            if ($depth -eq 0) {
                $end = $i + 1
                break
            }
        }
    }

    if ($end -lt 0) {
        throw "Sync boundary guard: unbalanced braces while removing '$declaration'."
    }

    $newline = if ($text.Contains("`r`n")) { "`r`n" } else { "`n" }
    $head = $text.Substring(0, $start).TrimEnd()
    $tail = $text.Substring($end).TrimStart()

    $result = $head + $newline + $newline + "// Removed by tools/SyncBmesFromWeb.ps1: $reason" + $newline
    if ($tail.Length -gt 0) {
        $result += $newline + $tail + $newline
    }

    return $result
}

function Copy-Converted([string] $sourceRelative, [string] $targetRelative, [scriptblock] $postConvert = $null) {
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
    if ($null -ne $postConvert) {
        $text = & $postConvert $text
    }

    # Writing unconditionally rewrote every synced file on every build. The text is almost
    # always identical, but the fresh timestamp is enough to make Git report the file as
    # modified and to make MSBuild recompile it, so a clean checkout stopped looking clean
    # the first time it was built and an orchestrated worker saw files it never touched.
    if (Test-Path -LiteralPath $target) {
        $existing = [System.IO.File]::ReadAllText($target, [System.Text.Encoding]::UTF8)
        if ($existing -ceq $text) {
            return
        }
    }

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

# The BMES report calculation slice that Components\Pages\BmesFCostPage.razor needs: the
# F-Cost calculation service it injects, the models it returns, the projection helper that
# service calls, and the DTO contracts underneath. The contracts are self-contained data
# types, so all seven come over as one unit. Everything else under Services\BmesReports
# stays in the web app: the daily, weekly, cause-monthly and KPI calculation services, the
# orchestrator that drives them, and the viewer bootstrap, which belongs to the web report
# host and its token store. The standalone ships none of that.
$reportCalculationFiles = @(
    'Services\BmesReports\Contracts\BmesReportDocumentDto.cs',
    'Services\BmesReports\Contracts\CauseMonthlyTabDtos.cs',
    'Services\BmesReports\Contracts\DailyTabDtos.cs',
    'Services\BmesReports\Contracts\FCostTabDtos.cs',
    'Services\BmesReports\Contracts\KpiTabDtos.cs',
    'Services\BmesReports\Contracts\ReportCommonDtos.cs',
    'Services\BmesReports\Contracts\WeeklyTabDtos.cs',
    'Services\BmesReports\BmesReportProjection.cs',
    'Services\BmesReports\BmesFCostReportCalculationService.cs'
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

foreach ($file in $reportCalculationFiles) {
    Copy-Converted $file $file
}

# BmesReportModels.cs holds the calculation snapshots the F-Cost page needs, but it also
# declares the orchestrator's result type. That one type is the only thing in the synced
# set that reaches past the boundary above: it carries FCostCorePartsKpiSnapshot and
# IpgDefectKpiSnapshot, owned by two web-only services the standalone does not build.
# Dropping the type keeps the boundary closed without copying those services in or
# declaring hollow stand-ins for them here.
Copy-Converted 'Services\BmesReports\BmesReportModels.cs' 'Services\BmesReports\BmesReportModels.cs' {
    param([string] $text)
    Remove-TypeDeclaration $text 'public sealed class BmesReportGenerationResult' `
        'orchestrator result type, outside the standalone F-Cost calculation boundary'
}

foreach ($file in $componentFiles) {
    Copy-Converted $file $file
}

Copy-Converted 'wwwroot\js\app.js' 'wwwroot\js\app.js'
