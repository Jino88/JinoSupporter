$ErrorActionPreference = "Continue"

$serviceRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $serviceRoot

$runRoot = Join-Path $serviceRoot "outputs\corpus-ingest\full-989-v1"
$stdoutPath = Join-Path $runRoot "benchmark-repair-false-pass-9-v35.stdout.log"
$stderrPath = Join-Path $runRoot "benchmark-repair-false-pass-9-v35.stderr.log"
$resultPath = Join-Path $runRoot "benchmark-repair-false-pass-9-v35.result.json"
$exitPath = Join-Path $runRoot "benchmark-repair-false-pass-9-v35.exit.json"
$sourceRoot = "D:\000. MyWorks\test\result\InputDataFinish"

"Started: $((Get-Date).ToString('o'))" |
    Out-File -LiteralPath $stdoutPath -Encoding utf8 -Force
"" | Out-File -LiteralPath $stderrPath -Encoding utf8 -Force

$relativePaths = @(
    "01. BRS-161016DT  Report  New bond EV 562850 and improve NG separate VP+CD Date  05.5.2025_1778470595_clean.xlsx",
    "01. TIU C11-20  Report test VP new film 2-2 roll high function NG rate ( A2 rate 42% ) recheck low air pressure NG rate date 17.12.2025 - Copy_1778470674_clean.xlsx",
    "013.MSU-20S15-07 Result test new frame improve_clean.xlsx",
    "015.MSU - L20S15-07 REPORT TEST CM+SM FOLLOW SUPPLIER IMPROVE GAUSS FOR MODULE date 2025.02.13_clean.xlsx",
    "02. L20S15-07DT Report test new lot CD (3-17)  ( Size 510)date   2025.04.14_clean.xlsx",
    "02. TIU L5S3-01 R Report  test F-PCB Improve Solder Pad of vendor CSY TECH VINA -Date 2026.02.04_clean.xlsx",
    "022.MSU - L20S15-07 REPORT TEST NEW BP+SM ASS'Y GUIDE JIG  date 2025.02.21_clean.xlsx",
    "027. MSU-20S15-07 Result check Height dimension C-MG, S-MG - Date 2025.02.28_clean.xlsx",
    "03. TIU L5S3 Result check gauss value after setting machine - date 2025.11.13_clean.xlsx"
)

$items = @()
$exitCode = 0
foreach ($relativePath in $relativePaths) {
    $sourcePath = Join-Path $sourceRoot $relativePath
    $arguments = @(
        "inference_data_ai_cli.py",
        "ingest-workbook",
        "--db", "outputs/universal-grid/InputDataFinish.sqlite",
        "--input", $sourcePath,
        "--artifact-root", "outputs/corpus-ingest/full-989-v1/workbooks",
        "--dataset", "InputDataFinish",
        "--workers", "2",
        "--batch-size", "6",
        "--batch-max-bytes", "240000",
        "--draft-monolithic-max-bytes", "1000000",
        "--draft-fragment-max-chunks", "8",
        "--draft-fragment-max-cells", "2000",
        "--draft-fragment-max-bytes", "400000",
        "--draft-fragment-workers", "3",
        "--reasoning-effort", "medium"
    )
    try {
        $output = & python @arguments 2>> $stderrPath
        $itemExitCode = $LASTEXITCODE
        $output | Out-File -LiteralPath $stdoutPath -Encoding utf8 -Append
    }
    catch {
        $_ | Out-String |
            Out-File -LiteralPath $stderrPath -Encoding utf8 -Append
        $itemExitCode = 1
    }
    if ($itemExitCode -ne 0) {
        $exitCode = 1
    }
    $items += [ordered]@{
        relativePath = $relativePath
        exitCode = $itemExitCode
    }
}

if ($exitCode -eq 0) {
    $reconcileArguments = @(
        "inference_data_ai_cli.py",
        "ingest-corpus",
        "--db", "outputs/universal-grid/InputDataFinish.sqlite",
        "--input", $sourceRoot,
        "--artifact-root", "outputs/corpus-ingest/full-989-v1",
        "--source-manifest", "pilot/corpus-benchmark-repair-false-pass-9-v34.json",
        "--journal", "outputs/corpus-ingest/full-989-v1/corpus-journal.json",
        "--out", $resultPath,
        "--dataset", "InputDataFinish",
        "--inventory-only"
    )
    & python @reconcileArguments 1>> $stdoutPath 2>> $stderrPath
    if ($LASTEXITCODE -ne 0) {
        $exitCode = 1
    }
}

$terminal = [ordered]@{
    finishedAt = (Get-Date).ToString("o")
    exitCode = $exitCode
    items = $items
}
$terminal | ConvertTo-Json -Depth 5 |
    Out-File -LiteralPath $exitPath -Encoding utf8 -Force
"Finished: $($terminal.finishedAt); ExitCode: $exitCode" |
    Out-File -LiteralPath $stdoutPath -Encoding utf8 -Append

exit $exitCode
