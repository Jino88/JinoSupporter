$ErrorActionPreference = "Continue"

$serviceRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $serviceRoot

$runRoot = Join-Path $serviceRoot "outputs\corpus-ingest\full-989-v1"
$stdoutPath = Join-Path $runRoot "benchmark-staged-b24-b25-v31.stdout.log"
$stderrPath = Join-Path $runRoot "benchmark-staged-b24-b25-v31.stderr.log"
$resultPath = Join-Path $runRoot "benchmark-staged-b24-b25-v31.result.json"
$exitPath = Join-Path $runRoot "benchmark-staged-b24-b25-v31.exit.json"

"Started: $((Get-Date).ToString('o'))" |
    Out-File -LiteralPath $stdoutPath -Encoding utf8 -Force
"" | Out-File -LiteralPath $stderrPath -Encoding utf8 -Force

$arguments = @(
    "inference_data_ai_cli.py",
    "ingest-corpus",
    "--db", "outputs/universal-grid/InputDataFinish.sqlite",
    "--input", "D:\000. MyWorks\test\result\InputDataFinish",
    "--artifact-root", "outputs/corpus-ingest/full-989-v1",
    "--source-manifest", "pilot/corpus-benchmark-staged-b24-b25-v31.json",
    "--journal", "outputs/corpus-ingest/full-989-v1/corpus-journal.json",
    "--out", $resultPath,
    "--dataset", "InputDataFinish",
    "--workbook-workers", "2",
    "--locator-workers", "2",
    "--batch-size", "6",
    "--batch-max-bytes", "240000",
    "--draft-fragment-max-chunks", "8",
    "--draft-fragment-max-cells", "2000",
    "--draft-fragment-max-bytes", "400000",
    "--draft-fragment-workers", "3",
    "--reasoning-effort", "medium",
    "--retry-failed"
)

try {
    & python @arguments 1>> $stdoutPath 2>> $stderrPath
    $exitCode = $LASTEXITCODE
}
catch {
    $_ | Out-String |
        Out-File -LiteralPath $stderrPath -Encoding utf8 -Append
    $exitCode = 1
}

$terminal = [ordered]@{
    finishedAt = (Get-Date).ToString("o")
    exitCode = $exitCode
}
$terminal | ConvertTo-Json |
    Out-File -LiteralPath $exitPath -Encoding utf8 -Force
"Finished: $($terminal.finishedAt); ExitCode: $exitCode" |
    Out-File -LiteralPath $stdoutPath -Encoding utf8 -Append

exit $exitCode
