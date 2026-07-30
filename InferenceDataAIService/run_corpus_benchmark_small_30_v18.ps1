$ErrorActionPreference = "Continue"

$serviceRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $serviceRoot

$stdoutPath = Join-Path $serviceRoot "outputs\corpus-ingest\full-989-v1\benchmark-small-30-v18.stdout.log"
$stderrPath = Join-Path $serviceRoot "outputs\corpus-ingest\full-989-v1\benchmark-small-30-v18.stderr.log"

"Started: $((Get-Date).ToString('o'))" | Out-File -LiteralPath $stdoutPath -Encoding utf8 -Force
"" | Out-File -LiteralPath $stderrPath -Encoding utf8 -Force

$arguments = @(
    "inference_data_ai_cli.py",
    "ingest-corpus",
    "--db", "outputs/universal-grid/InputDataFinish.sqlite",
    "--input", "D:\000. MyWorks\test\result\InputDataFinish",
    "--artifact-root", "outputs/corpus-ingest/full-989-v1",
    "--source-manifest", "pilot/corpus-benchmark-small-30-v1.json",
    "--journal", "outputs/corpus-ingest/full-989-v1/corpus-journal.json",
    "--out", "outputs/corpus-ingest/full-989-v1/benchmark-small-30-v18.result.json",
    "--dataset", "InputDataFinish",
    "--workbook-workers", "3",
    "--locator-workers", "2",
    "--batch-size", "6",
    "--batch-max-bytes", "240000",
    "--reasoning-effort", "medium"
)

try {
    & python @arguments 1>> $stdoutPath 2>> $stderrPath
    $exitCode = $LASTEXITCODE
}
catch {
    $_ | Out-String | Out-File -LiteralPath $stderrPath -Encoding utf8 -Append
    $exitCode = 1
}

"Finished: $((Get-Date).ToString('o')); ExitCode: $exitCode" |
    Out-File -LiteralPath $stdoutPath -Encoding utf8 -Append

exit $exitCode
