$ErrorActionPreference = "Stop"

$serviceRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $serviceRoot

$stdout = Join-Path $serviceRoot "outputs\corpus-ingest\full-989-v1\capture-989-v26.stdout.log"
$stderr = Join-Path $serviceRoot "outputs\corpus-ingest\full-989-v1\capture-989-v26.stderr.log"

"Started: $(Get-Date -Format o)" | Set-Content -LiteralPath $stdout -Encoding utf8
"" | Set-Content -LiteralPath $stderr -Encoding utf8

& python inference_data_ai_cli.py openxml-index `
    --db outputs/universal-grid/InputDataFinish.sqlite `
    --input "D:\000. MyWorks\test\result\InputDataFinish" `
    --dataset InputDataFinish `
    --workers 4 `
    1>> $stdout `
    2>> $stderr

$exitCode = $LASTEXITCODE
"Finished: $(Get-Date -Format o); ExitCode: $exitCode" |
    Add-Content -LiteralPath $stdout -Encoding utf8
exit $exitCode
