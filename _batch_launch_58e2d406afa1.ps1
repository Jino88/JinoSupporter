$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try { chcp 65001 | Out-Null } catch {}
try { $Host.UI.RawUI.WindowTitle = 'AI Batch 1/3' } catch {}
Set-Location -LiteralPath 'D:\000. MyWorks\005. Program\Repository\JinoSupporter'

$promptPath = 'D:\000. MyWorks\005. Program\Repository\JinoSupporter\_batch_prompt_58e2d406afa1.txt'
$scriptPath = 'D:\000. MyWorks\005. Program\Repository\JinoSupporter\_batch_launch_58e2d406afa1.ps1'
$prompt = Get-Content -LiteralPath $promptPath -Raw -Encoding UTF8

& claude $prompt
$exitCode = if ($null -eq $LASTEXITCODE) { 0 } else { $LASTEXITCODE }
if ($exitCode -ne 0) {
    Write-Host ('Claude exited with code ' + $exitCode + '.') -ForegroundColor Yellow
}

Remove-Item -LiteralPath $promptPath -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $scriptPath -Force -ErrorAction SilentlyContinue