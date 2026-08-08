$ErrorActionPreference = 'Stop'

$root = 'D:\000. MyWorks\005. Program\Repository\JinoSupporter\.codex\worktrees\codex-s31-codex-session-31-bacfac04\JinoSupporter.Web\Components\Pages'
$files = @(
    'DataInferenceDbPage.razor',
    'DataInferenceDetailPage.razor',
    'DataInferenceInputTestPage.razor',
    'DataInferenceAskPage.razor',
    'DataInferenceBatchPage.razor',
    'DataInferenceAnalysisPage.razor',
    'DataInferenceModelAnalysisPage.razor',
    'DataInferenceValidationPage.razor',
    'DataInferencePage_Test.razor'
)

$map = @{}
function Map-Colors([string[]] $colors, [string] $token) {
    foreach ($color in $colors) { $script:map[$color.ToLowerInvariant()] = $token }
}

Map-Colors @('#fff', '#ffffff') 'var(--panel, #ffffff)'
Map-Colors @('#0f1319') 'var(--ink, #0f1319)'
Map-Colors @('#16202b', '#17324a', '#28405c') 'var(--signal, #16202b)'
Map-Colors @('#4a5461', '#344054', '#354253', '#475467', '#475569', '#4b5563', '#526171', '#526173', '#5b6b7e') 'var(--ink-2, #4a5461)'
Map-Colors @('#7f8a98', '#64748b', '#667085', '#758195', '#8b98a8', '#94a3b8', '#98a2b3', '#aebdca') 'var(--ink-3, #7f8a98)'
Map-Colors @('#e8eaee') 'var(--plane, #e8eaee)'
Map-Colors @('#f5f6f9', '#f3f6f9', '#f3f7fb', '#f4f6f8', '#f4f8fb', '#f5f7fb', '#f6f7fb', '#f8fafb', '#f8fafc', '#f8fbfe', '#fafbfc', '#fbfcfe') 'var(--panel-2, #f5f6f9)'
Map-Colors @('#eceef2', '#edf1f5', '#edf2f7', '#e2e8f0') 'var(--panel-3, #eceef2)'
Map-Colors @('#dde1e7', '#d4dee8', '#d7e1ee', '#d8dee9', '#d9e0e7', '#c8d1db', '#c9d6e2', '#cbd5e1', '#cfd8e3') 'var(--line, #dde1e7)'
Map-Colors @('#c6ccd5', '#d0d5dd') 'var(--line-2, #c6ccd5)'
Map-Colors @('#edf0f4') 'var(--signal-wash, #edf0f4)'

Map-Colors @('#1f5fa8', '#1d4ed8', '#2563eb', '#3b82f6', '#0369a1', '#075985', '#1e3a8a', '#3866a3', '#38bdf8', '#7dd3fc') 'var(--ok, #1f5fa8)'
Map-Colors @('#eaf1f9', '#dbeafe', '#eff6ff', '#bfdbfe', '#e0f2fe', '#bae6fd', '#f0f9ff', '#f0f7ff', '#eaf2ff', '#bfd4ff', '#93c5fd') 'var(--ok-wash, #eaf1f9)'
Map-Colors @('#10b981', '#16a34a', '#027a48', '#1f7a7a', '#21673c', '#2f7d4b') 'var(--ok, #1f5fa8)'
Map-Colors @('#d1fae5', '#ecfdf3', '#abefc6', '#a9d9b8', '#6ee7b7', '#edf8f0', '#f3fbf5', '#f4fbf6') 'var(--ok-wash, #eaf1f9)'

Map-Colors @('#c8342a', '#9b2419', '#c24135', '#b42318', '#dc2626') 'var(--crit, #c8342a)'
Map-Colors @('#fbedec', '#fca5a5', '#efb2aa', '#fecdca', '#fef3f2', '#fff1ef', '#fff7f5', '#fff7f7') 'var(--crit-wash, #fbedec)'

Map-Colors @('#fab219', '#c47a1b', '#fb923c', '#fdba74', '#fbbf24', '#facc15') 'var(--warn, #fab219)'
Map-Colors @('#92610a', '#8a570f', '#78350f', '#9a3412', '#7c2d12', '#6e470d') 'var(--warn-ink, #92610a)'
Map-Colors @('#f0cf8a', '#edcf94', '#fed7aa') 'var(--warn-line, #f0cf8a)'
Map-Colors @('#fdf4e0', '#fff7ed', '#ffedd5', '#fff8eb', '#fffaf0') 'var(--warn-wash, #fdf4e0)'

Map-Colors @('#6366f1', '#7c3aed', '#3440a0', '#86198f') 'var(--signal, #16202b)'
Map-Colors @('#fae8ff') 'var(--signal-wash, #edf0f4)'
Map-Colors @('#f0abfc') 'var(--signal-line, #c3cad3)'

$hexPattern = '#[0-9a-fA-F]{3,8}\b'

function Is-ProtectedLine([string] $line) {
    if ($line -match '^\s*(?://|///)') { return $true }
    if ($line -match 'style\s*=\s*"[^"#]*#') { return $false }
    return $line -match '^\s*(?:return\s+["@]|["@]|(?:sb|html)\.(?:Append|AppendLine)|html\s*=)'
}

function Replace-SheetColors([string] $text) {
    return [regex]::Replace($text, $hexPattern, {
        param($match)
        $before = $text.Substring(0, $match.Index)
        $lineStart = $before.LastIndexOf("`n") + 1
        $line = $text.Substring($lineStart, $text.IndexOf("`n", $match.Index) - $lineStart)
        if (Is-ProtectedLine $line) { return $match.Value }

        $lastVar = $before.LastIndexOf('var(')
        $lastClose = $before.LastIndexOf(')')
        if ($lastVar -gt $lastClose) { return $match.Value }

        $key = $match.Value.ToLowerInvariant()
        if ($map.ContainsKey($key)) { return $map[$key] }
        return $match.Value
    })
}

foreach ($file in $files) {
    $path = Join-Path $root $file
    $text = [IO.File]::ReadAllText($path)
    $text = Replace-SheetColors $text

    # Keep the actual page styles on the shared design-system font tokens.
    $text = [regex]::Replace($text, 'font-family:\s*''Segoe UI'',\s*sans-serif', 'font-family: var(--sans)')
    $text = [regex]::Replace($text, 'font-family:\s*var\(--mono,\s*Consolas,\s*monospace\)', 'font-family: var(--mono)')

    # The sheet uses tiny square corners; circles are status dots/avatars, not corners.
    $text = [regex]::Replace($text, '(?i)(border-radius\s*:\s*)(?!50%)([^;}{"\r\n]+)', {
        param($match)
        $before = $text.Substring(0, $match.Index)
        $lineStart = $before.LastIndexOf("`n") + 1
        $lineEnd = $text.IndexOf("`n", $match.Index)
        if ($lineEnd -lt 0) { $lineEnd = $text.Length }
        $line = $text.Substring($lineStart, $lineEnd - $lineStart)
        if (Is-ProtectedLine $line) { return $match.Value }
        return $match.Groups[1].Value + '2px'
    })

    [IO.File]::WriteAllText($path, $text, [Text.UTF8Encoding]::new($false))
}

# Hosted pages must size against their flex parent, not the browser viewport.
$detail = Join-Path $root 'DataInferenceDetailPage.razor'
$text = [IO.File]::ReadAllText($detail)
$text = $text.Replace('min-height: 100vh; background:', 'min-height: 0; background:')
[IO.File]::WriteAllText($detail, $text, [Text.UTF8Encoding]::new($false))

$model = Join-Path $root 'DataInferenceModelAnalysisPage.razor'
$text = [IO.File]::ReadAllText($model)
$text = $text.Replace('height: 100vh;', 'height: 100%;')
$text = $text.Replace('max-height: 100vh;', 'max-height: 100%;')
[IO.File]::WriteAllText($model, $text, [Text.UTF8Encoding]::new($false))

foreach ($file in @('DataInferenceDbPage.razor', 'DataInferenceAskPage.razor')) {
    $path = Join-Path $root $file
    $text = [IO.File]::ReadAllText($path)
    $text = $text.Replace('border-collapse: separate;', 'border-collapse: collapse;')
    $text = $text.Replace('transition: transform .12s, box-shadow .12s;', 'transition: transform .12s;')
    [IO.File]::WriteAllText($path, $text, [Text.UTF8Encoding]::new($false))
}
