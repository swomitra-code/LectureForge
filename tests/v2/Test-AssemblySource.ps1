# Read-only source checks. No provider, media, COM, or production state actions.
$ErrorActionPreference = 'Stop'
$repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Set-Location $repo
$files = @((git diff --name-only), (git ls-files --others --exclude-standard)) | ForEach-Object { $_ }
$count = 0
foreach ($file in $files) {
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Missing changed file: $file" }
    $body = [IO.File]::ReadAllText((Join-Path $repo $file))
    foreach ($secret in @($env:ELEVENLABS_API_KEY, $env:HEYGEN_API_KEY)) {
        if ($secret -and $body.Contains($secret)) { throw "Credential literal found in $file" }
    }
    if ($file -match '\.ps(m)?1$') {
        $errors = $null; $tokens = $null
        $null = [Management.Automation.Language.Parser]::ParseInput($body, [ref]$tokens, [ref]$errors)
        if ($errors) { throw "$file : $($errors.Message -join '; ')" }
        $count++
    }
    if ($file -match '\.(js|cjs)$') {
        & node --check $file
        if ($LASTEXITCODE) { throw "JavaScript syntax failed: $file" }
    }
}
git diff --check
if ($LASTEXITCODE) { throw 'Git whitespace check failed.' }
Write-Output "PASS: $count PowerShell files; JavaScript syntax; credential literals; Git whitespace."
