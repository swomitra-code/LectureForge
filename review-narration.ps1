param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$ManifestPath,

    [ValidateRange(1024, 65535)]
    [int]$Port = 8765,

    [switch]$NoBrowser
)

$ErrorActionPreference = 'Stop'

function Resolve-ProjectPath([string]$manifestDirectory, [string]$projectRoot, [string]$path) {
    if ([IO.Path]::IsPathRooted($path)) { return [IO.Path]::GetFullPath($path) }
    $manifestRelative = [IO.Path]::GetFullPath((Join-Path $manifestDirectory $path))
    if (Test-Path -LiteralPath $manifestRelative) { return $manifestRelative }
    return [IO.Path]::GetFullPath((Join-Path $projectRoot $path))
}

function Get-TakePath([object]$slide, [int]$takeNumber, [string]$manifestDirectory, [string]$projectRoot) {
    $filename = "slide-$($slide.slide_number)-take-$takeNumber.mp3"
    foreach ($configuredTake in @($slide.narration.generated_takes)) {
        if ([string]::IsNullOrWhiteSpace([string]$configuredTake)) { continue }
        $path = Resolve-ProjectPath $manifestDirectory $projectRoot ([string]$configuredTake)
        if ([IO.Path]::GetFileName($path) -eq $filename) { return $path }
    }
    return $null
}

function Get-Duration([string]$path) {
    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
    $text = ((& ffprobe -v error -show_entries 'format=duration' -of 'default=noprint_wrappers=1:nokey=1' $path) -join '').Trim()
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($text)) { return $null }
    $duration = 0.0
    if (-not [double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$duration)) { return $null }
    return $duration
}

function Read-Manifest([string]$path) {
    return Get-Content -Raw -LiteralPath $path | ConvertFrom-Json
}

function Get-FormValues([string]$body) {
    $values = @{}
    foreach ($pair in $body -split '&') {
        $parts = $pair -split '=', 2
        $name = [Uri]::UnescapeDataString(($parts[0] -replace '\+', ' '))
        $value = if ($parts.Count -gt 1) { [Uri]::UnescapeDataString(($parts[1] -replace '\+', ' ')) } else { '' }
        $values[$name] = $value
    }
    return $values
}

function Send-Bytes($stream, [byte[]]$bytes, [string]$contentType, [int]$statusCode = 200, [hashtable]$extraHeaders = @{}) {
    $statusText = switch ($statusCode) { 200 {'OK'} 204 {'No Content'} 206 {'Partial Content'} 303 {'See Other'} 400 {'Bad Request'} 404 {'Not Found'} 416 {'Range Not Satisfiable'} default {'Internal Server Error'} }
    $header = "HTTP/1.1 $statusCode $statusText`r`nContent-Type: $contentType`r`nContent-Length: $($bytes.Length)`r`nConnection: close`r`nCache-Control: no-store`r`nX-Content-Type-Options: nosniff`r`n"
    foreach ($name in $extraHeaders.Keys) { $header += "$name`: $($extraHeaders[$name])`r`n" }
    $header += "`r`n"
    $headerBytes = [Text.Encoding]::ASCII.GetBytes($header)
    $stream.Write($headerBytes, 0, $headerBytes.Length)
    if ($bytes.Length -gt 0) { $stream.Write($bytes, 0, $bytes.Length) }
    $stream.Flush()
}

function Send-Text($stream, [string]$text, [string]$contentType = 'text/html; charset=utf-8', [int]$statusCode = 200, [hashtable]$extraHeaders = @{}) {
    Send-Bytes $stream ([Text.Encoding]::UTF8.GetBytes($text)) $contentType $statusCode $extraHeaders
}

function Send-Audio($stream, [hashtable]$requestHeaders, [string]$path) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { Send-Text $stream 'Not found.' 'text/plain; charset=utf-8' 404; return }
    $file = Get-Item -LiteralPath $path
    $start = 0L
    $end = $file.Length - 1
    $partial = $false
    $range = [string]$requestHeaders['Range']
    if ($range -match '^bytes=(\d+)-(\d*)$') {
        $start = [int64]$matches[1]
        if (-not [string]::IsNullOrWhiteSpace($matches[2])) { $end = [math]::Min([int64]$matches[2], $end) }
        if ($start -gt $end -or $start -ge $file.Length) { Send-Text $stream '' 'text/plain' 416; return }
        $partial = $true
    }
    $length = $end - $start + 1
    $statusCode = if ($partial) { 206 } else { 200 }
    $extraHeaders = @{ 'Accept-Ranges' = 'bytes' }
    if ($partial) { $extraHeaders['Content-Range'] = "bytes $start-$end/$($file.Length)" }
    $statusText = if ($partial) { 'Partial Content' } else { 'OK' }
    $header = "HTTP/1.1 $statusCode $statusText`r`nContent-Type: audio/mpeg`r`nContent-Length: $length`r`nConnection: close`r`nCache-Control: no-store`r`nX-Content-Type-Options: nosniff`r`n"
    foreach ($name in $extraHeaders.Keys) { $header += "$name`: $($extraHeaders[$name])`r`n" }
    $header += "`r`n"
    $headerBytes = [Text.Encoding]::ASCII.GetBytes($header)
    $stream.Write($headerBytes, 0, $headerBytes.Length)
    $fileStream = [IO.File]::OpenRead($path)
    try {
        [void]$fileStream.Seek($start, [IO.SeekOrigin]::Begin)
        $buffer = New-Object byte[] 65536
        $remaining = $length
        while ($remaining -gt 0) {
            $read = $fileStream.Read($buffer, 0, [math]::Min($buffer.Length, $remaining))
            if ($read -le 0) { break }
            $stream.Write($buffer, 0, $read)
            $remaining -= $read
        }
    }
    finally { $fileStream.Dispose(); $stream.Flush() }
}

function Build-Page([object]$manifest, [int[]]$reviewSlideNumbers, [string]$manifestDirectory, [string]$projectRoot, [string]$flash) {
    $encoder = [Net.WebUtility]
    $cards = [Text.StringBuilder]::new()
    $selectedCount = 0
    foreach ($number in $reviewSlideNumbers) {
        $slide = $manifest.slides | Where-Object { [int]$_.slide_number -eq $number }
        if ($null -eq $slide) { continue }
        $currentFilename = if ([string]::IsNullOrWhiteSpace([string]$slide.narration.instructor_selected_take)) { '' } else { [IO.Path]::GetFileName(([string]$slide.narration.instructor_selected_take -replace '/', '\')) }
        $selectedTake = 0
        for ($take = 1; $take -le 3; $take++) { if ($currentFilename -eq "slide-$number-take-$take.mp3") { $selectedTake = $take } }
        if ($selectedTake -gt 0) { $selectedCount++ }
        $scriptText = ''
        if (-not [string]::IsNullOrWhiteSpace([string]$slide.narration.speech_script)) {
            $scriptPath = Resolve-ProjectPath $manifestDirectory $projectRoot ([string]$slide.narration.speech_script)
            if (Test-Path -LiteralPath $scriptPath -PathType Leaf) { $scriptText = [IO.File]::ReadAllText($scriptPath, [Text.Encoding]::UTF8) }
        }
        [void]$cards.Append("<section class='card'><h2>Slide $number</h2><div class='script'>$($encoder::HtmlEncode($scriptText))</div><div class='takes'>")
        for ($take = 1; $take -le 3; $take++) {
            $takePath = Get-TakePath $slide $take $manifestDirectory $projectRoot
            $duration = Get-Duration $takePath
            $durationText = if ($null -eq $duration) { 'Unavailable' } else { '{0:0.000} seconds' -f $duration }
            $selectedClass = if ($selectedTake -eq $take) { ' take selected' } else { ' take' }
            $badge = if ($selectedTake -eq $take) { "<span class='badge'>Selected</span>" } else { '' }
            $audio = if ($null -eq $duration) { "<p class='error'>Audio unavailable</p>" } else { "<audio controls preload='metadata' src='/audio?slide=$number&amp;take=$take'></audio>" }
            $disabled = if ($null -eq $duration) { ' disabled' } else { '' }
            [void]$cards.Append("<article class='$selectedClass'><h3>Take $take $badge</h3>$audio<p class='duration'>$durationText</p><form method='post' action='/select'><input type='hidden' name='slide' value='$number'><input type='hidden' name='take' value='$take'><button type='submit'$disabled>Select Take $take</button></form></article>")
        }
        [void]$cards.Append('</div></section>')
    }
    $total = $reviewSlideNumbers.Count
    $remaining = $total - $selectedCount
    $completion = if ($remaining -eq 0) { "<div class='complete'><strong>All narration selections complete.</strong><br>Next command:<br><code>.\lectureforge.ps1 .\config\module1.json</code></div>" } else { '' }
    $flashHtml = if ([string]::IsNullOrWhiteSpace($flash)) { '' } else { "<div class='flash'>$($encoder::HtmlEncode($flash))</div>" }
    return @"
<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>LectureForge Narration Review</title>
<style>
body{margin:0;background:#f3f5f7;color:#17212b;font-family:Segoe UI,Arial,sans-serif}main{max-width:1100px;margin:0 auto;padding:32px 20px 60px}h1{margin:0 0 8px}.summary,.complete,.flash{background:#fff;border:1px solid #d8dee5;border-radius:10px;padding:16px;margin:16px 0}.complete{border-color:#78b28a;background:#eef8f1}.flash{border-color:#79a7d3;background:#eef6fd}.card{background:#fff;border:1px solid #d8dee5;border-radius:12px;padding:22px;margin:22px 0}.script{white-space:pre-wrap;line-height:1.55;background:#f8fafb;border-left:4px solid #3977a8;padding:14px}.takes{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px;margin-top:18px}.take{border:1px solid #d8dee5;border-radius:10px;padding:14px}.take.selected{border:2px solid #2f855a;background:#f0fff4}.badge{font-size:.75rem;color:#fff;background:#2f855a;border-radius:999px;padding:3px 8px}audio{width:100%}.duration{color:#52606d}button{width:100%;padding:10px;border:0;border-radius:7px;background:#22689b;color:#fff;font-weight:600;cursor:pointer}button:disabled{background:#9aa5b1;cursor:not-allowed}code{display:inline-block;margin-top:8px;padding:7px;background:#17212b;color:#fff;border-radius:5px}@media(max-width:760px){.takes{grid-template-columns:1fr}}
</style></head><body><main><h1>LectureForge Narration Review</h1><div class="summary"><strong>$total slides awaiting review</strong><br>$selectedCount slides selected<br>$remaining remaining</div>$flashHtml$completion$cards</main></body></html>
"@
}

$manifestFile = (Resolve-Path -LiteralPath $ManifestPath).Path
$manifestDirectory = Split-Path -Parent $manifestFile
$projectRoot = Split-Path -Parent $manifestDirectory
$selector = Join-Path $projectRoot 'select-narration.ps1'
if (-not (Test-Path -LiteralPath $selector -PathType Leaf)) { throw "Missing selection tool: $selector" }

$startupManifest = Read-Manifest $manifestFile
$reviewSlideNumbers = @($startupManifest.slides | Where-Object {
    $null -ne $_.narration -and
    @($_.narration.generated_takes).Count -gt 0 -and
    [string]::IsNullOrWhiteSpace([string]$_.narration.instructor_selected_take)
} | Sort-Object slide_number | ForEach-Object { [int]$_.slide_number })
$reviewSet = [Collections.Generic.HashSet[int]]::new()
foreach ($number in $reviewSlideNumbers) { [void]$reviewSet.Add($number) }

$prefix = "http://127.0.0.1:$Port/"
$listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, $Port)
$flash = ''
try {
    $listener.Start()
    Write-Output "LectureForge narration review: $prefix"
    Write-Output 'Press Ctrl+C or close this PowerShell process to stop the local review server.'
    if (-not $NoBrowser) { Start-Process $prefix }
    while ($true) {
        $client = $listener.AcceptTcpClient()
        $network = $null
        $reader = $null
        try {
            $network = $client.GetStream()
            $reader = [IO.StreamReader]::new($network, [Text.Encoding]::ASCII, $false, 8192, $true)
            $requestLine = $reader.ReadLine()
            if ([string]::IsNullOrWhiteSpace($requestLine)) { continue }
            $requestParts = $requestLine -split ' ', 3
            if ($requestParts.Count -lt 2) { Send-Text $network 'Bad request.' 'text/plain; charset=utf-8' 400; continue }
            $method = $requestParts[0]
            $target = $requestParts[1]
            $headers = @{}
            while (($line = $reader.ReadLine()) -ne '') {
                $separator = $line.IndexOf(':')
                if ($separator -gt 0) { $headers[$line.Substring(0, $separator).Trim()] = $line.Substring($separator + 1).Trim() }
            }
            $body = ''
            $contentLength = 0
            if ($headers.ContainsKey('Content-Length')) { [void][int]::TryParse([string]$headers['Content-Length'], [ref]$contentLength) }
            if ($contentLength -gt 0) {
                $characters = New-Object char[] $contentLength
                $offset = 0
                while ($offset -lt $contentLength) {
                    $read = $reader.Read($characters, $offset, $contentLength - $offset)
                    if ($read -le 0) { break }
                    $offset += $read
                }
                $body = -join $characters[0..($offset - 1)]
            }
            $uri = [Uri]::new("http://127.0.0.1:$Port$target")
            if ($method -eq 'GET' -and $uri.AbsolutePath -eq '/') {
                $page = Build-Page (Read-Manifest $manifestFile) $reviewSlideNumbers $manifestDirectory $projectRoot $flash
                $flash = ''
                Send-Text $network $page
            }
            elseif ($method -eq 'GET' -and $uri.AbsolutePath -eq '/audio') {
                $query = Get-FormValues $uri.Query.TrimStart('?')
                $slideNumber = 0; $takeNumber = 0
                if (-not [int]::TryParse([string]$query.slide, [ref]$slideNumber) -or -not [int]::TryParse([string]$query.take, [ref]$takeNumber) -or
                    -not $reviewSet.Contains($slideNumber) -or $takeNumber -lt 1 -or $takeNumber -gt 3) { Send-Text $network 'Not found.' 'text/plain; charset=utf-8' 404; continue }
                $slide = (Read-Manifest $manifestFile).slides | Where-Object { [int]$_.slide_number -eq $slideNumber }
                Send-Audio $network $headers (Get-TakePath $slide $takeNumber $manifestDirectory $projectRoot)
            }
            elseif ($method -eq 'POST' -and $uri.AbsolutePath -eq '/select') {
                $form = Get-FormValues $body
                $slideNumber = 0; $takeNumber = 0
                if (-not [int]::TryParse([string]$form.slide, [ref]$slideNumber) -or -not [int]::TryParse([string]$form.take, [ref]$takeNumber) -or
                    -not $reviewSet.Contains($slideNumber) -or $takeNumber -lt 1 -or $takeNumber -gt 3) { Send-Text $network 'Invalid selection.' 'text/plain; charset=utf-8' 400; continue }
                $selectionOutput = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $selector $manifestFile $slideNumber $takeNumber 2>&1
                if ($LASTEXITCODE -ne 0) { throw "Selection failed: $($selectionOutput -join ' ')" }
                $flash = "Selected Slide $slideNumber Take $takeNumber. You may change it before closing this page."
                Send-Text $network '' 'text/plain; charset=utf-8' 303 @{ Location = '/' }
            }
            elseif ($uri.AbsolutePath -eq '/favicon.ico') { Send-Text $network '' 'text/plain' 204 }
            else { Send-Text $network 'Not found.' 'text/plain; charset=utf-8' 404 }
        }
        catch {
            if ($null -ne $network -and $network.CanWrite) {
                try { Send-Text $network 'Local review error. See the PowerShell window.' 'text/plain; charset=utf-8' 500 }
                catch { }
            }
            Write-Warning "Local review request failed; listener remains active: $($_.Exception.Message)"
        }
        finally {
            if ($null -ne $reader) { $reader.Dispose() }
            if ($null -ne $network) { $network.Dispose() }
            $client.Close()
        }
    }
}
finally {
    $listener.Stop()
}
