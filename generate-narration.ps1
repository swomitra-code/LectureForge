param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$ManifestPath
)

$ErrorActionPreference = 'Stop'

function Resolve-ManifestPath([string]$baseDirectory, [string]$path) {
    if ([IO.Path]::IsPathRooted($path)) { return [IO.Path]::GetFullPath($path) }
    return [IO.Path]::GetFullPath((Join-Path $baseDirectory $path))
}

function Get-ManifestTakePath([string]$configuredOutputDirectory, [int]$slideNumber, [int]$takeNumber) {
    $filename = "slide-$slideNumber-take-$takeNumber.mp3"
    if ([IO.Path]::IsPathRooted($configuredOutputDirectory)) { return (Join-Path $configuredOutputDirectory $filename) }
    return ((Join-Path $configuredOutputDirectory $filename) -replace '\\', '/')
}

function Protect-NarrationDiagnostic([string]$Value, [string[]]$SensitiveValues) {
    foreach ($sensitive in $SensitiveValues) {
        if ([string]::IsNullOrEmpty($sensitive)) { continue }
        $Value = $Value.Replace($sensitive, '[REDACTED]')
        $escaped = ConvertTo-Json -InputObject $sensitive -Compress
        $Value = $Value.Replace($escaped.Substring(1, $escaped.Length - 2), '[REDACTED]')
    }
    # Suppress credential/header lines and any echoed request fields, including partial text.
    $Value = $Value -replace '(?im)^.*(?:\b(?:authorization|xi-api-key|api[_-]?key)["\s]*[:=]|bearer\s|"(?:text|input|body|payload|headers)"\s*:).*$', '[REDACTED]'
    return $Value
}

function Get-NarrationFailureDiagnostic($Failure, [string[]]$SensitiveValues) {
    $parts = [Collections.Generic.List[string]]::new()
    $exception = $Failure.Exception
    $parts.Add("Exception: $($exception.GetType().FullName): $($exception.Message)")
    $inner = $exception.InnerException
    while ($null -ne $inner) {
        $parts.Add("Inner exception: $($inner.GetType().FullName): $($inner.Message)")
        $inner = $inner.InnerException
    }
    $response = $exception.Response
    if ($null -ne $response -and $null -ne $response.StatusCode) {
        $parts.Add("HTTP status: $([int]$response.StatusCode)")
    }
    if ($null -ne $response) {
        foreach ($header in @('request-id', 'x-request-id')) {
            try {
                $requestId = $response.Headers.GetValues($header) -join ', '
                if ($requestId) { $parts.Add("Request ID: $requestId") }
            } catch { } # Header may be absent; never dump the headers collection.
        }
    }
    $providerBody = [string]$Failure.ErrorDetails.Message
    if ([string]::IsNullOrWhiteSpace($providerBody) -and $null -ne $response) {
        try {
            if ($null -ne $response.Content) {
                $providerBody = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            } else {
                $reader = [IO.StreamReader]::new($response.GetResponseStream())
                try { $providerBody = $reader.ReadToEnd() } finally { $reader.Dispose() }
            }
        } catch { } # An unavailable body must not replace the original failure.
    }
    if (-not [string]::IsNullOrWhiteSpace($providerBody)) {
        try {
            $provider = $providerBody | ConvertFrom-Json
            # Allowlist diagnostic fields; validation input and request payloads are excluded.
            foreach ($entry in @($provider, $provider.detail)) {
                if ($null -eq $entry) { continue }
                foreach ($field in @('status', 'message', 'request_id')) {
                    if ($entry.$field -is [string]) { $parts.Add("Provider ${field}: $($entry.$field)") }
                }
                if ($entry -is [string]) { $parts.Add("Provider detail: $entry") }
            }
        } catch {
            $parts.Add('Provider body unavailable as structured JSON; raw body omitted for privacy.')
        }
    }
    return Protect-NarrationDiagnostic ($parts -join [Environment]::NewLine) $SensitiveValues
}

$manifestFile = (Resolve-Path -LiteralPath $ManifestPath).Path
$manifestDirectory = Split-Path -Parent $manifestFile
$manifest = Get-Content -Raw -LiteralPath $manifestFile | ConvertFrom-Json

if ($null -eq $manifest.defaults -or $null -eq $manifest.defaults.narration) {
    throw 'Manifest is missing defaults.narration.'
}
if ($null -eq $manifest.slides) { throw 'Manifest is missing slides.' }

$requestedSlides = @($manifest.slides | Where-Object {
    $null -ne $_.narration -and
    $_.narration.enabled -eq $true -and
    [string]::IsNullOrWhiteSpace([string]$_.narration.instructor_selected_take)
} | Sort-Object slide_number)

if ($requestedSlides.Count -eq 0) {
    Write-Output 'No slides require narration generation. No provider calls were made.'
    exit 0
}

$narrationDefaults = $manifest.defaults.narration
if ([string]::IsNullOrWhiteSpace([string]$narrationDefaults.voice_id)) {
    throw 'Set defaults.narration.voice_id to the approved ElevenLabs voice ID.'
}
if ([string]::IsNullOrWhiteSpace($env:ELEVENLABS_API_KEY)) {
    throw 'ELEVENLABS_API_KEY is not set in the environment.'
}

$requiredDefaults = @('model_id', 'output_dir', 'stability', 'similarity_boost', 'style', 'use_speaker_boost', 'speed')
foreach ($name in $requiredDefaults) {
    if ($null -eq $narrationDefaults.$name) { throw "Manifest defaults.narration is missing '$name'." }
}

$outputDirectory = Resolve-ManifestPath $manifestDirectory $narrationDefaults.output_dir
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
$headers = @{
    'xi-api-key' = $env:ELEVENLABS_API_KEY
    'Accept' = 'audio/mpeg'
}
$voiceId = [Uri]::EscapeDataString([string]$narrationDefaults.voice_id)
$endpoint = "https://api.elevenlabs.io/v1/text-to-speech/$voiceId`?output_format=mp3_44100_128"
$results = [Collections.Generic.List[object]]::new()

foreach ($slide in $requestedSlides) {
    $slideNumber = [int]$slide.slide_number
    if ($slideNumber -lt 1) { throw "Invalid slide_number: $slideNumber" }
    if ([string]::IsNullOrWhiteSpace([string]$slide.narration.speech_script)) {
        throw "Slide $slideNumber is missing narration.speech_script."
    }
    $scriptPath = Resolve-ManifestPath $manifestDirectory $slide.narration.speech_script
    if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) { throw "Missing speech script for Slide $slideNumber`: $scriptPath" }

    # Read the supplied script verbatim. Do not trim, rewrite, or normalize it.
    $scriptText = [IO.File]::ReadAllText($scriptPath, [Text.Encoding]::UTF8)
    if ([string]::IsNullOrEmpty($scriptText)) { throw "Speech script is empty for Slide $slideNumber`: $scriptPath" }

    $generatedTakePaths = [Collections.Generic.List[string]]::new()
    for ($takeNumber = 1; $takeNumber -le 3; $takeNumber++) {
        $manifestTakePath = Get-ManifestTakePath $narrationDefaults.output_dir $slideNumber $takeNumber
        $outputPath = Resolve-ManifestPath $manifestDirectory $manifestTakePath
        $payload = [ordered]@{
            text = $scriptText
            model_id = [string]$narrationDefaults.model_id
            voice_settings = [ordered]@{
                stability = [double]$narrationDefaults.stability
                similarity_boost = [double]$narrationDefaults.similarity_boost
                style = [double]$narrationDefaults.style
                use_speaker_boost = [bool]$narrationDefaults.use_speaker_boost
                speed = [double]$narrationDefaults.speed
            }
        }
        $json = $payload | ConvertTo-Json -Depth 5 -Compress
        $body = [Text.Encoding]::UTF8.GetBytes($json)

        try {
            Invoke-WebRequest -UseBasicParsing -Method Post -Uri $endpoint -Headers $headers -ContentType 'application/json' -Body $body -OutFile $outputPath | Out-Null
        }
        catch {
            $failure = $_
            $diagnostic = Get-NarrationFailureDiagnostic $failure @($env:ELEVENLABS_API_KEY, $json, $scriptText)
            if (Test-Path -LiteralPath $outputPath) { Remove-Item -LiteralPath $outputPath -Force }
            throw "ElevenLabs generation failed for Slide $slideNumber Take $takeNumber.`n$diagnostic"
        }

        $file = Get-Item -LiteralPath $outputPath
        if ($file.Length -le 0) { throw "Generated narration is empty: $outputPath" }
        & ffmpeg -v error -i $outputPath -f null NUL
        if ($LASTEXITCODE -ne 0) { throw "Generated narration does not decode: $outputPath" }
        $durationText = ((& ffprobe -v error -show_entries 'format=duration' -of 'default=noprint_wrappers=1:nokey=1' $outputPath) -join '').Trim()
        if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($durationText)) { throw "Could not read narration duration: $outputPath" }
        $duration = 0.0
        if (-not [double]::TryParse($durationText, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$duration) -or $duration -le 0) {
            throw "Invalid narration duration: $outputPath"
        }

        $generatedTakePaths.Add($manifestTakePath)
        $results.Add([pscustomobject]@{
            slide_number = $slideNumber
            take_number = $takeNumber
            path = $outputPath
            bytes = $file.Length
            duration_seconds = [math]::Round($duration, 3)
            decode_ok = $true
        })
    }

    if ($null -eq $slide.narration.generated_takes) {
        $slide.narration | Add-Member -NotePropertyName generated_takes -NotePropertyValue @($generatedTakePaths)
    } else {
        $slide.narration.generated_takes = @($generatedTakePaths)
    }
}

# Persist only deterministic take paths. Instructor selection remains manual.
$manifest | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $manifestFile -Encoding UTF8
$results | Format-Table slide_number, take_number, path, bytes, duration_seconds, decode_ok -AutoSize
Write-Output 'Generated exactly three takes per requested slide. No HeyGen or PowerPoint production was run.'
