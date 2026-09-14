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

function ConvertTo-LFSpokenText([string]$Text) {
    # Only recognized bracketed directives are stage directions. Keep all other
    # text verbatim; a space prevents words on either side of a cue joining.
    $cue='\[\s*(?:(?:POINT|TRACE|CIRCLE|SWEEP)\s*:[^\[\]]*|REVEAL\s*|PAUSE(?:\s+[^\[\]]*)?)\]'
    return [regex]::Replace($Text,$cue,' ',[Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

function Invoke-LFElevenLabs($Settings,[string]$Text,[string]$Output) {
    if (-not $env:ELEVENLABS_API_KEY) { throw 'ELEVENLABS_API_KEY is not set.' }
    if (Test-Path -LiteralPath $Output) { throw 'Provider output already exists.' }
    $spokenText=ConvertTo-LFSpokenText $Text
    # Local, per-attempt record of exactly what is sent; never rewrite the script.
    [IO.File]::WriteAllText(($Output+'.spoken.txt'),$spokenText,[Text.UTF8Encoding]::new($false))
    $payload=@{text=$spokenText;model_id=$Settings.model_id;voice_settings=@{stability=$Settings.stability;similarity_boost=$Settings.similarity_boost;style=$Settings.style;use_speaker_boost=$Settings.use_speaker_boost;speed=$Settings.speed}}
    $json=$payload|ConvertTo-Json -Depth 5 -Compress
    $endpoint='https://api.elevenlabs.io/v1/text-to-speech/'+[Uri]::EscapeDataString($Settings.voice_id)+'?output_format=mp3_44100_128'
    Invoke-WebRequest -UseBasicParsing -TimeoutSec 120 -Method Post -Uri $endpoint -Headers @{'xi-api-key'=$env:ELEVENLABS_API_KEY;Accept='audio/mpeg'} -ContentType 'application/json' -Body ([Text.Encoding]::UTF8.GetBytes($json)) -OutFile $Output | Out-Null
}
Export-ModuleMember -Function Protect-NarrationDiagnostic,Get-NarrationFailureDiagnostic,ConvertTo-LFSpokenText,Invoke-LFElevenLabs
