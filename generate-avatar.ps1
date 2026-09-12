param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$ManifestPath
)

$ErrorActionPreference = 'Stop'

function Resolve-ProjectPath([string]$manifestDirectory, [string]$projectRoot, [string]$path) {
    if ([IO.Path]::IsPathRooted($path)) { return [IO.Path]::GetFullPath($path) }
    $manifestRelative = [IO.Path]::GetFullPath((Join-Path $manifestDirectory $path))
    if (Test-Path -LiteralPath $manifestRelative) { return $manifestRelative }
    return [IO.Path]::GetFullPath((Join-Path $projectRoot $path))
}

function Invoke-HeyGenJson([string]$method, [string]$uri, [hashtable]$headers, [object]$body = $null) {
    $parameters = @{ Method = $method; Uri = $uri; Headers = $headers; UseBasicParsing = $true }
    if ($null -ne $body) {
        $parameters.ContentType = 'application/json'
        $parameters.Body = ($body | ConvertTo-Json -Depth 10 -Compress)
    }
    try { return Invoke-RestMethod @parameters }
    catch {
        $status = $null
        try { $status = [int]$_.Exception.Response.StatusCode } catch { }
        if ($null -ne $status) { throw "HeyGen request failed (HTTP $status)." }
        throw 'HeyGen request failed.'
    }
}

function Get-DurationSeconds([string]$path) {
    $text = ((& ffprobe -v error -show_entries 'format=duration' -of 'default=noprint_wrappers=1:nokey=1' $path) -join '').Trim()
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($text)) { throw "Duration is unreadable: $path" }
    $value = 0.0
    if (-not [double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$value) -or $value -le 0) {
        throw "Duration is invalid: $path"
    }
    return $value
}

$manifestFile = (Resolve-Path -LiteralPath $ManifestPath).Path
$manifestDirectory = Split-Path -Parent $manifestFile
$projectRoot = Split-Path -Parent $manifestDirectory
$manifest = Get-Content -Raw -LiteralPath $manifestFile | ConvertFrom-Json
if ([string]::IsNullOrWhiteSpace($env:HEYGEN_API_KEY)) { throw 'HEYGEN_API_KEY is not set in the environment.' }
if ($null -eq $manifest.defaults.avatar_generation) { throw 'Manifest is missing defaults.avatar_generation.' }

$defaults = $manifest.defaults.avatar_generation
$required = @('avatar_name', 'avatar_id', 'output_dir', 'remove_background', 'output_format', 'aspect_ratio', 'resolution', 'burn_captions')
foreach ($name in $required) {
    if ($null -eq $defaults.$name) { throw "Manifest defaults.avatar_generation is missing '$name'." }
}
if ([string]$defaults.avatar_name -ne 'SKM-BLUE-7') { throw 'The approved avatar/look must be SKM-BLUE-7.' }
if ([string]$defaults.avatar_id -ne 'f04bc24f9fd7480885c266c4f1d71f98') { throw 'The approved SKM-BLUE-7 avatar ID is not configured.' }
if ([string]$defaults.output_format -ne 'webm' -or [string]$defaults.aspect_ratio -ne '16:9' -or [string]$defaults.resolution -ne '1080p') {
    throw 'Avatar output must be WebM, 16:9, and 1080p.'
}
if ($defaults.remove_background -ne $true -or $defaults.burn_captions -ne $false) {
    throw 'Avatar output must remove the background and must not burn captions.'
}

$requestedSlides = @($manifest.slides | Where-Object {
    $null -ne $_.avatar_generation -and $_.avatar_generation.enabled -eq $true -and
    [string]::IsNullOrWhiteSpace([string]$_.avatar_webm) -and $null -ne $_.narration -and
    -not [string]::IsNullOrWhiteSpace([string]$_.narration.instructor_selected_take)
} | Sort-Object slide_number)
if ($requestedSlides.Count -eq 0) {
    Write-Output 'No slides require avatar generation. No provider calls were made.'
    exit 0
}

$headers = @{ 'x-api-key' = $env:HEYGEN_API_KEY }
$outputDirectory = Resolve-ProjectPath $manifestDirectory $projectRoot ([string]$defaults.output_dir)
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
$results = [Collections.Generic.List[object]]::new()

foreach ($slide in $requestedSlides) {
    $slideNumber = [int]$slide.slide_number
    $narrationPath = Resolve-ProjectPath $manifestDirectory $projectRoot ([string]$slide.narration.instructor_selected_take)
    if (-not (Test-Path -LiteralPath $narrationPath -PathType Leaf)) { throw "Selected narration is missing for Slide $slideNumber`: $narrationPath" }
    if ((Get-Item -LiteralPath $narrationPath).Length -le 0) { throw "Selected narration is empty for Slide $slideNumber." }
    & ffmpeg -v error -i $narrationPath -f null NUL
    if ($LASTEXITCODE -ne 0) { throw "Selected narration does not decode for Slide $slideNumber." }
    $narrationDuration = Get-DurationSeconds $narrationPath

    $filename = "slide-$slideNumber-$($defaults.avatar_name)-production-transparent.webm"
    $outputPath = Join-Path $outputDirectory $filename
    if (Test-Path -LiteralPath $outputPath) { throw "Refusing to overwrite existing avatar output: $outputPath" }
    $downloadPath = "$outputPath.download"

    $uploadText = & curl.exe -sS --fail -X POST 'https://api.heygen.com/v3/assets' -H "X-Api-Key: $($env:HEYGEN_API_KEY)" -F "file=@$narrationPath"
    if ($LASTEXITCODE -ne 0) { throw "HeyGen audio upload failed for Slide $slideNumber." }
    $upload = ($uploadText -join '') | ConvertFrom-Json
    $assetId = [string]$upload.data.asset_id
    if ([string]::IsNullOrWhiteSpace($assetId)) { $assetId = [string]$upload.data.id }
    if ([string]::IsNullOrWhiteSpace($assetId)) { throw "HeyGen did not return an audio asset ID for Slide $slideNumber." }

    $body = [ordered]@{
        type = 'avatar'; avatar_id = [string]$defaults.avatar_id
        title = "LectureForge Slide $slideNumber $($defaults.avatar_name) approved narration"
        resolution = [string]$defaults.resolution; aspect_ratio = [string]$defaults.aspect_ratio
        remove_background = [bool]$defaults.remove_background; output_format = [string]$defaults.output_format
        audio_asset_id = $assetId
    }
    $created = Invoke-HeyGenJson 'Post' 'https://api.heygen.com/v3/videos' $headers $body
    $videoId = [string]$created.data.video_id
    if ([string]::IsNullOrWhiteSpace($videoId)) { throw "HeyGen did not return a video ID for Slide $slideNumber." }

    $deadline = [DateTime]::UtcNow.AddMinutes(45)
    $video = $null
    while ([DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Seconds 10
        try {
            $video = (Invoke-HeyGenJson 'Get' "https://api.heygen.com/v3/videos/$videoId" $headers).data
        }
        catch {
            # A transient polling transport failure must not create a duplicate job.
            # Keep polling the existing video ID until the bounded deadline.
            continue
        }
        if ([string]$video.status -eq 'completed') { break }
        if ([string]$video.status -eq 'failed') { throw "HeyGen generation failed for Slide $slideNumber." }
    }
    if ($null -eq $video -or [string]$video.status -ne 'completed') { throw "HeyGen generation timed out for Slide $slideNumber." }
    if ([string]::IsNullOrWhiteSpace([string]$video.video_url)) { throw "HeyGen completed Slide $slideNumber without a download URL." }

    try {
        Invoke-WebRequest -UseBasicParsing -Uri ([string]$video.video_url) -OutFile $downloadPath | Out-Null
        if (-not (Test-Path -LiteralPath $downloadPath -PathType Leaf) -or (Get-Item -LiteralPath $downloadPath).Length -le 0) { throw "Downloaded avatar video is empty for Slide $slideNumber." }
        & ffmpeg -v error -c:v libvpx-vp9 -i $downloadPath -f null NUL
        if ($LASTEXITCODE -ne 0) { throw "Avatar video does not decode for Slide $slideNumber." }
        $probe = (& ffprobe -v error -show_streams -of json $downloadPath) -join '' | ConvertFrom-Json
        $videoStream = @($probe.streams | Where-Object codec_type -eq 'video') | Select-Object -First 1
        $audioStream = @($probe.streams | Where-Object codec_type -eq 'audio') | Select-Object -First 1
        if ($null -eq $videoStream) { throw "Avatar output has no video stream for Slide $slideNumber." }
        if ($null -eq $audioStream) { throw "Avatar output has no audio stream for Slide $slideNumber." }
        $alphaPresent = ([string]$videoStream.tags.alpha_mode -eq '1')
        if (-not $alphaPresent) { throw "Avatar output has no verified alpha channel for Slide $slideNumber." }
        $avatarDuration = Get-DurationSeconds $downloadPath
        $tolerance = [math]::Max(1.0, $narrationDuration * 0.05)
        if ([math]::Abs($avatarDuration - $narrationDuration) -gt $tolerance) { throw "Avatar duration does not approximately match narration for Slide $slideNumber." }

        Move-Item -LiteralPath $downloadPath -Destination $outputPath
        $manifestAvatarPath = "media/$filename"
        $slide.avatar_webm = $manifestAvatarPath
        $results.Add([pscustomobject]@{
            slide_number = $slideNumber; selected_narration = [string]$slide.narration.instructor_selected_take
            avatar_webm = $manifestAvatarPath; narration_duration_seconds = [math]::Round($narrationDuration, 3)
            avatar_duration_seconds = [math]::Round($avatarDuration, 3); alpha_ok = $true
            decode_ok = $true; audio_present = $true
        })
    }
    finally {
        if (Test-Path -LiteralPath $downloadPath) { Remove-Item -LiteralPath $downloadPath -Force }
    }
}

$manifest | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $manifestFile -Encoding UTF8
$results | Format-Table -AutoSize
Write-Output 'Avatar generation and technical validation completed. PowerPoint was not run.'
