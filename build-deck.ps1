param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$ManifestPath
)

$ErrorActionPreference = 'Stop'

function Release-ComObject($object) {
    if ($null -ne $object) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object) }
}

function Resolve-ManifestPath([string]$baseDirectory, [string]$path) {
    if ([IO.Path]::IsPathRooted($path)) { return [IO.Path]::GetFullPath($path) }
    return [IO.Path]::GetFullPath((Join-Path $baseDirectory $path))
}

function Get-NonMediaSlideSignature($slide) {
    $items = [Collections.Generic.List[string]]::new()
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $shape = $slide.Shapes.Item($i)
        if ($shape.Type -eq 16) { continue }
        $text = ''
        try {
            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) { $text = $shape.TextFrame.TextRange.Text }
        } catch { }
        $items.Add("$($shape.Type)|$([math]::Round($shape.Left,2))|$([math]::Round($shape.Top,2))|$([math]::Round($shape.Width,2))|$([math]::Round($shape.Height,2))|$text")
    }
    return ($items -join "`n")
}

function Get-FullSlideSignature($slide) {
    $items = [Collections.Generic.List[string]]::new()
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $shape = $slide.Shapes.Item($i)
        $text = ''
        try {
            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) { $text = $shape.TextFrame.TextRange.Text }
        } catch { }
        $items.Add("$($shape.Type)|$([math]::Round($shape.Left,2))|$([math]::Round($shape.Top,2))|$([math]::Round($shape.Width,2))|$([math]::Round($shape.Height,2))|$text")
    }
    return ($items -join "`n")
}

function Invoke-FreshPlaybackProbe([string]$deckPath, [int]$slideNumber) {
    $ppt = $null
    $presentation = $null
    $slideShowWindow = $null
    try {
        $ppt = New-Object -ComObject PowerPoint.Application
        $presentation = $ppt.Presentations.Open($deckPath, $true, $false, $false)
        $settings = $presentation.SlideShowSettings
        $settings.RangeType = 2
        $settings.StartingSlide = $slideNumber
        $settings.EndingSlide = $slideNumber
        $settings.ShowType = 1
        $slideShowWindow = $settings.Run()
        Start-Sleep -Milliseconds 750
        $liveSlide = $slideShowWindow.View.Slide
        $liveMediaId = $null
        for ($i = 1; $i -le $liveSlide.Shapes.Count; $i++) {
            if ($liveSlide.Shapes.Item($i).Type -eq 16) { $liveMediaId = $liveSlide.Shapes.Item($i).Id }
        }
        if ($null -eq $liveMediaId) {
            return [pscustomobject]@{ advanced = $false; inconclusive = $true; error = 'No media shape found in the live Slide Show view.' }
        }
        $player = $slideShowWindow.View.Player($liveMediaId)
        $before = [double]$player.CurrentPosition
        $player.Play()
        Start-Sleep -Seconds 2
        $after = [double]$player.CurrentPosition
        return [pscustomobject]@{ advanced = ($after -gt $before); inconclusive = $false; error = $null }
    }
    catch {
        return [pscustomobject]@{ advanced = $false; inconclusive = $false; error = $_.Exception.Message }
    }
    finally {
        if ($slideShowWindow) { try { $slideShowWindow.View.Exit() } catch { } }
        if ($presentation) { try { $presentation.Close() } catch { } }
        if ($ppt) { try { $ppt.Quit() } catch { } }
        Release-ComObject $slideShowWindow
        Release-ComObject $presentation
        Release-ComObject $ppt
    }
}

$manifestFile = (Resolve-Path -LiteralPath $ManifestPath).Path
$manifestDirectory = Split-Path -Parent $manifestFile
$manifest = Get-Content -Raw -LiteralPath $manifestFile | ConvertFrom-Json

foreach ($required in 'source_pptx', 'output_pptx', 'output_media_dir', 'defaults', 'slides') {
    if ($null -eq $manifest.$required) { throw "Manifest is missing required property '$required'." }
}

$sourceDeck = Resolve-ManifestPath $manifestDirectory $manifest.source_pptx
$outputDeck = Resolve-ManifestPath $manifestDirectory $manifest.output_pptx
$outputMediaDirectory = Resolve-ManifestPath $manifestDirectory $manifest.output_media_dir
$workDirectory = Join-Path $manifestDirectory '..\work\manifest-build'
$workDirectory = [IO.Path]::GetFullPath($workDirectory)
$validationReport = [IO.Path]::ChangeExtension($outputDeck, '.validation.json')

if ($sourceDeck -eq $outputDeck) { throw 'source_pptx and output_pptx must be different files.' }
if (-not (Test-Path -LiteralPath $sourceDeck -PathType Leaf)) { throw "Missing source deck: $sourceDeck" }

$enabledSlides = @($manifest.slides | Where-Object { $_.enabled -eq $true } | Sort-Object slide_number)
if ($enabledSlides.Count -eq 0) { throw 'Manifest has no enabled slides.' }
$duplicates = @($enabledSlides | Group-Object slide_number | Where-Object Count -gt 1)
if ($duplicates.Count -gt 0) { throw "Duplicate enabled slide number: $($duplicates[0].Name)" }

$defaults = $manifest.defaults.presenter
foreach ($required in 'crop', 'scale', 'position') {
    if ($null -eq $defaults.$required) { throw "Manifest defaults.presenter is missing '$required'." }
}

$jobs = [Collections.Generic.List[object]]::new()
foreach ($slideConfig in $enabledSlides) {
    $number = [int]$slideConfig.slide_number
    if ($number -lt 1) { throw "Invalid slide_number: $number" }
    if ([string]::IsNullOrWhiteSpace($slideConfig.avatar_webm)) { throw "Slide $number is missing avatar_webm." }
    $avatar = Resolve-ManifestPath $manifestDirectory $slideConfig.avatar_webm
    if (-not (Test-Path -LiteralPath $avatar -PathType Leaf)) { throw "Missing approved avatar WebM for Slide $number`: $avatar" }

    $scaleWidth = [int]$defaults.scale.width
    $scaleHeight = [int]$defaults.scale.height
    $positionX = [int]$defaults.position.x
    $positionY = [int]$defaults.position.y
    if ($null -ne $slideConfig.presenter -and $null -ne $slideConfig.presenter.scale) {
        if ($null -ne $slideConfig.presenter.scale.width) { $scaleWidth = [int]$slideConfig.presenter.scale.width }
        if ($null -ne $slideConfig.presenter.scale.height) { $scaleHeight = [int]$slideConfig.presenter.scale.height }
    }
    if ($null -ne $slideConfig.presenter -and $null -ne $slideConfig.presenter.position) {
        if ($null -ne $slideConfig.presenter.position.x) { $positionX = [int]$slideConfig.presenter.position.x }
        if ($null -ne $slideConfig.presenter.position.y) { $positionY = [int]$slideConfig.presenter.position.y }
    }

    $jobs.Add([pscustomobject]@{
        slide_number = $number
        avatar = $avatar
        rendered_slide = Join-Path $workDirectory "slide-$number.png"
        output_video = Join-Path $outputMediaDirectory "slide-$number-final.mp4"
        scale_width = $scaleWidth
        scale_height = $scaleHeight
        position_x = $positionX
        position_y = $positionY
    })
}

New-Item -ItemType Directory -Force -Path $workDirectory, $outputMediaDirectory, (Split-Path -Parent $outputDeck) | Out-Null

$sourceSlideSignatures = @{}
$sourceUnderlaySignatures = @{}
$sourceMediaCounts = @{}
$ppt = $null
$source = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $source = $ppt.Presentations.Open($sourceDeck, $true, $false, $false)
    $sourceSlideCount = $source.Slides.Count
    foreach ($job in $jobs) {
        if ($job.slide_number -gt $sourceSlideCount) { throw "Source deck has fewer than $($job.slide_number) slides." }
    }

    $enabledNumbers = @($jobs | ForEach-Object { $_.slide_number })
    for ($n = 1; $n -le $sourceSlideCount; $n++) {
        if ($enabledNumbers -notcontains $n) { $sourceSlideSignatures[$n] = Get-FullSlideSignature $source.Slides.Item($n) }
    }
    foreach ($job in $jobs) {
        $slide = $source.Slides.Item($job.slide_number)
        $sourceUnderlaySignatures[$job.slide_number] = Get-NonMediaSlideSignature $slide
        $mediaCount = 0
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) { if ($slide.Shapes.Item($i).Type -eq 16) { $mediaCount++ } }
        $sourceMediaCounts[$job.slide_number] = $mediaCount
        $slide.Export($job.rendered_slide, 'PNG', 1920, 1080)
    }
}
finally {
    if ($source) { try { $source.Close() } catch { } }
    if ($ppt) { try { $ppt.Quit() } catch { } }
    Release-ComObject $source
    Release-ComObject $ppt
}

$videoReports = [ordered]@{}
foreach ($job in $jobs) {
    $crop = $defaults.crop
    # Locked Slides 10-12 golden-path filter: explicit VP9 alpha decode, split alpha,
    # RGB-only green despill, alpha merge, Lanczos scale, and lower-right overlay.
    $filter = "[1:v]crop=$($crop.width):$($crop.height):$($crop.x):$($crop.y),format=rgba,split=2[rgb0][a0];[a0]alphaextract[a];[rgb0]despill=type=green:mix=0.30:expand=0.05:alpha=0[rgb];[rgb][a]alphamerge,scale=$($job.scale_width):$($job.scale_height):flags=lanczos[p];[0:v][p]overlay=x=$($job.position_x):y=$($job.position_y):format=auto:shortest=1[v]"
    & ffmpeg -y -hide_banner -loglevel warning `
        -loop 1 -framerate 30 -i $job.rendered_slide `
        -c:v libvpx-vp9 -i $job.avatar `
        -filter_complex $filter `
        -map '[v]' -map '1:a:0' `
        -c:v libx264 -preset medium -crf 18 -pix_fmt yuv420p -r 30 `
        -c:a aac -b:a 192k -ar 48000 -ac 2 `
        -movflags '+faststart' -shortest $job.output_video
    if ($LASTEXITCODE -ne 0) { throw "FFmpeg failed for Slide $($job.slide_number) with exit code $LASTEXITCODE" }

    $probeText = (& ffprobe -v error -show_entries 'format=duration:stream=codec_type,codec_name,width,height' -of json $job.output_video) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "ffprobe failed for Slide $($job.slide_number)." }
    $probe = $probeText | ConvertFrom-Json
    $videoStream = @($probe.streams | Where-Object codec_type -eq 'video')[0]
    $audioStream = @($probe.streams | Where-Object codec_type -eq 'audio')[0]
    if ($videoStream.codec_name -ne 'h264' -or $audioStream.codec_name -ne 'aac' -or $videoStream.width -ne 1920 -or $videoStream.height -ne 1080) {
        throw "Slide $($job.slide_number) output is not 1920x1080 H.264/AAC."
    }
    & ffmpeg -v error -i $job.output_video -map '0:v:0' -map '0:a:0' -f null NUL
    if ($LASTEXITCODE -ne 0) { throw "Decode verification failed for Slide $($job.slide_number)." }
    $videoReports[[string]$job.slide_number] = [ordered]@{
        path = $job.output_video
        duration_seconds = [double]$probe.format.duration
        video_codec = $videoStream.codec_name
        audio_codec = $audioStream.codec_name
        width = $videoStream.width
        height = $videoStream.height
        decode_ok = $true
    }
}

# Build one output copy, open it once, embed every enabled slide, then save once.
Copy-Item -LiteralPath $sourceDeck -Destination $outputDeck -Force
$ppt = $null
$result = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $result = $ppt.Presentations.Open($outputDeck, $false, $false, $false)
    foreach ($job in $jobs) {
        $slide = $result.Slides.Item($job.slide_number)
        $media = $slide.Shapes.AddMediaObject2($job.output_video, $false, $true, 0, 0, $result.PageSetup.SlideWidth, $result.PageSetup.SlideHeight)
        $media.Left = 0
        $media.Top = 0
        $media.Width = $result.PageSetup.SlideWidth
        $media.Height = $result.PageSetup.SlideHeight
        $media.AnimationSettings.Animate = -1
        $media.AnimationSettings.PlaySettings.PlayOnEntry = -1
        for ($i = 1; $i -le $slide.TimeLine.MainSequence.Count; $i++) {
            $effect = $slide.TimeLine.MainSequence.Item($i)
            if ($effect.Shape.Id -eq $media.Id -and $effect.EffectType -eq 83) { $effect.Timing.TriggerType = 2 }
        }
    }
    $result.Save()
}
finally {
    if ($result) { try { $result.Close() } catch { } }
    if ($ppt) { try { $ppt.Quit() } catch { } }
    Release-ComObject $result
    Release-ComObject $ppt
}

# Reopen once to verify unrelated slides, then isolate each media playback probe in
# its own PowerPoint lifecycle for reliable desktop validation.
$unrelatedVerified = [Collections.Generic.List[int]]::new()
$ppt = $null
$result = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $result = $ppt.Presentations.Open($outputDeck, $true, $false, $false)
    $reopenedSlideCount = $result.Slides.Count
    foreach ($n in $sourceSlideSignatures.Keys) {
        if ($sourceSlideSignatures[$n] -eq (Get-FullSlideSignature $result.Slides.Item([int]$n))) { $unrelatedVerified.Add([int]$n) }
    }
}
finally {
    if ($result) { try { $result.Close() } catch { } }
    if ($ppt) { try { $ppt.Quit() } catch { } }
    Release-ComObject $result
    Release-ComObject $ppt
}

$slideReports = [ordered]@{}
foreach ($job in $jobs) {
    $ppt = $null
    $result = $null
    try {
        $ppt = New-Object -ComObject PowerPoint.Application
        $result = $ppt.Presentations.Open($outputDeck, $true, $false, $false)
        if ($result.Slides.Count -ne $sourceSlideCount) { throw 'Slide count changed after reopening the output deck.' }
        $slide = $result.Slides.Item($job.slide_number)
        $underlayPreserved = ($sourceUnderlaySignatures[$job.slide_number] -eq (Get-NonMediaSlideSignature $slide))

        $mediaShapes = @()
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) { if ($slide.Shapes.Item($i).Type -eq 16) { $mediaShapes += $slide.Shapes.Item($i) } }
        $expectedMediaCount = [int]$sourceMediaCounts[$job.slide_number] + 1
        if ($mediaShapes.Count -ne $expectedMediaCount) { throw "Unexpected media count on Slide $($job.slide_number)." }
        $media = $mediaShapes[-1]
        $linked = $false
        try { $linked = [bool]$media.LinkFormat.SourceFullName } catch { $linked = $false }
        $automaticEffect = $false
        for ($i = 1; $i -le $slide.TimeLine.MainSequence.Count; $i++) {
            $effect = $slide.TimeLine.MainSequence.Item($i)
            if ($effect.Shape.Id -eq $media.Id -and $effect.EffectType -eq 83 -and $effect.Timing.TriggerType -eq 2) { $automaticEffect = $true }
        }
        $structuralReport = [ordered]@{
            media_shape_count = $mediaShapes.Count
            expected_media_shape_count = $expectedMediaCount
            media_is_linked = $linked
            media_fills_slide = ([math]::Abs($media.Left) -lt 0.1 -and [math]::Abs($media.Top) -lt 0.1 -and [math]::Abs($media.Width - $result.PageSetup.SlideWidth) -lt 0.1 -and [math]::Abs($media.Height - $result.PageSetup.SlideHeight) -lt 0.1)
            starts_automatically = $automaticEffect
            play_on_entry = ($media.AnimationSettings.PlaySettings.PlayOnEntry -eq -1)
            editable_underlay_preserved = $underlayPreserved
        }
    }
    finally {
        if ($result) { try { $result.Close() } catch { } }
        if ($ppt) { try { $ppt.Quit() } catch { } }
        Release-ComObject $result
        Release-ComObject $ppt
    }

    $probeResults = [Collections.Generic.List[object]]::new()
    $probe = Invoke-FreshPlaybackProbe $outputDeck $job.slide_number
    $probeResults.Add($probe)
    if ($probe.inconclusive) {
        $probe = Invoke-FreshPlaybackProbe $outputDeck $job.slide_number
        $probeResults.Add($probe)
    }
    $slideReports[[string]$job.slide_number] = [ordered]@{
        media_shape_count = $structuralReport.media_shape_count
        expected_media_shape_count = $structuralReport.expected_media_shape_count
        media_is_linked = $structuralReport.media_is_linked
        media_fills_slide = $structuralReport.media_fills_slide
        starts_automatically = $structuralReport.starts_automatically
        play_on_entry = $structuralReport.play_on_entry
        playback_position_advanced = $probe.advanced
        playback_error = $probe.error
        playback_probe_attempts = $probeResults.Count
        playback_probe_results = @($probeResults)
        editable_underlay_preserved = $structuralReport.editable_underlay_preserved
    }
}

$report = [ordered]@{
    manifest = $manifestFile
    source_pptx = $sourceDeck
    output_pptx = $outputDeck
    source_slide_count = $sourceSlideCount
    output_slide_count = $reopenedSlideCount
    slide_count_unchanged = ($sourceSlideCount -eq $reopenedSlideCount)
    requested_slides = @($jobs | ForEach-Object { $_.slide_number })
    unrelated_slides_verified = @($unrelatedVerified | Sort-Object)
    unrelated_slides_unchanged = ($unrelatedVerified.Count -eq $sourceSlideSignatures.Count)
    videos = $videoReports
    slides = $slideReports
}
$report | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $validationReport -Encoding UTF8
$report | ConvertTo-Json -Depth 7

$passed = $report.slide_count_unchanged -and $report.unrelated_slides_unchanged
foreach ($job in $jobs) {
    $s = $slideReports[[string]$job.slide_number]
    $passed = $passed -and -not $s.media_is_linked -and $s.media_fills_slide -and $s.starts_automatically -and $s.play_on_entry -and $s.playback_position_advanced -and $s.editable_underlay_preserved
}
if (-not $passed) { throw "Build validation failed. See $validationReport" }

Write-Output "Built and validated: $outputDeck"
Write-Output "Validation report: $validationReport"
