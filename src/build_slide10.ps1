param(
    [string]$ConfigPath = 'module1.pilot.json',
    [int]$SlideNumber = 10,
    [switch]$SkipEncode
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$configFile = Join-Path $root $ConfigPath
if (-not (Test-Path -LiteralPath $configFile -PathType Leaf)) { throw "Missing configuration: $configFile" }
$config = Get-Content -Raw -LiteralPath $configFile | ConvertFrom-Json
$slideConfig = $config.slides.PSObject.Properties[[string]$SlideNumber].Value
if ($null -eq $slideConfig) { throw "Slide $SlideNumber is not configured in $configFile" }
$defaults = $config.defaults

$sourceDeck = Join-Path $root $config.deck
$avatar = Join-Path $root $slideConfig.avatar_video
$work = Join-Path $root 'work'
$output = Join-Path $root 'output'
$slidePng = Join-Path $root $slideConfig.rendered_slide
$video = Join-Path $root $slideConfig.output_video
$outputDeck = Join-Path $root $config.output_deck
$reportPath = Join-Path $root $slideConfig.validation_report

New-Item -ItemType Directory -Force -Path $work, $output | Out-Null
foreach ($path in $sourceDeck, $avatar) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing required input: $path" }
}

function Release-ComObject($object) {
    if ($null -ne $object) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object) }
}

function Get-SlideSignature($slide) {
    $items = [Collections.Generic.List[string]]::new()
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $shape = $slide.Shapes.Item($i)
        $text = ''
        try {
            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                $text = $shape.TextFrame.TextRange.Text
            }
        } catch { }
        $items.Add("$($shape.Type)|$([math]::Round($shape.Left,2))|$([math]::Round($shape.Top,2))|$([math]::Round($shape.Width,2))|$([math]::Round($shape.Height,2))|$text")
    }
    return ($items -join "`n")
}

$ppt = $null
$presentation = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $presentation = $ppt.Presentations.Open($sourceDeck, $true, $false, $false)
    if ($presentation.Slides.Count -lt $SlideNumber) { throw "Source deck has fewer than $SlideNumber slides." }
    $presentation.Slides.Item($SlideNumber).Export($slidePng, $defaults.render.format, $defaults.render.width, $defaults.render.height)
}
finally {
    if ($presentation) { $presentation.Close() }
    if ($ppt) { $ppt.Quit() }
    Release-ComObject $presentation
    Release-ComObject $ppt
}

if (-not $SkipEncode) {
    # The VP9 alpha decoder is selected explicitly. Despill touches RGB only; the
    # original alpha plane is split before cleanup and merged back afterward.
    $crop = $defaults.ffmpeg.presenter_crop
    $scale = $defaults.ffmpeg.presenter_scale
    $position = $defaults.ffmpeg.presenter_position
    $despill = $defaults.ffmpeg.despill
    $filter = "[1:v]crop=$($crop.width):$($crop.height):$($crop.x):$($crop.y),format=rgba,split=2[rgb0][a0];[a0]alphaextract[a];[rgb0]despill=type=$($despill.type):mix=$($despill.mix):expand=$($despill.expand):alpha=0[rgb];[rgb][a]alphamerge,scale=$($scale.width):$($scale.height):flags=$($scale.filter)[p];[0:v][p]overlay=x=$($position.x):y=$($position.y):format=auto:shortest=1[v]"
    $videoConfig = $defaults.ffmpeg.video
    $audioConfig = $defaults.ffmpeg.audio
    & ffmpeg -y -hide_banner -loglevel warning `
        -loop 1 -framerate $defaults.render.frame_rate -i $slidePng `
        -c:v $defaults.ffmpeg.alpha_decoder -i $avatar `
        -filter_complex $filter `
        -map '[v]' -map '1:a:0' `
        -c:v $videoConfig.codec -preset $videoConfig.preset -crf $videoConfig.crf -pix_fmt $videoConfig.pixel_format -r $videoConfig.frame_rate `
        -c:a $audioConfig.codec -b:a $audioConfig.bitrate -ar $audioConfig.sample_rate -ac $audioConfig.channels `
        -movflags '+faststart' -shortest $video
    if ($LASTEXITCODE -ne 0) { throw "FFmpeg failed with exit code $LASTEXITCODE" }
}
if (-not (Test-Path -LiteralPath $video -PathType Leaf)) { throw "Missing composed video: $video" }

Copy-Item -LiteralPath $sourceDeck -Destination $outputDeck -Force

$ppt = $null
$presentation = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $presentation = $ppt.Presentations.Open($outputDeck, $false, $false, $false)
    $slide = $presentation.Slides.Item($SlideNumber)
    $media = $slide.Shapes.AddMediaObject2($video, $defaults.powerpoint.link_to_file, $defaults.powerpoint.save_with_document, 0, 0, $presentation.PageSetup.SlideWidth, $presentation.PageSetup.SlideHeight)
    $media.Left = 0
    $media.Top = 0
    $media.Width = $presentation.PageSetup.SlideWidth
    $media.Height = $presentation.PageSetup.SlideHeight
    $media.AnimationSettings.Animate = -1
    $media.AnimationSettings.PlaySettings.PlayOnEntry = -1
    # AddMediaObject2 creates a media-play effect that defaults to on-click.
    # Change that existing effect in place so there is one unambiguous autoplay.
    for ($i = 1; $i -le $slide.TimeLine.MainSequence.Count; $i++) {
        $effect = $slide.TimeLine.MainSequence.Item($i)
        if ($effect.Shape.Id -eq $media.Id -and $effect.EffectType -eq 83) { $effect.Timing.TriggerType = 2 }
    }
    $presentation.Save()
}
finally {
    if ($presentation) { $presentation.Close() }
    if ($ppt) { $ppt.Quit() }
    Release-ComObject $presentation
    Release-ComObject $ppt
}

# Reopen both decks in desktop PowerPoint and validate the saved result.
$ppt = $null
$source = $null
$result = $null
$slideShowWindow = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $source = $ppt.Presentations.Open($sourceDeck, $true, $false, $false)
    $result = $ppt.Presentations.Open($outputDeck, $true, $false, $false)
    $sourceSlideCount = $source.Slides.Count
    $outputSlideCount = $result.Slides.Count
    $unchanged = [Collections.Generic.List[int]]::new()
    for ($n = 1; $n -le $source.Slides.Count; $n++) {
        if ($n -eq $SlideNumber) { continue }
        if ((Get-SlideSignature $source.Slides.Item($n)) -eq (Get-SlideSignature $result.Slides.Item($n))) { $unchanged.Add($n) }
    }

    $targetSlide = $result.Slides.Item($SlideNumber)
    $mediaShapes = @()
    for ($i = 1; $i -le $targetSlide.Shapes.Count; $i++) {
        $candidate = $targetSlide.Shapes.Item($i)
        if ($candidate.Type -eq 16) { $mediaShapes += $candidate }
    }
    if ($mediaShapes.Count -lt 1) { throw "No media shape found on reopened Slide $SlideNumber." }
    $savedMedia = $mediaShapes[-1]
    $linked = $false
    try { $linked = [bool]$savedMedia.LinkFormat.SourceFullName } catch { $linked = $false }

    $automaticEffect = $false
    for ($i = 1; $i -le $targetSlide.TimeLine.MainSequence.Count; $i++) {
        $effect = $targetSlide.TimeLine.MainSequence.Item($i)
        if ($effect.Shape.Id -eq $savedMedia.Id -and $effect.EffectType -eq 83 -and $effect.Timing.TriggerType -eq 2) { $automaticEffect = $true }
    }

    $playbackAdvanced = $false
    $playbackError = $null
    try {
        $settings = $result.SlideShowSettings
        $settings.RangeType = 2
        $settings.StartingSlide = $SlideNumber
        $settings.EndingSlide = $SlideNumber
        $settings.ShowType = 1
        $slideShowWindow = $settings.Run()
        Start-Sleep -Milliseconds 750
        $liveMediaId = $null
        $liveSlide = $slideShowWindow.View.Slide
        for ($i = 1; $i -le $liveSlide.Shapes.Count; $i++) {
            if ($liveSlide.Shapes.Item($i).Type -eq 16) { $liveMediaId = $liveSlide.Shapes.Item($i).Id }
        }
        if ($null -eq $liveMediaId) { throw 'No media shape found in the live Slide Show view.' }
        $player = $slideShowWindow.View.Player($liveMediaId)
        $before = [double]$player.CurrentPosition
        $player.Play()
        Start-Sleep -Seconds $defaults.powerpoint.playback_probe_seconds
        $after = [double]$player.CurrentPosition
        $playbackAdvanced = ($after -gt $before)
    } catch {
        $playbackError = $_.Exception.Message
    } finally {
        if ($slideShowWindow) { try { $slideShowWindow.View.Exit() } catch { } }
    }

    $report = [ordered]@{
        source_deck = $sourceDeck
        output_deck = $outputDeck
        output_video = $video
        reopened_in_powerpoint = $true
        source_slide_count = $sourceSlideCount
        output_slide_count = $outputSlideCount
        slide_count_unchanged = ($sourceSlideCount -eq $outputSlideCount)
        embedded_media_shape_count_slide_10 = $mediaShapes.Count
        media_is_linked = $linked
        media_fills_slide = ([math]::Abs($savedMedia.Left) -lt 0.1 -and [math]::Abs($savedMedia.Top) -lt 0.1 -and [math]::Abs($savedMedia.Width - $result.PageSetup.SlideWidth) -lt 0.1 -and [math]::Abs($savedMedia.Height - $result.PageSetup.SlideHeight) -lt 0.1)
        starts_automatically = $automaticEffect
        playback_position_advanced = $playbackAdvanced
        playback_error = $playbackError
        unrelated_slides_unchanged = ($unchanged.Count -eq ($source.Slides.Count - 1))
        unrelated_slides_verified = @($unchanged)
    }
    $report | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $reportPath -Encoding UTF8
    $report | ConvertTo-Json -Depth 4

    $failed = -not ($report.slide_count_unchanged -and -not $report.media_is_linked -and $report.media_fills_slide -and $report.starts_automatically -and $report.playback_position_advanced -and $report.unrelated_slides_unchanged)
    if ($failed) { throw "Validation failed. See $reportPath" }
}
finally {
    if ($result) { try { $result.Close() } catch { } }
    if ($source) { try { $source.Close() } catch { } }
    if ($ppt) { try { $ppt.Quit() } catch { } }
    Release-ComObject $slideShowWindow
    Release-ComObject $result
    Release-ComObject $source
    Release-ComObject $ppt
}
