$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$sourceDeck = Join-Path $root 'source\Module1_Energy_Landscape_2026_Refresh_Cleaned_No_Old_Photo_FINAL2.pptx'
$approvedSlide10Deck = Join-Path $root 'output\Module1_Energy_Landscape_VIDEO_PILOT.pptx'
$outputDeck = Join-Path $root 'output\Module1_Energy_Landscape_VIDEO_SLIDES_10_12.pptx'
$reportPath = Join-Path $root 'output\validation-slides10-12.json'
$work = Join-Path $root 'work'
$output = Join-Path $root 'output'

$slides = @(
    [ordered]@{
        number = 11
        avatar = Join-Path $root 'media\slide-11-SKM-BLUE-7-production-transparent.webm'
        png = Join-Path $work 'slide-11.png'
        video = Join-Path $output 'slide-11-final.mp4'
    },
    [ordered]@{
        number = 12
        avatar = Join-Path $root 'media\slide-12-SKM-BLUE-7-production-transparent.webm'
        png = Join-Path $work 'slide-12.png'
        video = Join-Path $output 'slide-12-final.mp4'
    }
)

New-Item -ItemType Directory -Force -Path $work, $output | Out-Null
foreach ($path in @($sourceDeck, $approvedSlide10Deck) + @($slides | ForEach-Object { $_.avatar })) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing required input: $path" }
}

function Release-ComObject($object) {
    if ($null -ne $object) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object) }
}

function Get-SlideSignature($slide, [switch]$ExcludeMedia) {
    $items = [Collections.Generic.List[string]]::new()
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $shape = $slide.Shapes.Item($i)
        if ($ExcludeMedia -and $shape.Type -eq 16) { continue }
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

# Render the original editable slides exactly as PowerPoint displays them.
$ppt = $null
$presentation = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $presentation = $ppt.Presentations.Open($sourceDeck, $true, $false, $false)
    foreach ($item in $slides) {
        $presentation.Slides.Item($item.number).Export($item.png, 'PNG', 1920, 1080)
    }
}
finally {
    if ($presentation) { $presentation.Close() }
    if ($ppt) { $ppt.Quit() }
    Release-ComObject $presentation
    Release-ComObject $ppt
}

# Locked Slide 10 golden-path filter and encoding settings.
$filter = '[1:v]crop=640:620:870:80,format=rgba,split=2[rgb0][a0];[a0]alphaextract[a];[rgb0]despill=type=green:mix=0.30:expand=0.05:alpha=0[rgb];[rgb][a]alphamerge,scale=280:271:flags=lanczos[p];[0:v][p]overlay=x=1620:y=430:format=auto:shortest=1[v]'
foreach ($item in $slides) {
    & ffmpeg -y -hide_banner -loglevel warning `
        -loop 1 -framerate 30 -i $item.png `
        -c:v libvpx-vp9 -i $item.avatar `
        -filter_complex $filter `
        -map '[v]' -map '1:a:0' `
        -c:v libx264 -preset medium -crf 18 -pix_fmt yuv420p -r 30 `
        -c:a aac -b:a 192k -ar 48000 -ac 2 `
        -movflags '+faststart' -shortest $item.video
    if ($LASTEXITCODE -ne 0) { throw "FFmpeg failed for Slide $($item.number) with exit code $LASTEXITCODE" }
}

# Preserve the approved Slide 10 deck exactly, then add only Slides 11 and 12.
Copy-Item -LiteralPath $approvedSlide10Deck -Destination $outputDeck -Force

$ppt = $null
$presentation = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $presentation = $ppt.Presentations.Open($outputDeck, $false, $false, $false)
    foreach ($item in $slides) {
        $slide = $presentation.Slides.Item($item.number)
        $media = $slide.Shapes.AddMediaObject2($item.video, $false, $true, 0, 0, $presentation.PageSetup.SlideWidth, $presentation.PageSetup.SlideHeight)
        $media.Left = 0
        $media.Top = 0
        $media.Width = $presentation.PageSetup.SlideWidth
        $media.Height = $presentation.PageSetup.SlideHeight
        $media.AnimationSettings.Animate = -1
        $media.AnimationSettings.PlaySettings.PlayOnEntry = -1
        for ($i = 1; $i -le $slide.TimeLine.MainSequence.Count; $i++) {
            $effect = $slide.TimeLine.MainSequence.Item($i)
            if ($effect.Shape.Id -eq $media.Id -and $effect.EffectType -eq 83) { $effect.Timing.TriggerType = 2 }
        }
    }
    $presentation.Save()
}
finally {
    if ($presentation) { $presentation.Close() }
    if ($ppt) { $ppt.Quit() }
    Release-ComObject $presentation
    Release-ComObject $ppt
}

# Close/reopen in desktop PowerPoint and validate structure, embedding, autoplay, and playback.
$ppt = $null
$source = $null
$approved = $null
$result = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $source = $ppt.Presentations.Open($sourceDeck, $true, $false, $false)
    $approved = $ppt.Presentations.Open($approvedSlide10Deck, $true, $false, $false)
    $result = $ppt.Presentations.Open($outputDeck, $true, $false, $false)

    $unrelatedVerified = [Collections.Generic.List[int]]::new()
    foreach ($n in @(1..9) + @(13..15)) {
        if ((Get-SlideSignature $source.Slides.Item($n)) -eq (Get-SlideSignature $result.Slides.Item($n))) { $unrelatedVerified.Add($n) }
    }

    $editableUnderlayVerified = [Collections.Generic.List[int]]::new()
    foreach ($n in 10, 11, 12) {
        if ((Get-SlideSignature $source.Slides.Item($n) -ExcludeMedia) -eq (Get-SlideSignature $result.Slides.Item($n) -ExcludeMedia)) { $editableUnderlayVerified.Add($n) }
    }

    $slide10Preserved = ((Get-SlideSignature $approved.Slides.Item(10)) -eq (Get-SlideSignature $result.Slides.Item(10)))
    $slideReports = [ordered]@{}
    foreach ($n in 10, 11, 12) {
        $targetSlide = $result.Slides.Item($n)
        $mediaShapes = @()
        for ($i = 1; $i -le $targetSlide.Shapes.Count; $i++) {
            $candidate = $targetSlide.Shapes.Item($i)
            if ($candidate.Type -eq 16) { $mediaShapes += $candidate }
        }
        if ($mediaShapes.Count -ne 1) { throw "Expected exactly one media shape on reopened Slide $n; found $($mediaShapes.Count)." }
        $savedMedia = $mediaShapes[0]
        $linked = $false
        try { $linked = [bool]$savedMedia.LinkFormat.SourceFullName } catch { $linked = $false }

        $automaticEffect = $false
        for ($i = 1; $i -le $targetSlide.TimeLine.MainSequence.Count; $i++) {
            $effect = $targetSlide.TimeLine.MainSequence.Item($i)
            if ($effect.Shape.Id -eq $savedMedia.Id -and $effect.EffectType -eq 83 -and $effect.Timing.TriggerType -eq 2) { $automaticEffect = $true }
        }

        $playbackAdvanced = $false
        $playbackError = $null
        $slideShowWindow = $null
        try {
            $settings = $result.SlideShowSettings
            $settings.RangeType = 2
            $settings.StartingSlide = $n
            $settings.EndingSlide = $n
            $settings.ShowType = 1
            $slideShowWindow = $settings.Run()
            Start-Sleep -Milliseconds 750
            $liveSlide = $slideShowWindow.View.Slide
            $liveMediaId = $null
            for ($i = 1; $i -le $liveSlide.Shapes.Count; $i++) {
                if ($liveSlide.Shapes.Item($i).Type -eq 16) { $liveMediaId = $liveSlide.Shapes.Item($i).Id }
            }
            if ($null -eq $liveMediaId) { throw 'No media shape found in the live Slide Show view.' }
            $player = $slideShowWindow.View.Player($liveMediaId)
            $before = [double]$player.CurrentPosition
            $player.Play()
            Start-Sleep -Seconds 2
            $after = [double]$player.CurrentPosition
            $playbackAdvanced = ($after -gt $before)
        } catch {
            $playbackError = $_.Exception.Message
        } finally {
            if ($slideShowWindow) { try { $slideShowWindow.View.Exit() } catch { } }
            Release-ComObject $slideShowWindow
        }

        $slideReports[[string]$n] = [ordered]@{
            embedded_media_shape_count = $mediaShapes.Count
            media_is_linked = $linked
            media_fills_slide = ([math]::Abs($savedMedia.Left) -lt 0.1 -and [math]::Abs($savedMedia.Top) -lt 0.1 -and [math]::Abs($savedMedia.Width - $result.PageSetup.SlideWidth) -lt 0.1 -and [math]::Abs($savedMedia.Height - $result.PageSetup.SlideHeight) -lt 0.1)
            starts_automatically = $automaticEffect
            play_on_entry = ($savedMedia.AnimationSettings.PlaySettings.PlayOnEntry -eq -1)
            playback_position_advanced = $playbackAdvanced
            playback_error = $playbackError
        }
    }

    $report = [ordered]@{
        source_deck = $sourceDeck
        approved_slide_10_deck = $approvedSlide10Deck
        output_deck = $outputDeck
        output_videos = [ordered]@{
            '10' = Join-Path $root 'output\slide-10-final.mp4'
            '11' = $slides[0].video
            '12' = $slides[1].video
        }
        reopened_in_powerpoint = $true
        source_slide_count = $source.Slides.Count
        output_slide_count = $result.Slides.Count
        slide_count_unchanged = ($source.Slides.Count -eq $result.Slides.Count -and $result.Slides.Count -eq 15)
        approved_slide_10_preserved = $slide10Preserved
        editable_underlays_verified = @($editableUnderlayVerified)
        editable_underlays_preserved = ($editableUnderlayVerified.Count -eq 3)
        unrelated_slides_verified = @($unrelatedVerified)
        unrelated_slides_unchanged = ($unrelatedVerified.Count -eq 12)
        slides = $slideReports
    }
    $report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $reportPath -Encoding UTF8
    $report | ConvertTo-Json -Depth 6

    $allSlideMediaPassed = $true
    foreach ($n in 10, 11, 12) {
        $s = $slideReports[[string]$n]
        if (-not ($s.embedded_media_shape_count -eq 1 -and -not $s.media_is_linked -and $s.media_fills_slide -and $s.starts_automatically -and $s.play_on_entry -and $s.playback_position_advanced)) {
            $allSlideMediaPassed = $false
        }
    }
    if (-not ($report.slide_count_unchanged -and $report.approved_slide_10_preserved -and $report.editable_underlays_preserved -and $report.unrelated_slides_unchanged -and $allSlideMediaPassed)) {
        throw "Validation failed. See $reportPath"
    }
}
finally {
    if ($result) { try { $result.Close() } catch { } }
    if ($approved) { try { $approved.Close() } catch { } }
    if ($source) { try { $source.Close() } catch { } }
    if ($ppt) { try { $ppt.Quit() } catch { } }
    Release-ComObject $result
    Release-ComObject $approved
    Release-ComObject $source
    Release-ComObject $ppt
}
