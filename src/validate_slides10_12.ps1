$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$sourceDeck = Join-Path $root 'source\Module1_Energy_Landscape_2026_Refresh_Cleaned_No_Old_Photo_FINAL2.pptx'
$approvedDeck = Join-Path $root 'output\Module1_Energy_Landscape_VIDEO_PILOT.pptx'
$outputDeck = Join-Path $root 'output\Module1_Energy_Landscape_VIDEO_SLIDES_10_12.pptx'
$reportPath = Join-Path $root 'output\validation-slides10-12.json'

function Release-ComObject($object) {
    if ($null -ne $object) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object) }
}

function Get-PptxEntryHash([string]$pptxPath, [string]$entryName) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($pptxPath)
    try {
        $entry = $archive.GetEntry($entryName)
        if ($null -eq $entry) { throw "Missing PPTX entry $entryName in $pptxPath" }
        $stream = $entry.Open()
        try {
            $sha = [Security.Cryptography.SHA256]::Create()
            try { return ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-', '').ToLowerInvariant() }
            finally { $sha.Dispose() }
        }
        finally { $stream.Dispose() }
    }
    finally { $archive.Dispose() }
}

function Get-SlideSignature($slide, [switch]$ExcludeMedia) {
    $items = [Collections.Generic.List[string]]::new()
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $shape = $slide.Shapes.Item($i)
        if ($ExcludeMedia -and $shape.Type -eq 16) { continue }
        $text = ''
        if ($shape.Type -ne 16) {
            try {
                if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) { $text = $shape.TextFrame.TextRange.Text }
            } catch { }
        }
        $items.Add("$($shape.Type)|$([math]::Round($shape.Left,2))|$([math]::Round($shape.Top,2))|$([math]::Round($shape.Width,2))|$([math]::Round($shape.Height,2))|$text")
    }
    return ($items -join "`n")
}

# Close only the hidden automation instance left by the preceding build, if present.
try {
    $active = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    if ($active.Presentations.Count -eq 0) { $active.Quit() }
    Release-ComObject $active
} catch { }

$ppt = $null
$source = $null
$approved = $null
$result = $null
try {
    $slide10Entries = @('ppt/slides/slide10.xml', 'ppt/slides/_rels/slide10.xml.rels', 'ppt/media/media1.mp4')
    $slide10Preserved = $true
    foreach ($entry in $slide10Entries) {
        if ((Get-PptxEntryHash $approvedDeck $entry) -ne (Get-PptxEntryHash $outputDeck $entry)) { $slide10Preserved = $false }
    }

    $ppt = New-Object -ComObject PowerPoint.Application
    $sourceSignatures = @{}
    $sourceUnderlaySignatures = @{}
    $source = $ppt.Presentations.Open($sourceDeck, $true, $false, $false)
    $sourceSlideCount = $source.Slides.Count
    foreach ($n in 1..15) { $sourceSignatures[$n] = Get-SlideSignature $source.Slides.Item($n) }
    foreach ($n in 10, 11, 12) { $sourceUnderlaySignatures[$n] = Get-SlideSignature $source.Slides.Item($n) -ExcludeMedia }
    $source.Close()
    Release-ComObject $source
    $source = $null

    $result = $ppt.Presentations.Open($outputDeck, $true, $false, $false)

    $unrelatedVerified = [Collections.Generic.List[int]]::new()
    foreach ($n in @(1..9) + @(13..15)) {
        if ($sourceSignatures[$n] -eq (Get-SlideSignature $result.Slides.Item($n))) { $unrelatedVerified.Add($n) }
    }

    $editableUnderlayVerified = [Collections.Generic.List[int]]::new()
    foreach ($n in 10, 11, 12) {
        if ($sourceUnderlaySignatures[$n] -eq (Get-SlideSignature $result.Slides.Item($n) -ExcludeMedia)) { $editableUnderlayVerified.Add($n) }
    }

    $slideReports = [ordered]@{}
    foreach ($n in 10, 11, 12) {
        $targetSlide = $result.Slides.Item($n)
        $mediaShapes = @()
        for ($i = 1; $i -le $targetSlide.Shapes.Count; $i++) {
            $candidate = $targetSlide.Shapes.Item($i)
            if ($candidate.Type -eq 16) { $mediaShapes += $candidate }
        }
        if ($mediaShapes.Count -ne 1) { throw "Expected exactly one media shape on Slide $n; found $($mediaShapes.Count)." }
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
        output_deck = $outputDeck
        reopened_in_powerpoint = $true
        source_slide_count = $sourceSlideCount
        output_slide_count = $result.Slides.Count
        slide_count_unchanged = ($sourceSlideCount -eq $result.Slides.Count -and $result.Slides.Count -eq 15)
        approved_slide_10_preserved = $slide10Preserved
        editable_underlays_verified = @($editableUnderlayVerified)
        editable_underlays_preserved = ($editableUnderlayVerified.Count -eq 3)
        unrelated_slides_verified = @($unrelatedVerified)
        unrelated_slides_unchanged = ($unrelatedVerified.Count -eq 12)
        slides = $slideReports
    }
    $report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $reportPath -Encoding UTF8
    $report | ConvertTo-Json -Depth 6

    $allPassed = $report.slide_count_unchanged -and $report.approved_slide_10_preserved -and $report.editable_underlays_preserved -and $report.unrelated_slides_unchanged
    foreach ($n in 10, 11, 12) {
        $s = $slideReports[[string]$n]
        $allPassed = $allPassed -and $s.embedded_media_shape_count -eq 1 -and -not $s.media_is_linked -and $s.media_fills_slide -and $s.starts_automatically -and $s.play_on_entry -and $s.playback_position_advanced
    }
    if (-not $allPassed) { throw "Validation failed. See $reportPath" }
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
