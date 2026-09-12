$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$sourceDeck = Join-Path $root 'source\Module1_Energy_Landscape_2026_Refresh_Cleaned_No_Old_Photo_FINAL2.pptx'
$outputDeck = Join-Path $root 'output\Module1_Energy_Landscape_VIDEO_SLIDES_10_12.pptx'
$reportPath = Join-Path $root 'output\validation-unrelated-slides.json'

function Release-ComObject($object) {
    if ($null -ne $object) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object) }
}

function Get-SlideSignature($slide) {
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

$ppt = $null
$source = $null
$result = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $source = $ppt.Presentations.Open($sourceDeck, $true, $false, $false)
    $result = $ppt.Presentations.Open($outputDeck, $true, $false, $false)
    $verified = [Collections.Generic.List[int]]::new()
    foreach ($n in @(1..9) + @(13..15)) {
        if ((Get-SlideSignature $source.Slides.Item($n)) -eq (Get-SlideSignature $result.Slides.Item($n))) { $verified.Add($n) }
    }
    $report = [ordered]@{
        unrelated_slides_verified = @($verified)
        unrelated_slides_unchanged = ($verified.Count -eq 12)
    }
    $report | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $reportPath -Encoding UTF8
    $report | ConvertTo-Json -Depth 3
    if (-not $report.unrelated_slides_unchanged) { throw "Unrelated slide validation failed. See $reportPath" }
}
finally {
    if ($result) { try { $result.Close() } catch { } }
    if ($source) { try { $source.Close() } catch { } }
    if ($ppt) { try { $ppt.Quit() } catch { } }
    Release-ComObject $result
    Release-ComObject $source
    Release-ComObject $ppt
}
