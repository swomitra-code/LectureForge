$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$deck = Join-Path $root 'source\Module1_Energy_Landscape_2026_Refresh_Cleaned_No_Old_Photo_FINAL2.pptx'
$work = Join-Path $root 'work'
New-Item -ItemType Directory -Force -Path $work | Out-Null

$ppt = $null
$presentation = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $presentation = $ppt.Presentations.Open($deck, $true, $false, $false)
    foreach ($slideNumber in 11, 12) {
        $png = Join-Path $work "slide-$slideNumber.png"
        $presentation.Slides.Item($slideNumber).Export($png, 'PNG', 1920, 1080)
        Write-Output $png
    }
}
finally {
    if ($presentation) { $presentation.Close() }
    if ($ppt) { $ppt.Quit() }
}
