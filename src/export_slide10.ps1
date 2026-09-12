$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$deck = Join-Path $root 'source\Module1_Energy_Landscape_2026_Refresh_Cleaned_No_Old_Photo_FINAL2.pptx'
$work = Join-Path $root 'work'
$png = Join-Path $work 'slide-10.png'
New-Item -ItemType Directory -Force -Path $work | Out-Null

$ppt = $null
$presentation = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $presentation = $ppt.Presentations.Open($deck, $true, $false, $false)
    $presentation.Slides.Item(10).Export($png, 'PNG', 1920, 1080)
    Write-Output $png
}
finally {
    if ($presentation) { $presentation.Close() }
    if ($ppt) { $ppt.Quit() }
}
