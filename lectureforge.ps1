param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$ManifestPath
)

$ErrorActionPreference = 'Stop'

function Resolve-ManifestPath([string]$manifestDirectory, [string]$projectRoot, [string]$path) {
    if ([IO.Path]::IsPathRooted($path)) { return [IO.Path]::GetFullPath($path) }
    $manifestRelative = [IO.Path]::GetFullPath((Join-Path $manifestDirectory $path))
    if (Test-Path -LiteralPath $manifestRelative) { return $manifestRelative }
    return [IO.Path]::GetFullPath((Join-Path $projectRoot $path))
}

function Get-ManifestRelativePath([string]$manifestDirectory, [string]$absolutePath) {
    $baseUri = [Uri]::new(([IO.Path]::GetFullPath($manifestDirectory).TrimEnd('\') + '\'))
    $targetUri = [Uri]::new([IO.Path]::GetFullPath($absolutePath))
    return [Uri]::UnescapeDataString($baseUri.MakeRelativeUri($targetUri).ToString())
}

function Test-AudioFile([string]$path) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $false }
    if ((Get-Item -LiteralPath $path).Length -le 0) { return $false }
    & ffmpeg -v error -i $path -f null NUL 2>$null
    if ($LASTEXITCODE -ne 0) { return $false }
    $duration = ((& ffprobe -v error -show_entries 'format=duration' -of 'default=noprint_wrappers=1:nokey=1' $path) -join '').Trim()
    return ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($duration))
}

function Test-AvatarFile([string]$path) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $false }
    if ((Get-Item -LiteralPath $path).Length -le 0) { return $false }
    & ffmpeg -v error -c:v libvpx-vp9 -i $path -f null NUL 2>$null
    if ($LASTEXITCODE -ne 0) { return $false }
    $probeText = (& ffprobe -v error -show_streams -of json $path) -join ''
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($probeText)) { return $false }
    $probe = $probeText | ConvertFrom-Json
    $video = @($probe.streams | Where-Object codec_type -eq 'video') | Select-Object -First 1
    $audio = @($probe.streams | Where-Object codec_type -eq 'audio') | Select-Object -First 1
    return ($null -ne $video -and $null -ne $audio -and [string]$video.tags.alpha_mode -eq '1')
}

function Test-CompletedDeck([object]$manifest, [string]$manifestDirectory, [string]$projectRoot, [object[]]$enabledSlides) {
    $outputDeck = Resolve-ManifestPath $manifestDirectory $projectRoot ([string]$manifest.output_pptx)
    $validationPath = [IO.Path]::ChangeExtension($outputDeck, '.validation.json')
    if (-not (Test-Path -LiteralPath $outputDeck -PathType Leaf) -or (Get-Item -LiteralPath $outputDeck).Length -le 0) { return $null }
    if (-not (Test-Path -LiteralPath $validationPath -PathType Leaf)) { return $null }
    try { $report = Get-Content -Raw -LiteralPath $validationPath | ConvertFrom-Json } catch { return $null }
    if ($report.slide_count_unchanged -ne $true -or $report.unrelated_slides_unchanged -ne $true) { return $null }

    $expected = @($enabledSlides | ForEach-Object { [int]$_.slide_number } | Sort-Object)
    $reported = @($report.requested_slides | ForEach-Object { [int]$_ } | Sort-Object)
    if (($expected -join ',') -ne ($reported -join ',')) { return $null }

    foreach ($number in $expected) {
        $slide = $report.slides.PSObject.Properties[[string]$number].Value
        $video = $report.videos.PSObject.Properties[[string]$number].Value
        if ($null -eq $slide -or $null -eq $video) { return $null }
        if ($slide.media_is_linked -ne $false -or $slide.media_fills_slide -ne $true -or
            $slide.starts_automatically -ne $true -or $slide.play_on_entry -ne $true -or
            $slide.playback_position_advanced -ne $true -or $slide.editable_underlay_preserved -ne $true -or
            $video.decode_ok -ne $true) { return $null }
        if (-not (Test-Path -LiteralPath ([string]$video.path) -PathType Leaf)) { return $null }
    }
    return [pscustomobject]@{ Deck = $outputDeck; Validation = $validationPath; Report = $report }
}

$manifestFile = (Resolve-Path -LiteralPath $ManifestPath).Path
$manifestDirectory = Split-Path -Parent $manifestFile
$projectRoot = Split-Path -Parent $manifestDirectory
$narrationTool = Join-Path $projectRoot 'generate-narration.ps1'
$avatarTool = Join-Path $projectRoot 'generate-avatar.ps1'
$buildTool = Join-Path $projectRoot 'build-deck.ps1'
foreach ($tool in $narrationTool, $avatarTool, $buildTool) {
    if (-not (Test-Path -LiteralPath $tool -PathType Leaf)) { throw "Missing proven production tool: $tool" }
}

$manifest = Get-Content -Raw -LiteralPath $manifestFile | ConvertFrom-Json
$enabledSlides = @($manifest.slides | Where-Object enabled -eq $true | Sort-Object slide_number)
if ($enabledSlides.Count -eq 0) { throw 'Manifest has no enabled production slides.' }

$slidesNeedingNarration = [Collections.Generic.List[object]]::new()
$slidesAwaitingSelection = [Collections.Generic.List[object]]::new()
foreach ($slide in $enabledSlides) {
    if ($null -eq $slide.narration -or $slide.narration.enabled -ne $true) { continue }
    if (-not [string]::IsNullOrWhiteSpace([string]$slide.narration.instructor_selected_take)) { continue }
    $validTakes = [Collections.Generic.List[string]]::new()
    foreach ($take in @($slide.narration.generated_takes)) {
        if ([string]::IsNullOrWhiteSpace([string]$take)) { continue }
        $takePath = Resolve-ManifestPath $manifestDirectory $projectRoot ([string]$take)
        if (Test-AudioFile $takePath) { $validTakes.Add($takePath) }
    }
    if ($validTakes.Count -ne 3 -and $null -ne $manifest.defaults.narration) {
        $takeDirectory = Resolve-ManifestPath $manifestDirectory $projectRoot ([string]$manifest.defaults.narration.output_dir)
        $deterministicTakes = [Collections.Generic.List[string]]::new()
        for ($takeNumber = 1; $takeNumber -le 3; $takeNumber++) {
            $takePath = Join-Path $takeDirectory "slide-$($slide.slide_number)-take-$takeNumber.mp3"
            if (Test-AudioFile $takePath) { $deterministicTakes.Add($takePath) }
        }
        if ($deterministicTakes.Count -eq 3) { $validTakes = $deterministicTakes }
    }
    if ($validTakes.Count -eq 3) {
        $slidesAwaitingSelection.Add([pscustomobject]@{ Slide = $slide; Takes = @($validTakes) })
    } else {
        $slidesNeedingNarration.Add($slide)
    }
}

if ($slidesNeedingNarration.Count -gt 0) {
    Write-Output "Narration generation required for Slide(s): $(@($slidesNeedingNarration.slide_number) -join ', ')."
    & $narrationTool $manifestFile
    if ($LASTEXITCODE -ne 0) { throw 'Narration generation failed.' }
    $manifest = Get-Content -Raw -LiteralPath $manifestFile | ConvertFrom-Json
    foreach ($number in @($slidesNeedingNarration.slide_number)) {
        $slide = $manifest.slides | Where-Object slide_number -eq $number
        Write-Output "Slide $number narration takes ready for instructor review:"
        foreach ($take in @($slide.narration.generated_takes)) { Write-Output "  $take" }
    }
    Write-Output 'Stopped at the instructor narration-selection gate. No take was selected and HeyGen was not called.'
    exit 0
}

if ($slidesAwaitingSelection.Count -gt 0) {
    foreach ($item in $slidesAwaitingSelection) {
        Write-Output "Slide $($item.Slide.slide_number) requires instructor narration selection. Valid takes:"
        foreach ($take in $item.Takes) { Write-Output "  $take" }
    }
    Write-Output 'Stopped at the instructor narration-selection gate. HeyGen was not called.'
    exit 0
}

$slidesNeedingAvatar = [Collections.Generic.List[object]]::new()
foreach ($slide in $enabledSlides) {
    if ([string]::IsNullOrWhiteSpace([string]$slide.avatar_webm)) {
        if ($null -ne $slide.avatar_generation -and $slide.avatar_generation.enabled -eq $true -and
            $null -ne $slide.narration -and -not [string]::IsNullOrWhiteSpace([string]$slide.narration.instructor_selected_take)) {
            $avatarDefaults = $manifest.defaults.avatar_generation
            $candidateDirectory = Resolve-ManifestPath $manifestDirectory $projectRoot ([string]$avatarDefaults.output_dir)
            $candidate = Join-Path $candidateDirectory "slide-$($slide.slide_number)-$($avatarDefaults.avatar_name)-production-transparent.webm"
            if (Test-AvatarFile $candidate) {
                $slide.avatar_webm = Get-ManifestRelativePath $manifestDirectory $candidate
                $manifest | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $manifestFile -Encoding UTF8
                Write-Output "Recovered existing valid avatar WebM for Slide $($slide.slide_number); no HeyGen job was created."
                continue
            }
            $slidesNeedingAvatar.Add($slide)
            continue
        }
        throw "Slide $($slide.slide_number) has no avatar WebM and is not eligible for avatar generation."
    }
    $avatarPath = Resolve-ManifestPath $manifestDirectory $projectRoot ([string]$slide.avatar_webm)
    if (-not (Test-AvatarFile $avatarPath)) { throw "Slide $($slide.slide_number) references an invalid avatar WebM: $avatarPath" }
}

if ($slidesNeedingAvatar.Count -gt 0) {
    Write-Output "Avatar generation required for Slide(s): $(@($slidesNeedingAvatar.slide_number) -join ', ')."
    & $avatarTool $manifestFile
    if ($LASTEXITCODE -ne 0) { throw 'Avatar generation failed.' }
    $manifest = Get-Content -Raw -LiteralPath $manifestFile | ConvertFrom-Json
    $enabledSlides = @($manifest.slides | Where-Object enabled -eq $true | Sort-Object slide_number)
    foreach ($number in @($slidesNeedingAvatar.slide_number)) {
        $slide = $manifest.slides | Where-Object slide_number -eq $number
        $avatarPath = Resolve-ManifestPath $manifestDirectory $projectRoot ([string]$slide.avatar_webm)
        if (-not (Test-AvatarFile $avatarPath)) { throw "Avatar validation failed after generation for Slide $number." }
        $slide.avatar_webm = Get-ManifestRelativePath $manifestDirectory $avatarPath
        Write-Output "Slide $number avatar WebM is complete and technically valid: $avatarPath"
    }
    $manifest | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $manifestFile -Encoding UTF8
}

$completed = Test-CompletedDeck $manifest $manifestDirectory $projectRoot $enabledSlides
if ($null -ne $completed) {
    Write-Output 'LectureForge is already complete for all enabled slides. No provider or deck-build tool was called.'
    Write-Output "Final PPTX: $($completed.Deck)"
    Write-Output "Validation report: $($completed.Validation)"
    Write-Output "Validated slides: $(@($completed.Report.requested_slides) -join ', ')"
    exit 0
}

Write-Output "Deck production required for enabled Slide(s): $(@($enabledSlides.slide_number) -join ', ')."
& $buildTool $manifestFile
if ($LASTEXITCODE -ne 0) { throw 'Deck production failed.' }
$manifest = Get-Content -Raw -LiteralPath $manifestFile | ConvertFrom-Json
$completed = Test-CompletedDeck $manifest $manifestDirectory $projectRoot $enabledSlides
if ($null -eq $completed) { throw 'Deck production returned without a complete validation report.' }
Write-Output "Final PPTX: $($completed.Deck)"
Write-Output "Validation report: $($completed.Validation)"
Write-Output "Validated slides: $(@($completed.Report.requested_slides) -join ', ')"
