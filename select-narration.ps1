param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$ManifestPath,

    [Parameter(Mandatory = $true, Position = 1)]
    [int]$SlideNumber,

    [Parameter(Mandatory = $true, Position = 2)]
    [int]$TakeNumber
)

$ErrorActionPreference = 'Stop'

function Resolve-ManifestPath([string]$manifestDirectory, [string]$projectRoot, [string]$path) {
    if ([IO.Path]::IsPathRooted($path)) { return [IO.Path]::GetFullPath($path) }
    $manifestRelative = [IO.Path]::GetFullPath((Join-Path $manifestDirectory $path))
    if (Test-Path -LiteralPath $manifestRelative) { return $manifestRelative }
    return [IO.Path]::GetFullPath((Join-Path $projectRoot $path))
}

function Get-ProjectRelativePath([string]$projectRoot, [string]$absolutePath) {
    $baseUri = [Uri]::new(([IO.Path]::GetFullPath($projectRoot).TrimEnd('\') + '\'))
    $targetUri = [Uri]::new([IO.Path]::GetFullPath($absolutePath))
    return [Uri]::UnescapeDataString($baseUri.MakeRelativeUri($targetUri).ToString())
}

if ($TakeNumber -lt 1 -or $TakeNumber -gt 3) { throw 'Take number must be 1, 2, or 3.' }

$manifestFile = (Resolve-Path -LiteralPath $ManifestPath).Path
$manifestDirectory = Split-Path -Parent $manifestFile
$projectRoot = Split-Path -Parent $manifestDirectory
$manifestText = [IO.File]::ReadAllText($manifestFile, [Text.Encoding]::UTF8)
$manifest = $manifestText | ConvertFrom-Json

$slides = @($manifest.slides | Where-Object { [int]$_.slide_number -eq $SlideNumber })
if ($slides.Count -eq 0) { throw "Slide $SlideNumber does not exist in the manifest." }
if ($slides.Count -gt 1) { throw "Slide $SlideNumber appears more than once in the manifest." }
$slide = $slides[0]
if ($null -eq $slide.narration) { throw "Slide $SlideNumber has no narration configuration." }
if ($null -eq $manifest.defaults -or $null -eq $manifest.defaults.narration -or
    [string]::IsNullOrWhiteSpace([string]$manifest.defaults.narration.output_dir)) {
    throw 'Manifest is missing defaults.narration.output_dir.'
}

$filename = "slide-$SlideNumber-take-$TakeNumber.mp3"
$outputDirectory = Resolve-ManifestPath $manifestDirectory $projectRoot ([string]$manifest.defaults.narration.output_dir)
$takePath = Join-Path $outputDirectory $filename
if (-not (Test-Path -LiteralPath $takePath -PathType Leaf)) { throw "Generated narration take does not exist: $takePath" }
if ((Get-Item -LiteralPath $takePath).Length -le 0) { throw "Generated narration take is empty: $takePath" }

$recordedTake = $false
foreach ($configuredTake in @($slide.narration.generated_takes)) {
    if ([string]::IsNullOrWhiteSpace([string]$configuredTake)) { continue }
    $configuredPath = Resolve-ManifestPath $manifestDirectory $projectRoot ([string]$configuredTake)
    if ([string]::Equals($configuredPath, [IO.Path]::GetFullPath($takePath), [StringComparison]::OrdinalIgnoreCase)) {
        $recordedTake = $true
        break
    }
}
if (-not $recordedTake) { throw "Slide $SlideNumber Take $TakeNumber is not recorded in narration.generated_takes." }

& ffmpeg -v error -i $takePath -f null NUL
if ($LASTEXITCODE -ne 0) { throw "Generated narration take does not decode: $takePath" }

$manifestTakePath = Get-ProjectRelativePath $projectRoot $takePath
$currentSelection = [string]$slide.narration.instructor_selected_take
if (-not [string]::Equals($currentSelection, $manifestTakePath, [StringComparison]::Ordinal)) {
    $numberPattern = [regex]::Escape([string]$SlideNumber)
    $pattern = '(?s)(\{\s*"slide_number"\s*:\s*' + $numberPattern + '\s*,.*?"narration"\s*:\s*\{.*?"instructor_selected_take"\s*:\s*)("(?:\\.|[^"\\])*"|null)'
    $selectionRegex = [regex]::new($pattern)
    $match = $selectionRegex.Match($manifestText)
    if (-not $match.Success) { throw "Could not locate Slide $SlideNumber instructor_selected_take in the manifest text." }
    $secondMatch = $selectionRegex.Match($manifestText, $match.Index + $match.Length)
    if ($secondMatch.Success) { throw "Slide $SlideNumber selection field is ambiguous in the manifest text." }
    $jsonValue = $manifestTakePath | ConvertTo-Json -Compress
    $updatedText = $manifestText.Substring(0, $match.Groups[2].Index) + $jsonValue + $manifestText.Substring($match.Groups[2].Index + $match.Groups[2].Length)

    $bytes = [IO.File]::ReadAllBytes($manifestFile)
    $hasUtf8Bom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    $encoding = [Text.UTF8Encoding]::new($hasUtf8Bom)
    [IO.File]::WriteAllText($manifestFile, $updatedText, $encoding)
}

Write-Output "Selected Slide $SlideNumber Take ${TakeNumber}: $manifestTakePath"
Write-Output "Validated narration file: $takePath"
Write-Output 'No provider, avatar, PowerPoint, or LectureForge orchestration tool was called.'
