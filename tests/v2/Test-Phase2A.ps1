param([string]$ProjectPath=(Join-Path $PSScriptRoot '../../work/v2-phase2a/solar-thermal/project.json'),[string]$BaseCandidateName='slide_01_candidate_001')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot '../../studio/Project.psm1') -Force
Import-Module (Join-Path $PSScriptRoot '../../studio/Render.psm1') -Force
$results=[Collections.Generic.List[object]]::new()
function Assert-Proof([string]$Name,[bool]$Passed) {
    $results.Add([pscustomobject]@{test=$Name;passed=$Passed})
    if (-not $Passed) { throw "FAILED: $Name" }
}
function Expect-Rejection([string]$Name,[scriptblock]$Action) {
    $rejected=$false
    try { & $Action | Out-Null } catch { $rejected=$true }
    Assert-Proof $Name $rejected
}
$project=Read-V2ProofProject $ProjectPath
$root=$project.root
$record=$project.record
$sourceReport=Get-Content -Raw (Join-Path $root 'validation/source-render.json') | ConvertFrom-Json
$first=Get-Content -Raw (Join-Path $root "validation/$BaseCandidateName.json") | ConvertFrom-Json
$originalBefore=(Get-V2Asset $record.source.original_path).sha256
$importedBefore=(Get-V2Asset (Join-Path $root $record.source.path)).sha256
$renderBefore=(Get-V2Asset (Join-Path $root 'renders/slide_01.png')).sha256
$projectBefore=(Get-V2Asset $ProjectPath).sha256
Assert-Proof 'Source imported byte-for-byte and unchanged across COM rendering' ($sourceReport.original_sha256_before -eq $sourceReport.original_sha256_after -and $sourceReport.imported_sha256_before -eq $sourceReport.imported_sha256_after -and $originalBefore -eq $importedBefore -and $originalBefore -eq $record.source.sha256)
Assert-Proof 'Source remains 14 slides and 16:9' ((Get-V2DeckInfo (Join-Path $root $record.source.path)).is_16_9 -and $sourceReport.slide_count_before -eq 14 -and $sourceReport.slide_count_after -eq 14)

# Tamper only a test manifest, never any binary source or fixture.
$invalid=Get-Content -Raw $ProjectPath | ConvertFrom-Json
$invalid.source.sha256='INVALID_TEST_HASH'
$invalidPath=Join-Path $root 'project-invalid-hash-test.json'
Write-V2Json $invalidPath $invalid
Expect-Rejection 'Reject a source that does not match recorded hash' {Read-V2ProofProject $invalidPath}

# Phase 2A forbids all save/embed operations. The separate Recording module is
# deliberately outside this scope and may save/embed only into a derivative.
$codeFiles=@(Get-ChildItem (Join-Path $PSScriptRoot '../../studio') -File | Where-Object Name -in 'Project.psm1','Render.psm1','Invoke-Phase2A.ps1')
$forbidden='(?i)Invoke-WebRequest|Invoke-RestMethod|Start-BitsTransfer|curl(?:\.exe)?|HttpClient|WebRequest|generate-avatar\.ps1|generate-narration\.ps1|AddMediaObject|\.Save(?:As|CopyAs)?\s*\('
foreach($file in $codeFiles) {
    $tokens=$null;$errors=$null
    [void][Management.Automation.Language.Parser]::ParseFile($file.FullName,[ref]$tokens,[ref]$errors)
    Assert-Proof ("PowerShell syntax: "+$file.Name) ($errors.Count -eq 0)
    Assert-Proof ("No network/provider or PPTX save/embed code: "+$file.Name) (-not [regex]::IsMatch([IO.File]::ReadAllText($file.FullName),$forbidden))
}
$maps=@()
for($i=0;$i -lt $first.ffmpeg_arguments.Count-1;$i++){if($first.ffmpeg_arguments[$i] -eq '-map'){$maps+=$first.ffmpeg_arguments[$i+1]}}
Assert-Proof 'Only filtered video and explicit narration input are mapped' (($maps -join ',') -eq '[v],2:a:0' -and $first.ffmpeg_arguments -notcontains '-shortest')
Assert-Proof 'Actual AAC packets equal independent narration encode' ($first.actual_audio_packet_hash -eq $first.expected_audio_packet_hash -and $first.audio_routing_verified)
Assert-Proof 'Transparent and opaque pixels actually decode' $first.decoded_alpha_has_zero_and_255
$probe=Get-V2Probe $first.candidate
Assert-V2VideoProfile $probe $first.narration_duration
Assert-Proof 'Candidate profile, duration and full decode passed' ($first.output_profile_valid -and $first.full_decode_ok)
$bad=$probe | ConvertTo-Json -Depth 16 | ConvertFrom-Json
@($bad.streams | Where-Object codec_type -eq 'video')[0].width=1280
Expect-Rejection 'Reject wrong output dimensions' {Assert-V2VideoProfile $bad $first.narration_duration}
Expect-Rejection 'Reject output duration mismatch' {Assert-V2VideoProfile $probe ($first.narration_duration+2)}
Expect-Rejection 'Reject off-slide placement' {Get-V2Placement @{x=0.99;y=0.7;width=0.2} $record.avatar_crop}

# A real second full composition with the same project/assets and changed placement.
# No COM render or provider entry points are involved.
$name='slide_01_placement_test_'+[DateTime]::UtcNow.ToString('yyyyMMddHHmmss')
$second=Build-V2Slide $ProjectPath $name @{x=0.82;y=0.68;width=0.1458333333333333;approved=$false}
Assert-Proof 'Placement-only recomposition keeps all source assets' ($second.narration_sha256 -eq $first.narration_sha256 -and $second.avatar_sha256 -eq $first.avatar_sha256 -and $second.render_sha256 -eq $first.render_sha256)
Assert-Proof 'Placement changes video pixels without changing final audio' ($second.candidate_sha256 -ne $first.candidate_sha256 -and $second.actual_audio_packet_hash -eq $first.actual_audio_packet_hash)
Assert-Proof 'Recomposition invokes no providers or COM rendering' ($second.provider_calls -eq 0 -and $second.render_calls -eq 0)
Assert-Proof 'Recomposition preserves source, render and project metadata' ((Get-V2Asset $record.source.original_path).sha256 -eq $originalBefore -and (Get-V2Asset (Join-Path $root $record.source.path)).sha256 -eq $importedBefore -and (Get-V2Asset (Join-Path $root 'renders/slide_01.png')).sha256 -eq $renderBefore -and (Get-V2Asset $ProjectPath).sha256 -eq $projectBefore)
Assert-Proof 'Provisional placement was not promoted to project default' (-not $record.project_placement.approved -and $null -eq $record.project_placement.width)

# Check real background pixels through transparent parts of the presenter's box.
# H.264 is lossy, so allow a 20-level per-channel difference from the PNG.
Add-Type -AssemblyName System.Drawing
$background=[Drawing.Bitmap]::new((Join-Path $root 'renders/slide_01.png'))
$composed=[Drawing.Bitmap]::new($first.review_frame)
$alpha=[IO.File]::ReadAllBytes((Join-Path $root "validation/$BaseCandidateName-alpha.gray"))
$samples=0;$maxDelta=0
try {
    $c=$record.avatar_crop;$p=$first.placement_pixels
    foreach($cy in @(20,60,100,160,240,360,480 | Where-Object {$_ -lt $record.avatar_crop.height-5})) {
        foreach($cx in 20,60,120,520,580,620) {
            $clear=$true
            foreach($dy in -5,0,5){foreach($dx in -5,0,5){if($alpha[($cy+$dy)*$c.width+$cx+$dx] -ne 0){$clear=$false}}}
            if(-not $clear){continue}
            $px=$p.x+[int][math]::Round($cx*$p.width/$c.width)
            $py=$p.y+[int][math]::Round($cy*$p.height/$c.height)
            $a=$background.GetPixel($px,$py);$b=$composed.GetPixel($px,$py)
            foreach($channel in 'R','G','B'){$maxDelta=[math]::Max($maxDelta,[math]::Abs([int]$a.$channel-[int]$b.$channel))}
            $samples++
        }
    }
} finally {$background.Dispose();$composed.Dispose()}
Assert-Proof 'Slide background remains visible through avatar alpha' ($samples -ge 4 -and $maxDelta -le 20)

$repo=(Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$oldFiles=@(& git -C $repo ls-tree -r --name-only lectureforge-v1-baseline)
& git -C $repo diff --quiet lectureforge-v1-baseline -- $oldFiles
Assert-Proof 'All tracked V1 files remain unchanged' ($LASTEXITCODE -eq 0)
$report=[ordered]@{passed=$true;tests=@($results);count=$results.Count;provider_calls=0;alpha_background_samples=$samples;alpha_background_max_channel_delta=$maxDelta;recomposition_candidate=$second.candidate;limitations=@('No lip-sync acceptance','Energy Management technical fixtures, not Solar Thermal narration','Instructor placement approval pending')}
Write-V2Json (Join-Path $root 'validation/phase2a-tests.json') $report
$report | ConvertTo-Json -Depth 6
