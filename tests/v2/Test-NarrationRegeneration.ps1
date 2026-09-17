param([string]$RuntimeRoot=(Join-Path $PSScriptRoot '../../work/take-regeneration-tests'))
$ErrorActionPreference='Stop'
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Import-Module (Join-Path $repo 'studio/Narration.psm1') -Force
$provider=Import-Module (Join-Path $repo 'studio/NarrationProvider.psm1') -Force -PassThru
$root=Join-Path $RuntimeRoot ([guid]::NewGuid().ToString('N').Substring(0,8))
[void][IO.Directory]::CreateDirectory($root)
function Check($Name,$OK){if(-not $OK){throw "FAIL: $Name"};"PASS: $Name"}
function Reject($Name,[scriptblock]$Action){$rejected=$false;try{& $Action}catch{$rejected=$true};Check $Name $rejected}
# Minimal two-slide deck: no PowerPoint or live services required.
Add-Type -AssemblyName System.IO.Compression.FileSystem
$source=Join-Path $root 'input.pptx';$zip=[IO.Compression.ZipFile]::Open($source,'Create')
function Part($Name,$Text){$w=[IO.StreamWriter]::new($zip.CreateEntry($Name).Open());try{$w.Write($Text)}finally{$w.Dispose()}}
try{
 Part 'ppt/presentation.xml' '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>'
 Part 'ppt/_rels/presentation.xml.rels' '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/><Relationship Id="rId2" Target="slides/slide2.xml"/></Relationships>'
 Part 'ppt/slides/slide1.xml' '<slide><t>First</t></slide>'
 Part 'ppt/slides/slide2.xml' '<slide><t>Second</t></slide>'
}finally{$zip.Dispose()}
$fixture=Join-Path $root 'fixture.mp3'
& ffmpeg -v error -f lavfi -i 'sine=frequency=440:duration=0.2' -c:a libmp3lame $fixture
if($LASTEXITCODE){throw 'Fixture audio failed'}
$fixtureAsset=Get-LFAudioValidation $fixture
$firstId=$null
# Intercept only transport in the real provider; execute the real worker body
# with its already-loaded modules so -Force does not erase the offline transport.
& $provider {param($Fixture)
 $script:fixture=$Fixture;$script:requests=0;$script:heygen=0;$script:fail=$false;$script:payload=$null
 function script:Invoke-WebRequest {
  param([switch]$UseBasicParsing,$TimeoutSec,$Method,$Uri,$Headers,$ContentType,$Body,$OutFile)
  if($Uri -notlike 'https://api.elevenlabs.io/v1/text-to-speech/*'){$script:heygen++;throw 'Unexpected provider'}
  $script:requests++;$script:payload=[Text.Encoding]::UTF8.GetString($Body)|ConvertFrom-Json
  if($script:fail){throw 'Offline transport failure'}
  [IO.File]::Copy($script:fixture,$OutFile,$false)
 }
} $fixture
$workerText=[IO.File]::ReadAllText((Join-Path $repo 'studio/Invoke-NarrationTake.ps1'))
$worker=[scriptblock]::Create(($workerText -replace '(?m)^Import-Module[^\r\n]*',''))
$priorKey=$env:ELEVENLABS_API_KEY
try{
 $env:ELEVENLABS_API_KEY='offline-fixture'
 foreach($take in 1..3){foreach($selected in @($take,($take%3+1))){
  $p=New-LFProject $root 'Regeneration' 'input.pptx' ([IO.File]::ReadAllBytes($source))
  if(-not $firstId){$firstId=$p.id}
  foreach($s in $p.slides){$s.enabled=$true;$s.script="Saved [POINT: collector]narration $($s.number). [PAUSE 2, REVEAL]Read [kWh]."}
  $p=Update-LFProject $root $p.id ([pscustomobject]@{version=$p.version;slides=$p.slides})
  $dir=Get-LFProjectPath $root $p.id
  Add-LFNarrationBatch $root $p.id $p.version
  $state=Read-LFNarrationFile $dir
  foreach($r in $state.revisions){foreach($t in $r.takes){
   $a=$t.attempts[-1];$relative=$a.raw_path.Replace('raw.mp3','take.mp3');$path=Resolve-LFNarrationAsset $dir $relative
   [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path));[IO.File]::Copy($fixture,$path,$false)
   $t.asset=@{path=$relative;sha256=$fixtureAsset.sha256;duration_seconds=$fixtureAsset.duration_seconds;decode_ok=$true;post_speech_silence_seconds=0};$t.state='ready';$a.state='complete'
  }}
  $r=$state.revisions[0];$r.selected_take=$selected
  Write-LFJson (Join-Path $dir 'narration.json') $state
  $others=$r.takes|Where-Object number -ne $take|ConvertTo-Json -Depth 30 -Compress
  $otherSlide=$state.revisions[1]|ConvertTo-Json -Depth 30 -Compress
  $projectHash=(Get-FileHash (Join-Path $dir 'project.json')).Hash
  $t=$r.takes[$take-1];$old=$t.asset;$expected=$t.attempts[-1].id
  $calls=& $provider {$script:requests}
  Add-LFNarrationRegeneration $root $p.id $r.id $take $expected
  Reject 'Duplicate submission rejected while queued' {Add-LFNarrationRegeneration $root $p.id $r.id $take $expected}
  $task=Get-LFNextNarration $root
  Check 'Exactly the requested slot dispatched' ($task.id -eq $p.id -and $task.take -eq $take)
  & $worker $root $task.id $task.revision $task.take $task.attempt
  $after=Read-LFNarrationFile $dir;$current=$after.revisions[0];$new=$current.takes[$take-1]
  Check "Take $take replaced and verified (selected $selected)" ($new.state -eq 'ready' -and $new.asset.path -ne $old.path -and (Test-LFTake $dir $new))
  Check 'Other takes path and metadata identical' (($current.takes|Where-Object number -ne $take|ConvertTo-Json -Depth 30 -Compress) -ceq $others)
  Check 'Other takes byte-identical' (@($current.takes|Where-Object number -ne $take|Where-Object {-not (Test-LFTake $dir $_)}).Count -eq 0)
  Check 'Other slide unchanged' (($after.revisions[1]|ConvertTo-Json -Depth 30 -Compress) -ceq $otherSlide)
  Check 'Selected slot preserved' ($current.selected_take -eq $selected)
  Check 'Exactly one provider request recorded and sent' ((Get-LFNarration $root $p.id).provider_calls -eq 1 -and (& $provider {$script:requests}) -eq $calls+1)
  Check 'Current saved script sanitized by existing provider' ((& $provider {$script:payload.text}) -ceq 'Saved  narration 1.  Read [kWh].')
  Check 'Voice preset unchanged' ((& $provider {$script:payload.voice_settings.speed}) -eq $p.preset.narration.speed -and (& $provider {$script:payload.model_id}) -eq $p.preset.narration.model_id)
  $raw=Resolve-LFNarrationAsset $dir $new.attempts[-1].raw_path
  Check 'Exact spoken-text sidecar preserved' ([IO.File]::ReadAllText($raw+'.spoken.txt') -ceq (& $provider {$script:payload.text}))
  Check 'Previous valid audio retained with history' ((Get-FileHash (Resolve-LFNarrationAsset $dir $old.path)).Hash -eq $old.sha256 -and $new.attempts[-1].previous_asset.path -eq $old.path)
  Check 'Project file unchanged' ((Get-FileHash (Join-Path $dir 'project.json')).Hash -eq $projectHash)
  Reject 'Delayed duplicate rejected after completion' {Add-LFNarrationRegeneration $root $p.id $r.id $take $expected}
 }}
 # Failure and recovery on last project's selected slot.
 $take=$selected;$current=(Read-LFNarrationFile $dir).revisions[0];$old=$current.takes[$take-1].asset
 & $provider {$script:fail=$true}
 Add-LFNarrationRegeneration $root $p.id $current.id $take $current.takes[$take-1].attempts[-1].id
 $task=Get-LFNextNarration $root
 & $worker $root $task.id $task.revision $task.take $task.attempt
 $failed=(Read-LFNarrationFile $dir).revisions[0].takes[$take-1]
 Check 'Failure retains previous valid file and path' ($failed.state -eq 'uncertain' -and $failed.asset.path -eq $old.path -and (Get-FileHash (Resolve-LFNarrationAsset $dir $old.path)).Hash -eq $old.sha256)
 $calls=& $provider {$script:requests}
 Restore-LFNarrationTake $root $p.id $current.id $take
 $restored=(Read-LFNarrationFile $dir).revisions[0]
 Check 'Recovery restores playback and selected slot without request' ((Test-LFTake $dir $restored.takes[$take-1]) -and $restored.selected_take -eq $take -and (& $provider {$script:requests}) -eq $calls)
 Reject 'Uncertain regeneration cannot be resubmitted after recovery' {Add-LFNarrationRegeneration $root $p.id $current.id $take $task.attempt}
 $p.slides[0].script='Changed saved narration'
 $p=Update-LFProject $root $p.id ([pscustomobject]@{version=$p.version;slides=$p.slides})
 Reject 'Obsolete card cannot send outdated narration' {Add-LFNarrationRegeneration $root $p.id $current.id 3 $restored.takes[2].attempts[-1].id}
 Check 'No HeyGen request or avatar state created' ((& $provider {$script:heygen}) -eq 0 -and -not (Test-Path (Join-Path $dir 'avatars.json')) -and $workerText -notmatch 'Invoke-LFHeyGen')
 Write-LFJson (Join-Path $root 'ui-fixture.json') @{root=$root;id=$firstId}
 & node (Join-Path $PSScriptRoot 'Test-NarrationRegeneration.cjs') (Join-Path $root 'ui-fixture.json')
 if($LASTEXITCODE){throw 'Regeneration UI/HTTP regression failed'}
 'PASS: per-take regeneration regressions; transport mocked; live paid requests 0.'
}finally{$env:ELEVENLABS_API_KEY=$priorKey;Remove-Module $provider}
