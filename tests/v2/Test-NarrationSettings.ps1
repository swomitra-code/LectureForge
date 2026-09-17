param([string]$RuntimeRoot=(Join-Path $PSScriptRoot '../../work/narration-settings-tests'))
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
  $script:uri=$Uri
  if($script:fail){throw 'Offline transport failure'}
  [IO.File]::Copy($script:fixture,$OutFile,$false)
 }
} $fixture
$workerText=[IO.File]::ReadAllText((Join-Path $repo 'studio/Invoke-NarrationTake.ps1'))
$worker=[scriptblock]::Create(($workerText -replace '(?m)^Import-Module[^\r\n]*',''))
Import-Module (Join-Path $repo 'studio/NarrationSettings.psm1') -Force
Import-Module (Join-Path $repo 'studio/Authorization.psm1') -Force
$priorKey=$env:ELEVENLABS_API_KEY
try{
 $env:ELEVENLABS_API_KEY='offline-fixture'
 $p=New-LFProject $root 'Settings' 'input.pptx' ([IO.File]::ReadAllBytes($source));$id=$p.id;$dir=Get-LFProjectPath $root $id
 $p.preset.PSObject.Properties.Remove('narration');Write-LFJson (Join-Path $dir 'project.json') $p
 $oldHash=(Get-FileHash (Join-Path $dir 'project.json')).Hash
 $p=Read-LFProject $root $id;$defaults=Get-LFNarrationDefaults
 Check 'Legacy lecture missing settings reads exact defaults without migration write' ((ConvertTo-LFCanonicalJson $p.preset.narration) -eq (ConvertTo-LFCanonicalJson $defaults) -and (Get-FileHash (Join-Path $dir 'project.json')).Hash -eq $oldHash)
 $p.preset.narration.PSObject.Properties.Remove('takes_per_slide');Write-LFJson (Join-Path $dir 'project.json') $p
 $p=Read-LFProject $root $id
 Check 'Partial legacy settings also fill missing fields' ($p.preset.narration.takes_per_slide -eq 3)
 $p.slides[0].script='Saved [POINT: collector]narration.';$p.slides[1].enabled=$false
 $p=Update-LFProject $root $id ([pscustomobject]@{version=$p.version;slides=$p.slides})
 Check 'Default generation plan remains three requests' ((Get-LFNarrationPlan $root $id).generations -eq 3)
 Add-LFNarrationBatch $root $id $p.version
 for($i=0;$i -lt 3;$i++){$task=Get-LFNextNarration $root;& $worker $root $task.id $task.revision $task.take $task.attempt}
 $original=(Read-LFNarrationFile $dir).revisions[0]
 Check 'Existing generation works with default settings' (@($original.takes|Where-Object state -eq 'ready').Count -eq 3 -and (& $provider {$script:requests}) -eq 3)
 Set-LFNarrationSelection $root $id 1 $original.id 2
 $binding=Get-LFBinding $p $p.slides[0]
 # UI saves of unchanged settings must retain legacy narration binding.
 $p=Update-LFProject $root $id ([pscustomobject]@{version=$p.version;slides=$p.slides;narration_settings=$defaults})
 Check 'Default save preserves existing binding and selected take' ((Get-LFBinding $p $p.slides[0]) -eq $binding -and (Get-LFNarration $root $id).slides[0].selected_take -eq 2)
 $beforeProduction=ConvertTo-LFCanonicalJson (Read-LFNarrationFile $dir).revisions
 $approvalBefore=(Get-LFSelectionInputs $root $id $dir).binding_sha256
 $projectBefore=(Get-FileHash (Join-Path $dir 'project.json')).Hash
 $custom=Get-LFNarrationDefaults;$custom.voice_id='custom-voice';$custom.model_id='custom-model';$custom.stability=.47;$custom.style=.38;$custom.speed=1.1;$custom.use_speaker_boost=$false;$custom.takes_per_slide=5
 $calls=& $provider {$script:requests}
 $preview=Add-LFNarrationPreview $root $id 1 $p.version $custom 'Unsaved [REVEAL]test [kWh].'
 Reject 'Duplicate pending preview rejected' {Add-LFNarrationPreview $root $id 1 $p.version $custom 'Duplicate'}
 $task=Get-LFNextNarration $root;& $worker $root $task.id $task.revision $task.take $task.attempt
 $state=Read-LFNarrationFile $dir;$pr=$state.revisions[-1];$payload=& $provider {$script:payload}
 Check 'Test take sends exactly one request despite five configured takes' ((& $provider {$script:requests}) -eq $calls+1 -and $pr.takes.Count -eq 1 -and $pr.takes[0].state -eq 'ready' -and $null -eq (Get-LFNextNarration $root))
 Check 'Custom settings reach existing ElevenLabs payload' ($payload.model_id -eq 'custom-model' -and $payload.voice_settings.stability -eq .47 -and $payload.voice_settings.style -eq .38 -and $payload.voice_settings.speed -eq 1.1 -and $payload.voice_settings.use_speaker_boost -eq $false)
 Check 'Custom voice reaches existing provider endpoint' ((& $provider {$script:uri}) -eq 'https://api.elevenlabs.io/v1/text-to-speech/custom-voice?output_format=mp3_44100_128')
 Check 'Preview uses sanitizer and exact sidecar' ($payload.text -ceq 'Unsaved  test [kWh].' -and [IO.File]::ReadAllText((Resolve-LFNarrationAsset $dir $pr.takes[0].attempts[0].raw_path)+'.spoken.txt') -ceq $payload.text)
 Check 'Preview preserves project, production takes, selection, and authorization binding' ((Get-FileHash (Join-Path $dir 'project.json')).Hash -eq $projectBefore -and (ConvertTo-LFCanonicalJson @($state.revisions|Where-Object {$_.preview -ne $true})) -eq $beforeProduction -and (Get-LFSelectionInputs $root $id $dir).binding_sha256 -eq $approvalBefore)
 Reject 'Preview cannot be selected' {Set-LFNarrationSelection $root $id 1 $preview 1}
 Check 'Preview makes zero HeyGen requests' ((& $provider {$script:heygen}) -eq 0 -and -not (Test-Path (Join-Path $dir 'avatar-queue.json')) -and $workerText -notmatch 'HeyGenProvider|Invoke-LFHeyGen')
 foreach($count in 1,5){
  $custom.takes_per_slide=$count
  $p=Update-LFProject $root $id ([pscustomobject]@{version=$p.version;slides=$p.slides;narration_settings=$custom})
  $p=Read-LFProject $root $id
  Check "Settings save and reopen with $count takes; project schema unchanged" ($p.preset.narration.takes_per_slide -eq $count -and $p.preset.narration.style -eq .38 -and $p.schema -eq 'lectureforge-studio-1')
  Check "Plan requests $count takes" ((Get-LFNarrationPlan $root $id).generations -eq $count)
  $calls=& $provider {$script:requests};Add-LFNarrationBatch $root $id $p.version
  for($i=0;$i -lt $count;$i++){$task=Get-LFNextNarration $root;& $worker $root $task.id $task.revision $task.take $task.attempt}
  Check "Executor sends $count requests and preserves all slots on read" ((& $provider {$script:requests}) -eq $calls+$count -and (Read-LFNarrationFile $dir).revisions[-1].takes.Count -eq $count)
  Set-LFNarrationSelection $root $id 1 (Read-LFNarrationFile $dir).revisions[-1].id $count
  Check "Take $count can be selected" ((Get-LFNarration $root $id).slides[0].selected_take -eq $count)
 }
 $p=Update-LFProject $root $id ([pscustomobject]@{version=$p.version;slides=$p.slides;narration_settings=(Get-LFNarrationDefaults)})
 Check 'Restore defaults saves exact preset and recovers original selection' ((ConvertTo-LFCanonicalJson $p.preset.narration) -eq (ConvertTo-LFCanonicalJson $defaults) -and (Get-LFNarration $root $id).slides[0].selected_take -eq 2)
 foreach($case in @(@('stability',-0.1),@('similarity_boost',1.01),@('style',[double]::NaN),@('speed',.69),@('speed',1.21),@('takes_per_slide',0),@('takes_per_slide',6),@('takes_per_slide',2.5),@('style','0.5'),@('stability',$null),@('use_speaker_boost','true'))){
  $bad=Get-LFNarrationDefaults;$bad.($case[0])=$case[1]
  Reject "Backend rejects invalid $($case[0]) = $($case[1])" {Update-LFProject $root $id ([pscustomobject]@{version=$p.version;slides=$p.slides;narration_settings=$bad})}
  Reject 'Preview also rejects invalid settings' {Add-LFNarrationPreview $root $id 1 $p.version $bad 'Test'}
  Reject 'Provider rejects invalid settings before transport' {Invoke-LFElevenLabs $bad 'Test' (Join-Path $root 'invalid.mp3')}
 }
 Write-LFJson (Join-Path $root 'ui-fixture.json') @{root=$root;id=$id;preview=$preview}
 & node (Join-Path $PSScriptRoot 'Test-NarrationSettings.cjs') (Join-Path $root 'ui-fixture.json')
 if($LASTEXITCODE){throw 'Settings HTTP/UI regression failed'}
 'PASS: narration settings regression; real paid requests 0.'
}finally{$env:ELEVENLABS_API_KEY=$priorKey;Remove-Module $provider}
