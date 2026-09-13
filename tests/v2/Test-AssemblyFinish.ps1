param([string]$RuntimeRoot=(Join-Path $env:LOCALAPPDATA 'LectureForgeE'),[int]$Port=8788)
$ErrorActionPreference='Stop'
[Net.WebRequest]::DefaultWebProxy=$null
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Import-Module (Join-Path $repo 'studio/Assembly.psm1') -Force
$root=Get-LFRoot $RuntimeRoot
$evidence=Read-LFSharedJson (Join-Path $root 'results.json')
$id=$evidence.project_id;$dir=Get-LFProjectPath $root $id
$checks=[Collections.Generic.List[object]]::new()
function Check($Name,[bool]$OK){
 $checks.Add(@{name=$Name;passed=$OK})
 Write-LFJson (Join-Path $root 'finish-results.json') @{checks=@($checks);project_id=$id;elevenlabs_calls=0;heygen_calls=0}
 if(-not $OK){throw "FAIL: $Name"};"PASS: $Name"
}
function JobHashes {
 $q=Read-LFAvatarQueue $dir
 @(foreach($jid in $q.jobs){(Get-FileHash (Get-LFAvatarJobPath $dir $jid)).Hash}) -join ','
}
$hashes=JobHashes
$url="http://127.0.0.1:$Port/"
function Get($Path){Invoke-RestMethod ($url+$Path) -TimeoutSec 60}
$token=(Get 'api/bootstrap').token
function Post($Path,$Body){Invoke-RestMethod -Method Post -Uri ($url+$Path) -Headers @{'X-LF-Token'=$script:token} -ContentType application/json -Body ($Body|ConvertTo-Json -Depth 15) -TimeoutSec 60}
$before=Read-LFAssembly $dir
$ready=Get "api/assembly-ready?id=$id"
$queued=Post "api/assembly-create?id=$id" @{expected=$ready.binding;folder=(Join-Path $root 'Output')}
Check 'Retry uses fresh derivative and preserves failed attempt history' ($queued.current.id -ne $before.current.id -and $before.current.id -in $queued.history.id)
$deadline=[DateTime]::UtcNow.AddMinutes(20)
do {
 Start-Sleep -Seconds 3
 $s=Get "api/assembly-status?id=$id"
 "PROGRESS $($s.current.stage) slide=$($s.current.slide)"
 if($s.current.stage -eq 'Exception'){throw $s.current.error}
}while($s.current.stage -ne 'Complete' -and [DateTime]::UtcNow -lt $deadline)
Check 'Background assembly completes through responsive HTTP status' ($s.current.stage -eq 'Complete')
$v=Read-LFSharedJson (Join-Path $dir "r/$($s.current.id)/v/validation.json")
Check 'One save produces 14 slides and 14 embedded avatars' ($v.passed -and $v.save_calls -eq 1 -and $v.slide_count -eq 14 -and $v.avatar_count -eq 14)
Check 'All native structure and native renders are preserved' (@($v.slides|Where-Object {-not $_.native_preserved -or -not $_.native_render_identical}).Count -eq 0)
Check 'All avatars decode, autoplay and move/resize independently' (@($v.slides|Where-Object {-not $_.independent_move_resize -or -not $_.autoplay -or -not $_.package.embedded_full_decode_ok}).Count -eq 0)
Check 'Only Slide 6 has the placement override' ($v.slides[5].placement.left -eq .84 -and @($v.slides|Where-Object {$_.slide -ne 6 -and $_.placement.left -ne .875}).Count -eq 0)
Check 'Slide 4 embeds approved Take 2 media' ($ready.snapshot.slides[3].take -eq 2 -and $v.slides[3].package.embedded_video_sha256 -eq $ready.snapshot.slides[3].sha256)
Check 'Slide 8 binds the exact prepared narration and six-second pause' ($v.slides[7].preparation.post_speech_silence_seconds -eq 6 -and $v.slides[7].narration_sha256 -eq 'E234676D7F92B42D20E98A3CB6CADF5BAA89595C2DCA906005EB7913607D69A5')
$p=Read-LFProject $root $id
$existing=Join-Path $root 'Output/Solar Thermal_RECORDING.pptx'
Check 'Existing filename remains untouched and export uses revision suffix' ((Get-FileHash $existing).Hash -eq $p.source.sha256 -and $s.current.output -ne $existing)
Check 'Canonical source unchanged and deck reopened' ($v.reopened -and $v.source_sha256_before -eq $v.source_sha256_after -and (Get-FileHash (Join-Path $repo 'source/RE-M2-Solar-Thermal-Overview.pptx')).Hash -eq $p.source.sha256)
$stateHash=(Get-FileHash (Join-Path $dir 'assembly.json')).Hash
$deckHash=(Get-FileHash $s.current.output).Hash
$null=Post 'api/stop' @{}
for($i=0;$i -lt 60;$i++){Start-Sleep -Milliseconds 250;if((Read-LFSharedJson (Join-Path $root 'runtime.json')).status -eq 'stopped'){break}}
& (Join-Path $repo 'Start-LectureForge.ps1') -RuntimeRoot $root -Port $Port -NoBrowser
$script:token=(Get 'api/bootstrap').token
Check 'Restart preserves byte-identical assembly state and output' ((Get-FileHash (Join-Path $dir 'assembly.json')).Hash -eq $stateHash -and (Get-FileHash $s.current.output).Hash -eq $deckHash)
$opened=Post "api/recording-open?id=$id" @{kind='deck'}
Check 'Open Recording PowerPoint launches the validated output' ($opened.opened -and $opened.path -eq $s.current.output)
$opened=Post "api/recording-open?id=$id" @{kind='folder'}
Check 'Open Output Folder succeeds' $opened.opened
Check 'Avatar Ready records remain byte-identical' ((JobHashes) -eq $hashes -and (Get-LFAvatarStatus $root $id).ready -eq 14)
Check 'Zero provider calls throughout assembly and recovery' ((Get-LFNarration $root $id).provider_calls -eq 0 -and (Get-LFAvatarStatus $root $id).provider_calls -eq 0)
"RECORDING=$($s.current.output)"
