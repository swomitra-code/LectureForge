param([string]$RuntimeRoot=(Join-Path $env:LOCALAPPDATA 'LectureForgeE'),[int]$Port=8788)
$ErrorActionPreference='Stop';[Net.WebRequest]::DefaultWebProxy=$null
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Import-Module (Join-Path $repo 'studio/Assembly.psm1') -Force
$root=Get-LFRoot $RuntimeRoot
if(Test-Path (Join-Path $root 'Projects')){throw 'Use a fresh isolated short test root.'}
$priorRoot=Join-Path $env:LOCALAPPDATA 'LectureForgeDF';$prior=Read-LFSharedJson (Join-Path $priorRoot 'results.json');$id=$prior.project_id
$dir=Get-LFProjectPath $root $id;[void][IO.Directory]::CreateDirectory((Split-Path $dir))
Copy-Item -LiteralPath $prior.project_root -Destination $dir -Recurse
Write-LFJson (Join-Path $root 'avatar-policy.json') @{provider_calls_enabled=$false}
$p=Read-LFProject $root $id;$p.name='Solar Thermal';$p.slides[0].enabled=$true
$slide1=Join-Path $repo 'work/solar-thermal-slide1-production-r002'
$p.slides[0].script=[IO.File]::ReadAllText((Join-Path $slide1 'slide-1-approved-script.txt')).TrimEnd([char[]]"`r`n");$p.version++
Write-LFJson (Join-Path $dir 'project.json') $p
# Create only fixture metadata and copy the existing three Slide 1 takes. No generator runs.
Add-LFNarrationBatch $root $id $p.version
$n=Read-LFNarrationFile $dir;$r=Get-LFCurrentRevision $n $p $p.slides[0]
foreach($t in $r.takes){$a=$t.attempts[-1];$rel=$a.raw_path.Replace('raw.mp3','take.mp3');$dest=Resolve-LFNarrationAsset $dir $rel;[void][IO.Directory]::CreateDirectory((Split-Path $dest));Copy-Item -LiteralPath (Join-Path $slide1 "takes/slide-1-take-$($t.number).mp3") -Destination $dest;$asset=Get-LFAudioValidation $dest;$asset.path=$rel;$asset.pause_samples=0;$asset.pause_sample_rate=44100;$asset.post_speech_silence_seconds=0;$t.asset=$asset;$t.state='ready';$a.state='fixture';$a.finished_utc=[DateTime]::UtcNow.ToString('o')}
$r.selected_take=2;Write-LFJson (Join-Path $dir 'narration.json') $n
$catalog=Read-LFSharedJson (Join-Path $dir 'avatars.json');Copy-Item -LiteralPath (Join-Path $slide1 'slide-1-avatar-white.mp4') -Destination (Join-Path $dir 'avatars/s1.mp4')
$c=Get-LFSelectionInputs $root $id $dir
$catalog.assets+=@{id='solar-s1';slide_number=1;status='approved';rejected=$false;video_id='2ec644fbfe94f44e0068dfcf89224e88';narration_sha256=$r.takes[1].asset.sha256;preset_sha256=$c.snapshot.avatar_preset_sha256;path='avatars/s1.mp4';sha256=(Get-FileHash (Join-Path $dir 'avatars/s1.mp4')).Hash;validation=@{passed=$true;authoritative_audio_verified=$true}}
Write-LFJson (Join-Path $dir 'avatars.json') $catalog
$old=Get-LFAvatarStatus $root $id
Set-LFAvatarPause $root $id $false;$c=Get-LFSelectionReview $root $id;$null=Save-LFSelectionReview $root $id $c.current.binding_sha256 confirmation;$null=Confirm-LFAvatarAuthorization $root $id $c.current.binding_sha256 $c.current.binding_sha256
Start-LFAvatarBatch $root $id;$q=Read-LFAvatarQueue $dir;$jobs=$q.batches[-1].jobs
foreach($jid in $jobs){$j=Read-LFAvatarJob $dir $jid;$j.output_path='avatars/s'+$j.slide+'.mp4';$j.output_sha256=(Get-FileHash (Join-Path $dir $j.output_path)).Hash
 if($j.slide -eq 1){$j.stage='Validating';$j.lease=@{owner='fixture';pid=$PID;started=(Get-Process -Id $PID).StartTime.ToUniversalTime().Ticks.ToString()}}
 else{$prev=$old.jobs|Where-Object slide -eq $j.slide;$j.validation=$prev.validation;$j.stage='Avatar Ready';$j.error=$null;$j.exception_kind=$null;$j.lease=$null}
 Save-LFAvatarJob $dir $j
}
# Validate existing Slide 1 media with D's real local validator; never rebuild it.
& (Join-Path $repo 'studio/Invoke-AvatarStep.ps1') -RuntimeRoot $root -Id $id -Job $jobs[0]
if($LASTEXITCODE){throw 'Slide 1 existing avatar validation failed.'}
Set-LFAvatarPause $root $id $true
$checks=[Collections.Generic.List[object]]::new()
function Check($Name,[bool]$OK){$checks.Add(@{name=$Name;passed=$OK});Write-LFJson (Join-Path $root 'results.json') @{checks=@($checks);project_id=$id;project_root=$dir;url="http://127.0.0.1:$Port/";elevenlabs_calls=0;heygen_calls=0};if(-not $OK){throw "FAIL: $Name"};"PASS: $Name"}
$ready=Get-LFAssemblyReadiness $root $id
Check 'All 14 existing approved avatars pass readiness' ($ready.ready -and $ready.snapshot.slides.Count -eq 14)
$job=Read-LFAvatarJob $dir $jobs[7];$saved=$job.output_path;$job.output_path='avatars/missing.mp4';Save-LFAvatarJob $dir $job
Check 'Missing avatar blocks assembly' (-not (Get-LFAssemblyReadiness $root $id).ready)
$job.output_path=$saved;$hash=$job.validation.narration_sha256;$job.validation.narration_sha256='STALE';Save-LFAvatarJob $dir $job
Check 'Stale audio/avatar binding blocks assembly' (-not (Get-LFAssemblyReadiness $root $id).ready)
$job.validation.narration_sha256=$hash;Save-LFAvatarJob $dir $job
$hashes=@(foreach($jid in $jobs){(Get-FileHash (Get-LFAvatarJobPath $dir $jid)).Hash}) -join ','
$pos=@{left=.84;top=.80;width=$p.preset.placement.width;height=$p.preset.placement.height}
Set-LFAssemblyPlacement $root $id 6 $pos 0;$ready=Get-LFAssemblyReadiness $root $id
Check 'Placement override only changes Slide 6 insertion metadata' ($ready.ready -and $ready.snapshot.slides[5].placement.left -eq .84 -and $ready.snapshot.slides[4].placement.left -eq .875 -and ((@(foreach($jid in $jobs){(Get-FileHash (Get-LFAvatarJobPath $dir $jid)).Hash}) -join ',') -eq $hashes))
Set-LFAssemblyPlacement $root $id 6 $null 1
Check 'Reset to project default restores approved placement' ((Get-LFAssemblyReadiness $root $id).snapshot.slides[5].placement.left -eq .875)
Set-LFAssemblyPlacement $root $id 6 $pos 2
# Simulate a dead local COM executor, without running or killing PowerPoint.
$ready=Get-LFAssemblyReadiness $root $id;$null=Request-LFAssembly $root $id $ready.binding (Join-Path $root 'Output')
$a=Read-LFAssembly $dir;$failedId=$a.current.id;$a.current.stage='Inserting avatars';$a.current.slide=4;$a.current.lease=@{pid=2147483647;started='0'};Write-LFJson (Join-Path $dir 'assembly.json') $a
$null=Start-LFAssemblyTask $root
Check 'Interrupted local assembly becomes a retryable exception' ((Read-LFAssembly $dir).current.stage -eq 'Exception')
# Existing filename must survive publication.
[void][IO.Directory]::CreateDirectory((Join-Path $root 'Output'));$existing=Join-Path $root 'Output/Solar Thermal_RECORDING.pptx';Copy-Item -LiteralPath (Join-Path $dir 'source.pptx') -Destination $existing;$existingHash=(Get-FileHash $existing).Hash
& (Join-Path $repo 'Start-LectureForge.ps1') -RuntimeRoot $root -Port $Port -NoBrowser
$url="http://127.0.0.1:$Port/";function Get($Path){Invoke-RestMethod ($url+$Path) -TimeoutSec 60};$token=(Get 'api/bootstrap').token
function Post($Path,$Body){Invoke-RestMethod -Method Post -Uri ($url+$Path) -Headers @{'X-LF-Token'=$script:token} -ContentType application/json -Body ($Body|ConvertTo-Json -Depth 15) -TimeoutSec 60}
function Restart { $null=Post 'api/stop' @{};for($i=0;$i -lt 60;$i++){Start-Sleep -Milliseconds 250;if((Read-LFSharedJson (Join-Path $root 'runtime.json')).status -eq 'stopped'){break}};& (Join-Path $repo 'Start-LectureForge.ps1') -RuntimeRoot $root -Port $Port -NoBrowser;$script:token=(Get 'api/bootstrap').token }
$stateHash=(Get-FileHash (Join-Path $dir 'assembly.json')).Hash;Restart
Check 'Studio restart before assembly preserves retry state' ((Get-FileHash (Join-Path $dir 'assembly.json')).Hash -eq $stateHash)

$ready=Get ("api/assembly-ready?id=$id")
$queued=Post ("api/assembly-create?id=$id") @{expected=$ready.binding;folder=(Join-Path $root 'Output')}
$again=Post ("api/assembly-create?id=$id") @{expected=$ready.binding;folder=(Join-Path $root 'Output')}
Check 'Double-click queues one fresh retry attempt' ($queued.current.id -eq $again.current.id -and $queued.current.id -ne $failedId)
$deadline=[DateTime]::UtcNow.AddMinutes(20);$stages=@()
do{Start-Sleep -Seconds 3;$s=Get ("api/assembly-status?id=$id");$stages+=$s.current.stage;"PROGRESS $($s.current.stage) slide=$($s.current.slide)";if($s.current.stage -eq 'Exception'){throw $s.current.error}}while($s.current.stage -ne 'Complete' -and [DateTime]::UtcNow -lt $deadline)
Check 'Background assembly completes while HTTP remains responsive' ($s.current.stage -eq 'Complete' -and $stages.Count -gt 1)
$v=Read-LFSharedJson (Join-Path $dir "r/$($s.current.id)/v/validation.json")
Check 'One save produces 14 slides and 14 embedded avatar objects' ($v.passed -and $v.save_calls -eq 1 -and $v.slide_count -eq 14 -and $v.avatar_count -eq 14)
Check 'Every slide retains native structure and identical native rendering' (@($v.slides|Where-Object {-not $_.native_preserved -or -not $_.native_render_identical}).Count -eq 0)
Check 'Every avatar is embedded, decodes, autoplays and moves/resizes independently' (@($v.slides|Where-Object {-not $_.independent_move_resize -or -not $_.autoplay -or -not $_.package.embedded_full_decode_ok}).Count -eq 0)
Check 'Only Slide 6 receives placement override in reopened deck' ($v.slides[5].placement.left -eq .84 -and @($v.slides|Where-Object {$_.slide -ne 6 -and $_.placement.left -ne .875}).Count -eq 0)
Check 'Slide 4 embeds approved Take 2 and excludes rejected Take 1' ($ready.snapshot.slides[3].take -eq 2 -and $v.slides[3].package.embedded_video_sha256 -eq $ready.snapshot.slides[3].sha256)
Check 'Slide 8 embedded media retains authoritative prepared six-second pause' ($v.slides[7].preparation.post_speech_silence_seconds -eq 6 -and $v.slides[7].narration_sha256 -eq 'E234676D7F92B42D20E98A3CB6CADF5BAA89595C2DCA906005EB7913607D69A5')
Check 'Existing output preserved with revision-safe exported filename' ((Get-FileHash $existing).Hash -eq $existingHash -and $s.current.output -ne $existing)
Check 'Canonical source unchanged and completed deck reopens' ($v.reopened -and $v.source_sha256_before -eq $v.source_sha256_after -and (Get-FileHash (Join-Path $repo 'source/RE-M2-Solar-Thermal-Overview.pptx')).Hash -eq $p.source.sha256)
$stateHash=(Get-FileHash (Join-Path $dir 'assembly.json')).Hash;$deckHash=(Get-FileHash $s.current.output).Hash;Restart
Check 'Restart after completion preserves identical state and deck hash' ((Get-FileHash (Join-Path $dir 'assembly.json')).Hash -eq $stateHash -and (Get-FileHash $s.current.output).Hash -eq $deckHash)
$opened=Post ("api/recording-open?id=$id") @{kind='deck'}
Check 'Open Recording PowerPoint launches the validated output' ($opened.opened -and $opened.path -eq $s.current.output)
$opened=Post ("api/recording-open?id=$id") @{kind='folder'}
Check 'Open Output Folder launches the validated output folder' $opened.opened
Check 'All Avatar Ready records and paid media survive assembly unchanged' (((@(foreach($jid in $jobs){(Get-FileHash (Get-LFAvatarJobPath $dir $jid)).Hash}) -join ',') -eq $hashes) -and (Get-LFAvatarStatus $root $id).ready -eq 14)
Check 'Zero ElevenLabs and HeyGen calls during local assembly and recovery' ((Get-LFNarration $root $id).provider_calls -eq 0 -and (Get-LFAvatarStatus $root $id).provider_calls -eq 0)
"RECORDING=$($s.current.output)"
