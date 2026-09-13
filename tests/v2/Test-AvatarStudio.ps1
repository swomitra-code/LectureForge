param([string]$RuntimeRoot=(Join-Path $env:LOCALAPPDATA 'LectureForgeD'),[int]$Port=8778)
$ErrorActionPreference='Stop';[Net.WebRequest]::DefaultWebProxy=$null
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'));Import-Module (Join-Path $repo 'studio/AvatarProduction.psm1') -Force
$root=Get-LFRoot $RuntimeRoot;if((Test-Path (Join-Path $root 'runtime.json')) -or (Test-Path (Join-Path $root 'Projects'))){throw 'Use a fresh isolated test root.'}
[void][IO.Directory]::CreateDirectory($root);Write-LFJson (Join-Path $root 'avatar-policy.json') @{provider_calls_enabled=$false;purpose='Zero-provider Milestone D acceptance only'}
$id=[guid]::NewGuid().ToString('N').Substring(0,12);$dir=Get-LFProjectPath $root $id
[void][IO.Directory]::CreateDirectory((Split-Path $dir));Copy-Item -LiteralPath (Join-Path $env:LOCALAPPDATA 'LectureForgeB/Projects/4b9d067085b3') -Destination $dir -Recurse
$p=Get-Content -Raw -Encoding UTF8 (Join-Path $dir 'project.json')|ConvertFrom-Json;$p.id=$id;$p.name='Solar Thermal - D OFFLINE COPY';Write-LFJson (Join-Path $dir 'project.json') $p
$checks=[Collections.Generic.List[object]]::new();function Check($Name,[bool]$OK){$checks.Add(@{name=$Name;passed=$OK});if(-not $OK){throw "FAIL: $Name"};"PASS: $Name"}
$production=Join-Path $repo 'work/solar-thermal-slides2-14-production-r001';$history=Get-Content -Raw (Join-Path $production 'selection-workflow-v2.json')|ConvertFrom-Json
$inputs=Get-LFSelectionInputs $root $id $dir;$catalog=@{assets=@()};$records=@{}
[void][IO.Directory]::CreateDirectory((Join-Path $dir 'avatars'))
foreach($row in $inputs.snapshot.slides){
 $h=@($history.history|Where-Object {$_.slide -eq $row.number -and -not $_.rejected -and $_.sha256 -eq $row.narration.sha256})|Select-Object -Last 1
 if(-not $h){throw "Missing approved fixture slide $($row.number)"}
 $job=Get-Content -Raw $h.path|ConvertFrom-Json;$src=Join-Path (Split-Path $h.path) 'avatar-white.mp4';$records[$row.number]=$job
 $rel="avatars/s$($row.number).mp4";Copy-Item -LiteralPath $src -Destination (Join-Path $dir $rel)
 $catalog.assets+=@{id="solar-s$($row.number)";slide_number=$row.number;status='approved';rejected=$false;video_id=$job.video_id;narration_sha256=$row.narration.sha256;preset_sha256=$inputs.snapshot.avatar_preset_sha256;path=$rel;sha256=(Get-FileHash $src).Hash;validation=@{passed=$true;authoritative_audio_verified=$true}}
}
$rejected=Get-Content -Raw (Join-Path $production 'jobs/slide-4/job.json')|ConvertFrom-Json
Copy-Item -LiteralPath (Join-Path $production 'jobs/slide-4/avatar-white.mp4') -Destination (Join-Path $dir 'avatars/rejected4.mp4')
$catalog.assets+=@{id='rejected4';slide_number=4;status='rejected';rejected=$true;narration_sha256=$rejected.narration_sha256;preset_sha256=$inputs.snapshot.avatar_preset_sha256;path='avatars/rejected4.mp4';sha256=(Get-FileHash (Join-Path $dir 'avatars/rejected4.mp4')).Hash;validation=@{passed=$true;authoritative_audio_verified=$true}}
Write-LFJson (Join-Path $dir 'avatars.json') $catalog
$c=Get-LFSelectionReview $root $id
Check 'All 13 exact approved white-avatar fixtures reused with zero expected jobs' ($c.current.snapshot.reusable_avatars -eq 13 -and $c.current.snapshot.expected_new_provider_jobs -eq 0)
Check 'Slide 4 approved Take 2 excludes rejected Take 1 history' ($c.current.snapshot.slides[2].selected_take -eq 2 -and $c.current.snapshot.slides[2].reusable_avatar.id -eq 'solar-s4')
$null=Save-LFSelectionReview $root $id $c.current.binding_sha256 confirmation;$auth=Confirm-LFAvatarAuthorization $root $id $c.current.binding_sha256 $c.current.binding_sha256
Start-LFAvatarBatch $root $id;Start-LFAvatarBatch $root $id;$q=Read-LFAvatarQueue $dir
Check 'Confirmed authorization consumed once into 13 durable jobs' ($q.jobs.Count -eq 13 -and $q.batches.Count -eq 1)
$claimed=@();for($i=0;$i -lt 4;$i++){$claimed+=@(Get-LFAvatarTasks $root 'test' 1)}
Check 'Scheduler grants exactly three concurrent leases' ($claimed.Count -eq 3)
foreach($jid in $q.jobs){$j=Read-LFAvatarJob $dir $jid;$j.lease=$null;Save-LFAvatarJob $dir $j}
Set-LFAvatarPause $root $id $true
Check 'Pause queue prevents new work without cancelling provider jobs' (@(Get-LFAvatarTasks $root 'test' 1).Count -eq 0)
Set-LFAvatarPause $root $id $false
# Reuse actual persisted provider IDs for restart recovery without network polling.
$j=Read-LFAvatarJob $dir $q.jobs[0];$j.video_id=$records[2].video_id;$j.submission_intent=$true;$j.stage='HeyGen Processing';$j.lease=@{owner='dead';pid=2147483647;started='0'};Save-LFAvatarJob $dir $j
$t=@(Get-LFAvatarTasks $root 'recovery' 1)[0];$recovered=Read-LFAvatarJob $dir $j.id
Check 'Dead worker recovery retains known HeyGen ID and polling stage' ($recovered.video_id -eq $records[2].video_id -and $recovered.stage -eq 'HeyGen Processing' -and $t.job -eq $j.id)
# Cached completed response/raw download fixture enters the real local converter.
$rawSource=Join-Path (Split-Path (@($history.history|Where-Object {$_.slide -eq 2 -and -not $_.rejected})|Select-Object -Last 1).path) 'avatar.webm'
$raw="a/$($j.id)/raw.webm";Copy-Item -LiteralPath $rawSource -Destination (Join-Path $dir $raw)
$j=Read-LFAvatarJob $dir $j.id;$j.lease=$null;$j.raw_path=$raw;$j.raw_sha256=(Get-FileHash (Join-Path $dir $raw)).Hash;$j.provider_status='completed';$j.provider_terminal=$true;$j.stage='Avatar Downloaded';Save-LFAvatarJob $dir $j
# An unknown submission fails closed on lease recovery; no second paid request.
$u=Read-LFAvatarJob $dir $q.jobs[1];$u.stage='Submitting to HeyGen';$u.video_id=$null;$u.submission_intent=$true;$u.lease=@{owner='dead';pid=2147483647;started='0'};Save-LFAvatarJob $dir $u
$null=Get-LFAvatarTasks $root 'recover' 3;$null=Get-LFAvatarTasks $root 'recover' 3;$u=Read-LFAvatarJob $dir $u.id
Check 'Interrupted submission without ID becomes uncertain exception' ($u.stage -eq 'Exception' -and $u.exception_kind -eq 'submission uncertain')
$blocked=$false;try{Repair-LFAvatarJob $root $id $u.id 'paid-retry'}catch{$blocked=$true};Check 'Generic retry cannot submit a replacement' $blocked
# Restore that synthetic exception using its original approved reusable fixture.
$u.stage='Queued';$u.video_id=$records[3].video_id;$u.submission_intent=$false;$u.error=$null;$u.exception_kind=$null;$u.lease=$null;Save-LFAvatarJob $dir $u
foreach($jid in $q.jobs){$j=Read-LFAvatarJob $dir $jid;$j.lease=$null;Save-LFAvatarJob $dir $j}
# Stale authorization rejected before starting any step; restoring selection needs new authorization.
$n=Get-LFNarration $root $id;Set-LFNarrationSelection $root $id 5 $n.slides[4].revision 1
$blocked=$false;try{$null=Assert-LFAvatarInputs $root $id $dir $auth.record.id}catch{$blocked=$true};Check 'Changed narration rejects stale consumed authorization' $blocked
Set-LFNarrationSelection $root $id 5 $n.slides[4].revision 3
# Preserve old snapshots and create a new confirmed fixture authorization.
$c=Get-LFSelectionReview $root $id;$null=Save-LFSelectionReview $root $id $c.current.binding_sha256 confirmation;$auth=Confirm-LFAvatarAuthorization $root $id $c.current.binding_sha256 $c.current.binding_sha256
# Historical jobs remain as explicit exceptions; old Slide 2 already-completed provider record is history.
foreach($jid in $q.jobs){$j=Read-LFAvatarJob $dir $jid;$j.stage='Exception';$j.error='Historical fixture authorization superseded';$j.provider_terminal=$true;Save-LFAvatarJob $dir $j}
Start-LFAvatarBatch $root $id;$q=Read-LFAvatarQueue $dir;$active=$q.batches[-1].jobs
# Exercise local conversion once from existing raw media; all other approved MP4s are reused.
$j=Read-LFAvatarJob $dir $active[0];$j.stage='Avatar Downloaded';$j.raw_path=$raw;$j.raw_sha256=(Get-FileHash (Join-Path $dir $raw)).Hash;$j.video_id=$records[2].video_id;$j.submission_intent=$true;$j.provider_terminal=$true;$j.provider_status='completed';Save-LFAvatarJob $dir $j
& (Join-Path $repo 'Start-LectureForge.ps1') -RuntimeRoot $root -Port $Port -NoBrowser
$url="http://127.0.0.1:$Port/";$token=(Invoke-RestMethod ($url+'api/bootstrap')).token
$deadline=[DateTime]::UtcNow.AddMinutes(15)
do{
 Start-Sleep -Seconds 5;$status=Invoke-RestMethod ($url+"api/avatar-status?id=$id") -TimeoutSec 20
 "PROGRESS ready=$($status.ready) exceptions=$($status.exceptions) queued=$($status.queued)"
 if($status.exceptions){$status.jobs|Where-Object stage -eq Exception|Select-Object slide,error|ConvertTo-Json|Write-Output;throw 'Local production exception'}
}while($status.ready -ne 13 -and [DateTime]::UtcNow -lt $deadline)
Check 'Background queue validates all 13 avatars without provider calls' ($status.ready -eq 13 -and $status.provider_calls -eq 0)
Check 'Real shared white converter passes downloaded/building/validating/ready progression' ('White Avatar Building' -in $status.jobs[0].history -and $status.jobs[0].validation.authoritative_audio_verified)
$s8=$status.jobs|Where-Object slide -eq 8
Check 'Slide 8 full audio packets and prepared 6.000-second silence validated' ($s8.validation.pause_verified -and $s8.validation.prepared_pause_seconds -eq 6 -and $s8.validation.narration_sha256 -eq 'E234676D7F92B42D20E98A3CB6CADF5BAA89595C2DCA906005EB7913607D69A5')
Check 'Canonical source and all reused narration remain unchanged' ((Get-FileHash (Join-Path $dir 'source.pptx')).Hash -eq $p.source.sha256 -and (Get-LFNarration $root $id).provider_calls -eq 0)
Write-LFJson (Join-Path $root 'results.json') @{checks=@($checks);project_id=$id;project_root=$dir;url=$url;elevenlabs_calls=0;heygen_calls=0;reused_avatars=13}
"RESULTS=$root/results.json"
