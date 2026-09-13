param([string]$RuntimeRoot=(Join-Path $env:LOCALAPPDATA 'LectureForgeDG'))
$ErrorActionPreference='Stop';$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'));Import-Module (Join-Path $repo 'studio/AvatarProduction.psm1') -Force
$root=Get-LFRoot $RuntimeRoot;if(Test-Path (Join-Path $root 'Projects')){throw 'Use a fresh guard-test root.'}
$id=[guid]::NewGuid().ToString('N').Substring(0,12);$dir=Get-LFProjectPath $root $id;[void][IO.Directory]::CreateDirectory((Split-Path $dir))
Copy-Item -LiteralPath (Join-Path $env:LOCALAPPDATA 'LectureForgeB/Projects/4b9d067085b3') -Destination $dir -Recurse
Write-LFJson (Join-Path $root 'avatar-policy.json') @{provider_calls_enabled=$false}
$p=Get-Content -Raw -Encoding UTF8 (Join-Path $dir 'project.json')|ConvertFrom-Json;$p.id=$id;foreach($slide in $p.slides){$slide.enabled=$slide.number -in 2,3,4,5};$p.version++;Write-LFJson (Join-Path $dir 'project.json') $p
$checks=@();function Check($Name,[bool]$OK){if(-not $OK){throw "FAIL: $Name"};$script:checks+=@{name=$Name;passed=$true};"PASS: $Name"}
function Authorize {$c=Get-LFSelectionReview $root $id;$null=Save-LFSelectionReview $root $id $c.current.binding_sha256 confirmation;Confirm-LFAvatarAuthorization $root $id $c.current.binding_sha256 $c.current.binding_sha256}
$a=Authorize;Start-LFAvatarBatch $root $id;$q=Read-LFAvatarQueue $dir
Check 'Four enabled slides create exactly four authorization-bound queue records' ($q.jobs.Count -eq 4)
$jobHashes=@(foreach($jid in $q.jobs){(Get-FileHash (Get-LFAvatarJobPath $dir $jid)).Hash}) -join ','
Write-LFJson (Join-Path $dir 'avatar-queue.json') @{schema=$q.schema;paused=$false;batches=@();jobs=@()}
Start-LFAvatarBatch $root $id;$q=Read-LFAvatarQueue $dir
Check 'Receipt-to-queue crash recovery reuses identical job records' (($jobHashes -eq (@(foreach($jid in $q.jobs){(Get-FileHash (Get-LFAvatarJobPath $dir $jid)).Hash}) -join ',') -and $q.jobs.Count -eq 4))
$tasks=@();for($i=0;$i -lt 4;$i++){$tasks+=@(Get-LFAvatarTasks $root 'guard' 1)}
Check 'At most three unsubmitted provider jobs receive reservations' ($tasks.Count -eq 3)
foreach($jid in $q.jobs){$j=Read-LFAvatarJob $dir $jid;$j.lease=$null;if($j.slide -lt 5){$j.submission_intent=$true;$j.video_id="fixture-known-$($j.slide)";$j.stage='HeyGen Processing';$j.next_poll_utc='2099-01-01T00:00:00Z'};Save-LFAvatarJob $dir $j}
Check 'Three existing nonterminal jobs prevent a fourth paid start' (@(Get-LFAvatarTasks $root 'guard' 1).Count -eq 0)
$n=Get-LFNarration $root $id;Set-LFNarrationSelection $root $id 5 $n.slides[4].revision 1
$blocked=$false;try{$null=Assert-LFAvatarInputs $root $id $dir $a.record.id}catch{$blocked=$true};Check 'Consumed snapshot rejects changed selection before work' $blocked
$b=Authorize;Start-LFAvatarBatch $root $id;$new=(Read-LFAvatarQueue $dir).batches[-1].jobs;$jobs=@(foreach($jid in $new){Read-LFAvatarJob $dir $jid})
Check 'New authorization cannot duplicate three prior active submissions' (@($jobs|Where-Object exception_kind -eq 'existing submission').Count -eq 3)
$j=$jobs[3];$j.row.narration.sha256='BAD';$blocked=$false;try{Assert-LFAvatarJobBinding $b.record.snapshot $j}catch{$blocked=$true};Check 'Tampered job narration is rejected against immutable snapshot' $blocked
Check 'Authorization bytes stay unchanged after consumption and invalidation' ((Get-FileHash (Get-LFApprovalPath $dir $a.record.id authorization)).Hash -eq $a.file_sha256)
$code=Get-Content -Raw (Join-Path $repo 'studio/Invoke-AvatarStep.ps1')
Check 'Submission intent is persisted before POST and ID before polling (source boundary check)' ($code.IndexOf('$live.submission_intent=$true') -lt $code.IndexOf("Request POST '/v3/videos'") -and $code.IndexOf('video_id=[string]$created.video_id') -lt $code.IndexOf("elseif(`$j.stage -eq 'HeyGen Processing')"))
Check 'Guard tests create no narration or avatar provider calls' ((Get-LFNarration $root $id).provider_calls -eq 0 -and (Get-LFAvatarStatus $root $id).provider_calls -eq 0)
$old=Read-LFAvatarJob $dir $q.jobs[0];$old.provider_status='completed';$old.provider_terminal=$true;Save-LFAvatarJob $dir $old
Set-LFNarrationSelection $root $id 5 $n.slides[4].revision 3;$c=Authorize;Start-LFAvatarBatch $root $id
$latest=(Read-LFAvatarQueue $dir).batches[-1].jobs;$reused=Read-LFAvatarJob $dir $latest[0]
Check 'Completed provider ID is reused even when local conversion previously failed' ($reused.video_id -eq $old.video_id -and $reused.stage -eq 'HeyGen Processing' -and $reused.provider_calls -eq 0)
Write-LFJson (Join-Path $root 'results.json') @{checks=$checks;elevenlabs_calls=0;heygen_calls=0;project_id=$id}
