param([string]$RuntimeRoot=(Join-Path $env:LOCALAPPDATA 'LectureForgeDG'),[int]$Port=8893)
$ErrorActionPreference='Stop';$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'));Import-Module (Join-Path $repo 'studio/AvatarProduction.psm1') -Force
$root=Get-LFRoot $RuntimeRoot;if(Test-Path (Join-Path $root 'Projects')){throw 'Use a fresh guard-test root.'}
$id=[guid]::NewGuid().ToString('N').Substring(0,12);$dir=Get-LFProjectPath $root $id;[void][IO.Directory]::CreateDirectory((Split-Path $dir))
Copy-Item -LiteralPath (Join-Path $env:LOCALAPPDATA 'LectureForgeB/Projects/4b9d067085b3') -Destination $dir -Recurse
Write-LFJson (Join-Path $root 'avatar-policy.json') @{provider_calls_enabled=$false}
$p=Get-Content -Raw -Encoding UTF8 (Join-Path $dir 'project.json')|ConvertFrom-Json;$p.id=$id;foreach($slide in $p.slides){$slide.enabled=$slide.number -in 2,3,4,5};$p.version++;Write-LFJson (Join-Path $dir 'project.json') $p
$checks=@();function Check($Name,[bool]$OK){if(-not $OK){throw "FAIL: $Name"};$script:checks+=@{name=$Name;passed=$true};"PASS: $Name"}
function Authorize {$c=Get-LFSelectionReview $root $id;$null=Save-LFSelectionReview $root $id $c.current.binding_sha256 confirmation;Confirm-LFAvatarAuthorization $root $id $c.current.binding_sha256 $c.current.binding_sha256}
# Narration selections and read-only readiness must not authorize or enqueue work.
Import-Module (Join-Path $repo 'studio/Assembly.psm1')
Start-LFAvatarBatch $root $id
$readiness=Get-LFAssemblyReadiness $root $id
Check 'Selected narration and Check Readiness create no avatar queue or provider calls' (-not $readiness.ready -and (Read-LFAvatarQueue $dir).jobs.Count -eq 0 -and [int](Get-LFAvatarStatus $root $id).provider_calls -eq 0)
$n=Get-LFNarration $root $id;$selected=$n.slides[1].selected_take
Set-LFNarrationSelection $root $id 2 $n.slides[1].revision 0
$c=Get-LFSelectionReview $root $id
$blocked=$false;try{$null=Save-LFSelectionReview $root $id $c.current.binding_sha256 confirmation}catch{$blocked=$true}
Check 'Missing narration selection blocks production confirmation' ($blocked -and -not $c.current.snapshot.can_authorize)
Set-LFNarrationSelection $root $id 2 $n.slides[1].revision $selected
$c=Get-LFSelectionReview $root $id
$null=Save-LFSelectionReview $root $id $c.current.binding_sha256 review
$blocked=$false;try{$null=Confirm-LFAvatarAuthorization $root $id $c.current.binding_sha256 $c.current.binding_sha256}catch{$blocked=$true}
Check 'Authorization rejects review without final confirmation summary' $blocked
$null=Save-LFSelectionReview $root $id $c.current.binding_sha256 confirmation
Start-LFAvatarBatch $root $id
Check 'Viewing final production summary cannot queue jobs before explicit confirmation' ((Read-LFAvatarQueue $dir).jobs.Count -eq 0)
$a=Authorize;Start-LFAvatarBatch $root $id;$q=Read-LFAvatarQueue $dir
Check 'Four enabled slides create exactly four authorization-bound queue records' ($q.jobs.Count -eq 4)
$queuedSlides=@(foreach($jid in $q.jobs){(Read-LFAvatarJob $dir $jid).slide})
Check 'Confirmed queue contains exactly enabled slides 2, 3, 4, 5' (($queuedSlides -join ',') -eq '2,3,4,5')
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
$queue=Read-LFAvatarQueue $dir;$failed=Read-LFAvatarJob $dir $latest[-1];$failed.stage='Exception';$failed.provider_status='failed';$failed.provider_terminal=$true;$failed.error='HeyGen job failed. A replacement requires a new explicit authorization.';Save-LFAvatarJob $dir $failed
$untouched=@(foreach($jid in @($queue.jobs|Where-Object {$_ -ne $failed.id})){@{id=$jid;hash=(Get-FileHash (Get-LFAvatarJobPath $dir $jid)).Hash}})
$summary=Get-LFAvatarReplacementSummary $root $id $failed.id
Check 'Failed slide replacement summary is read-only and describes exactly one paid job' ($summary.slide -eq $failed.slide -and $summary.selected_take -eq $failed.take -and $summary.expected_new_provider_jobs -eq 1 -and $summary.reusable_avatars -eq 0 -and $summary.paid_replacement -and (Read-LFAvatarQueue $dir).jobs.Count -eq $queue.jobs.Count)
$blocked=$false;try{$null=Confirm-LFAvatarReplacement $root $id $failed.id $false $summary.selection_binding}catch{$blocked=$true};Check 'Failed avatar replacement requires explicit confirmation' $blocked
$replacement=Confirm-LFAvatarReplacement $root $id $failed.id $true $summary.selection_binding;$after=Read-LFAvatarQueue $dir;$new=Read-LFAvatarJob $dir $replacement.job
Check 'Confirmed replacement creates one new authorization and queues only failed slide' ($after.jobs.Count -eq $queue.jobs.Count+1 -and $new.slide -eq $failed.slide -and $new.stage -eq 'Queued' -and ($new.history -join ',') -eq 'Authorized,Queued' -and $new.authorization -ne $failed.authorization)
Check 'Replacement preserves narration, placement and all successful job bytes' ($new.take -eq $failed.take -and $new.row.narration.sha256 -eq $failed.row.narration.sha256 -and (ConvertTo-LFCanonicalJson $new.placement) -eq (ConvertTo-LFCanonicalJson $failed.placement) -and @($untouched|Where-Object {(Get-FileHash (Get-LFAvatarJobPath $dir $_.id)).Hash -ne $_.hash}).Count -eq 0)
$new.stage='Exception';$new.provider_status='failed';$new.provider_terminal=$true;$new.error='Second fixture failure';Save-LFAvatarJob $dir $new;$again=Get-LFAvatarReplacementSummary $root $id $new.id;$second=Confirm-LFAvatarReplacement $root $id $new.id $true $again.selection_binding
Check 'A second failure receives another fresh authorization and exactly one replacement job' ($second.authorization -ne $replacement.authorization -and (Read-LFAvatarQueue $dir).jobs.Count -eq $queue.jobs.Count+2)
$secondJob=Read-LFAvatarJob $dir $second.job;$secondJob.stage='Exception';$secondJob.provider_status='failed';$secondJob.provider_terminal=$true;$secondJob.error='HTTP fixture failure';Save-LFAvatarJob $dir $secondJob
$serverOut=Join-Path $root 'replacement-http.log';$serverErr=Join-Path $root 'replacement-http.err';$server=Start-Process (Join-Path $PSHOME 'powershell.exe') -WindowStyle Hidden -PassThru -RedirectStandardOutput $serverOut -RedirectStandardError $serverErr -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "'+(Join-Path $repo 'studio/Server.ps1')+'" -RuntimeRoot "'+$root+'" -Port '+$Port+' -Instance replacement-http -WorkerPid '+$PID)
$url="http://127.0.0.1:$Port/";$bootstrap=$null;for($i=0;$i -lt 40;$i++){Start-Sleep -Milliseconds 250;try{$bootstrap=Invoke-RestMethod ($url+'api/bootstrap') -TimeoutSec 1;break}catch{}}
if(-not $bootstrap){throw 'Replacement HTTP fixture did not start.'}
$beforeHttp=(Read-LFAvatarQueue $dir).jobs.Count;$preview=Invoke-RestMethod ($url+"api/avatar-replacement?id=$id&job=$($secondJob.id)") -TimeoutSec 10
Check 'Replacement preview endpoint is current and queues zero jobs' ($preview.paid_replacement -and $preview.expected_new_provider_jobs -eq 1 -and (Read-LFAvatarQueue $dir).jobs.Count -eq $beforeHttp)
$body=@{job=$secondJob.id;confirm=$true;expected=$preview.selection_binding}|ConvertTo-Json;$confirmed=Invoke-RestMethod -Method Post ($url+"api/avatar-replacement?id=$id") -Headers @{'X-LF-Token'=$bootstrap.token} -ContentType application/json -Body $body -TimeoutSec 10
Check 'Replacement confirmation endpoint queues exactly one current job' ($confirmed.stage -eq 'Queued' -and (Read-LFAvatarQueue $dir).jobs.Count -eq $beforeHttp+1 -and $confirmed.authorization -ne $secondJob.authorization)
$null=Invoke-RestMethod -Method Post ($url+'api/stop') -Headers @{'X-LF-Token'=$bootstrap.token} -ContentType application/json -Body '{}' -TimeoutSec 10;$server.WaitForExit(10000)|Out-Null
$serverText=(Get-Content -Raw $serverOut)+(Get-Content -Raw $serverErr);Check 'Current replacement endpoints never hit the Milestone E catch-all' ($serverText -notmatch [regex]::Escape('This action is not available in Milestone E.'))
Write-LFJson (Join-Path $root 'results.json') @{checks=$checks;elevenlabs_calls=0;heygen_calls=0;project_id=$id}
