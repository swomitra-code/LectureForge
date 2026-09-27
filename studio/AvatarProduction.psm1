$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Authorization.psm1')
Import-Module (Join-Path $PSScriptRoot 'HeyGenProvider.psm1')
function Test-LFAvatarSameAsset($A,$B){
 $A.slide -eq $B.slide -and $A.take -eq $B.take -and $A.row.narration.sha256 -eq $B.row.narration.sha256 -and (ConvertTo-LFCanonicalJson $A.preset) -eq (ConvertTo-LFCanonicalJson $B.preset)
}
function Resolve-LFAvatarSubmission($Dir,$Job){
 if($Job.exception_kind -ne 'existing submission'){return $Job}
 $oldId=$Job.existing_submission_id
 if(-not $oldId -and $Job.error -match '^Existing provider submission ([a-f0-9]{16}) must be reconciled before any replacement\.$'){$oldId=$Matches[1]}
 if(-not $oldId){throw 'Existing submission lineage is unavailable.'}
 $old=Read-LFAvatarJob $Dir $oldId
 if(-not $old.submission_intent -or -not (Test-LFAvatarSameAsset $old $Job)){throw 'Existing submission does not match the slide and narration.'}
 $old
}
function Assert-LFAvatarRecoveryInputs($Root,$Id,$Dir,$Job){
 # Finishing an existing paid asset does not authorize any new provider generation.
 if(-not $Job.submission_intent -or $Job.stage -in @('Queued','Submitting to HeyGen')){throw 'Recovery cannot enter a generation stage.'}
 $null=Assert-LFAvatarArtifactAuthorization $Dir $Job
 $current=Get-LFSelectionInputs $Root $Id $Dir
 $row=@($current.snapshot.slides|Where-Object number -eq $Job.slide)[0]
 if(-not $row -or $row.selected_take -ne $Job.take -or $row.narration.sha256 -ne $Job.row.narration.sha256 -or (ConvertTo-LFCanonicalJson $current.snapshot.avatar_preset) -ne (ConvertTo-LFCanonicalJson $Job.preset)){throw 'Recovered asset no longer matches the selected take, narration or avatar preset.'}
 if((Get-FileHash (Resolve-LFNarrationAsset $Dir $Job.row.narration.path)).Hash -ne $Job.row.narration.sha256){throw 'Authoritative narration changed.'}
}
function Recover-LFAvatarVideo($Root,$Id,$Job,[string]$Url){
 $candidate=Get-LFHeyGenUrlCandidate $Url
 Invoke-LFNarrationLock $Root $Id {param($dir)
  $q=Read-LFAvatarQueue $dir
  if($Job -notin $q.jobs){throw 'Submission is not in this project queue.'}
  $j=Resolve-LFAvatarSubmission $dir (Read-LFAvatarJob $dir $Job)
  if($j.video_id -and $j.video_id -ne $candidate){throw 'Submission is already associated with a different HeyGen video.'}
  if($j.recovery -and $j.video_id -eq $candidate){return @{recovered=$true;already_reconciled=$true;job=$j.id;stage=$j.stage}}
  if($j.lease){throw 'Work is active. Wait for completion before recovering.'}
  if(-not $j.submission_intent -or $j.stage -notin @('Exception','Submitting to HeyGen','Queued','HeyGen Processing')){throw 'No unresolved existing provider submission to recover.'}
  # A crashed queued/submitting record may be recovered, but never run as Queued.
  $originalStage=$j.stage;$j.stage='Exception'
  Assert-LFAvatarRecoveryInputs $Root $Id $dir $j
  $j.stage=$originalStage
  $policy=Join-Path $Root 'avatar-policy.json'
  if((Test-Path $policy) -and (Get-Content -Raw $policy|ConvertFrom-Json).provider_calls_enabled -eq $false){throw 'Provider calls disabled for this isolated fixture runtime.'}
  # No durable fields change until the read-only provider lookup and binding checks succeed.
  $video=Get-LFHeyGenRecoveryVideo $Url $Id $j.slide $j.take
  $j.video_id=[string]$video.id;$j.provider_status=$video.status;$j.provider_terminal=$video.status -in @('completed','failed');$j.provider_calls++
  $j|Add-Member -Force NoteProperty recovery ([pscustomobject]@{url=$Url;video_id=$video.id;verified_title=$video.title;timestamp=[DateTime]::UtcNow.ToString('o');submission=$j.id})
  $j.history+=@('ReconciliationStarted','ProviderJobRecovered');$j.error=$null;$j.exception_kind=$null;$j.next_poll_utc=$null
  $j.stage='HeyGen Processing'
  if($video.status -eq 'completed' -and $video.video_url){$j.download_url=$video.video_url;$j.stage='Downloading Avatar'}
  if($video.status -eq 'failed'){$j.stage='Exception';$j.exception_kind='provider recovery';$j.error='Recovered HeyGen video failed. A replacement requires a new explicit authorization.'}
  $j.history+=,$j.stage;Save-LFAvatarJob $dir $j
  @{recovered=$true;already_reconciled=$false;job=$j.id;stage=$j.stage}
 }
}
function Resolve-LFAvatarPath($Dir,$Relative){
 if($Relative -notmatch '^(avatars/[a-zA-Z0-9_-]{1,40}\.mp4|a/[a-f0-9]{16}/(raw[a-z0-9-]*\.webm|w-[a-f0-9]{8}\.mp4))$'){throw 'Invalid avatar media path.'}
 $path=[IO.Path]::GetFullPath((Join-Path $Dir $Relative));if($path.Length -ge 250 -or -not $path.StartsWith(([IO.Path]::GetFullPath($Dir).TrimEnd('\')+'\'),[StringComparison]::OrdinalIgnoreCase)){throw 'Unsafe avatar media path.'};$path
}
function Read-LFAvatarQueue($Dir){
 $file=Join-Path $Dir 'avatar-queue.json'
 if(Test-Path $file){Get-Content -Raw -Encoding UTF8 $file|ConvertFrom-Json}else{[pscustomobject]@{schema='lectureforge-avatar-queue-1';paused=$false;batches=@();jobs=@()}}
}
function Get-LFAvatarJobPath($Dir,$Job){
 if($Job -notmatch '^[a-f0-9]{16}$'){throw 'Invalid avatar job identifier.'};Join-Path $Dir "a/$Job/job.json"
}
function Read-LFAvatarJob($Dir,$Job){Get-Content -Raw -Encoding UTF8 (Get-LFAvatarJobPath $Dir $Job)|ConvertFrom-Json}
function Save-LFAvatarJob($Dir,$Job){$Job.updated_utc=[DateTime]::UtcNow.ToString('o');Write-LFJson (Get-LFAvatarJobPath $Dir $Job.id) $Job}
function Assert-LFAvatarJobBinding($Snapshot,$Job){
 $row=@($Snapshot.slides|Where-Object number -eq $Job.slide)[0]
 if(-not $row -or $Job.take -ne $row.selected_take -or (ConvertTo-LFCanonicalJson $Job.row) -ne (ConvertTo-LFCanonicalJson $row) -or (ConvertTo-LFCanonicalJson $Job.preset) -ne (ConvertTo-LFCanonicalJson $Snapshot.avatar_preset)){throw 'Avatar job inputs do not match the immutable authorization.'}
}
function Assert-LFAvatarInputs($Root,$Id,$Dir,$Authorization){
 $a=Read-LFApproval $Dir $Authorization 'authorization';$supported=Get-LFAvatarPreset ([pscustomobject]@{preset=Get-LFPreset});if((ConvertTo-LFCanonicalJson $supported) -ne (ConvertTo-LFCanonicalJson $a.snapshot.avatar_preset)){throw 'Unsupported avatar preset. Use the proven Renewable Energy preset.'};$c=Get-LFSelectionInputs $Root $Id $Dir
 $selectionBinding=if($a.snapshot.PSObject.Properties.Name -contains 'selection_binding'){$a.snapshot.selection_binding}else{$a.binding_sha256}
 if(-not $c.snapshot.can_authorize -or $c.binding_sha256 -ne $selectionBinding){throw 'Authorization stale. Return to Review Selections.'};$a
}
function Assert-LFAvatarArtifactAuthorization($Dir,$Job){
 $a=Read-LFApproval $Dir $Job.authorization 'authorization'
 Assert-LFAvatarJobBinding $a.snapshot $Job
 $receiptPath=Get-LFApprovalPath $Dir $Job.authorization 'consumed'
 if(-not (Test-Path $receiptPath)){throw 'Avatar job authorization was not consumed by the production worker.'}
 $receipt=Get-Content -Raw -Encoding UTF8 $receiptPath|ConvertFrom-Json
 $authorizationHash=(Get-FileHash (Get-LFApprovalPath $Dir $Job.authorization 'authorization')).Hash
 if($receipt.consumer -ne 'lectureforge-avatar-worker-1' -or $receipt.authorization -ne $Job.authorization -or $receipt.authorization_sha256 -ne $authorizationHash){throw 'Avatar job authorization receipt is invalid.'}
 $a
}
function Get-LFAvatarReplacementSummary($Root,$Id,$Job){
 Invoke-LFNarrationLock $Root $Id {param($dir)
  $j=Read-LFAvatarJob $dir $Job;if($j.stage -ne 'Exception' -or $j.provider_status -ne 'failed'){throw 'Only a failed HeyGen generation may be re-authorized.'}
  $current=Get-LFSelectionInputs $Root $Id $dir;$row=@($current.snapshot.slides|Where-Object number -eq $j.slide)[0]
  if(-not $row -or $row.selected_take -ne $j.take -or $row.narration.sha256 -ne $j.row.narration.sha256){throw 'Failed job no longer matches the selected narration.'}
  [ordered]@{job=$j.id;slide=$row.number;title=$row.title;selected_take=$row.selected_take;avatar_preset=$current.snapshot.avatar_preset;expected_new_provider_jobs=1;reusable_avatars=0;paid_replacement=$true;selection_binding=$current.binding_sha256;placement=$j.placement}
 }
}
function Confirm-LFAvatarReplacement($Root,$Id,$Job,[bool]$Confirm,[string]$Expected){
 if(-not $Confirm){throw 'Explicit paid replacement confirmation is required.'}
 Invoke-LFNarrationLock $Root $Id {param($dir)
  $j=Read-LFAvatarJob $dir $Job;if($j.stage -ne 'Exception' -or $j.provider_status -ne 'failed'){throw 'Only a failed HeyGen generation may be re-authorized.'}
  $current=Get-LFSelectionInputs $Root $Id $dir;$row=@($current.snapshot.slides|Where-Object number -eq $j.slide)[0]
  if($Expected -ne $current.binding_sha256 -or -not $row -or $row.selected_take -ne $j.take -or $row.narration.sha256 -ne $j.row.narration.sha256){throw 'Replacement summary is stale. Review it again.'}
  $q=Read-LFAvatarQueue $dir;$sequence=@($q.batches|Where-Object replacement_of -eq $j.id).Count+1
  $snapshot=[ordered]@{selection_binding=$current.binding_sha256;replacement_of=$j.id;replacement_sequence=$sequence;project_revision=$current.snapshot.project_revision;intent_revision=$current.snapshot.intent_revision;avatar_preset=$current.snapshot.avatar_preset;avatar_preset_sha256=$current.snapshot.avatar_preset_sha256;default_placement=$current.snapshot.default_placement;slides=@($row);expected_new_provider_jobs=1;reusable_avatars=0;paid_replacement=$true}
  $aid=Get-LFTextHash (ConvertTo-LFCanonicalJson $snapshot);[void][IO.Directory]::CreateDirectory((Join-Path $dir 'approvals'));$authPath=Get-LFApprovalPath $dir $aid authorization
  if(Test-Path (Get-LFApprovalPath $dir $aid consumed)){throw 'Replacement authorization already consumed. Review and authorize a new retry.'}
  if(-not (Test-Path $authPath)){Write-LFImmutableJson $authPath @{schema='lectureforge-avatar-replacement-authorization-1';id=$aid;timestamp=[DateTime]::UtcNow.ToString('o');project_id=$Id;binding_sha256=$aid;snapshot=$snapshot;authorization='one paid replacement avatar generation';provider_jobs_submitted=0;consumer_milestone='D'}}
  Write-LFImmutableJson (Get-LFApprovalPath $dir $aid consumed) @{consumer='lectureforge-avatar-worker-1';authorization=$aid;authorization_sha256=(Get-FileHash $authPath).Hash;timestamp=[DateTime]::UtcNow.ToString('o')}
  $jid=(Get-LFTextHash ($aid+'/'+$row.number)).Substring(0,16).ToLowerInvariant();$path=Get-LFAvatarJobPath $dir $jid;[void][IO.Directory]::CreateDirectory((Split-Path $path))
  $replacement=[ordered]@{id=$jid;authorization=$aid;slide=$row.number;take=$row.selected_take;row=$row;preset=$snapshot.avatar_preset;placement=$j.placement;stage='Queued';history=@('Authorized','Queued');error=$null;exception_kind=$null;lease=$null;submission_intent=$false;video_id=$null;audio_asset_id=$null;provider_status=$null;provider_terminal=$false;provider_calls=0;next_poll_utc=$null;download_url=$null;raw_path=$null;raw_sha256=$null;output_path=$null;output_sha256=$null;validation=$null;attempt=0;updated_utc=[DateTime]::UtcNow.ToString('o')}
  Write-LFJson $path $replacement;$q.jobs=@($q.jobs+$jid|Select-Object -Unique);$q.batches+=@{authorization=$aid;jobs=@($jid);replacement_of=$j.id};Write-LFJson (Join-Path $dir 'avatar-queue.json') $q
  @{authorization=$aid;job=$jid;stage='Queued'}
 }
}
function Start-LFAvatarBatch($Root,$Id){
 Invoke-LFNarrationLock $Root $Id {param($dir)
  $reviewFile=Join-Path $dir 'selection-review.json';if(-not (Test-Path $reviewFile)){return}
  $saved=Get-Content -Raw $reviewFile|ConvertFrom-Json
  if($saved.phase -ne 'authorized' -or -not $saved.review_id){return}
  $queue=Read-LFAvatarQueue $dir
  if($queue.paused){return}
  if($saved.review_id -in @($queue.batches.authorization)){return}
  $receipt=Get-LFApprovalPath $dir $saved.review_id 'consumed'
  try{$auth=Assert-LFAvatarInputs $Root $Id $dir $saved.review_id}catch{return}
  if(Test-Path $receipt){$claim=Get-Content -Raw $receipt|ConvertFrom-Json;if($claim.consumer -ne 'lectureforge-avatar-worker-1' -or $claim.authorization -ne $auth.id -or $claim.authorization_sha256 -ne (Get-FileHash (Get-LFApprovalPath $dir $auth.id 'authorization')).Hash){throw 'Incompatible consumed authorization. No work started.'}}
  else{Write-LFImmutableJson $receipt @{consumer='lectureforge-avatar-worker-1';authorization=$auth.id;authorization_sha256=(Get-FileHash (Get-LFApprovalPath $dir $auth.id 'authorization')).Hash;timestamp=[DateTime]::UtcNow.ToString('o')}}
  # Deterministic IDs recover a crash between receipt, job creation, and queue save.
  $ids=@(foreach($row in $auth.snapshot.slides){
   $jid=(Get-LFTextHash ($auth.id+'/'+$row.number)).Substring(0,16).ToLowerInvariant();$path=Get-LFAvatarJobPath $dir $jid
   if(-not (Test-Path $path)){
    [void][IO.Directory]::CreateDirectory((Split-Path $path))
    $job=[ordered]@{id=$jid;authorization=$auth.id;slide=$row.number;take=$row.selected_take;row=$row;preset=$auth.snapshot.avatar_preset;placement=if($row.placement_override){$row.placement_override}else{$auth.snapshot.default_placement};stage='Queued';history=@('Authorized','Queued');error=$null;exception_kind=$null;lease=$null;submission_intent=$false;video_id=if($row.reusable_avatar){$row.reusable_avatar.video_id}else{$null};audio_asset_id=$null;provider_status=$null;provider_terminal=$false;provider_calls=0;next_poll_utc=$null;download_url=$null;raw_path=$null;raw_sha256=$null;output_path=$null;output_sha256=$null;validation=$null;attempt=0;updated_utc=[DateTime]::UtcNow.ToString('o')}
    # Reauthorizing an identical take cannot silently create a duplicate live job.
    foreach($oldId in $queue.jobs){
     $old=Read-LFAvatarJob $dir $oldId
     if($old.row.narration.sha256 -eq $row.narration.sha256 -and (Get-LFTextHash (ConvertTo-LFCanonicalJson $old.preset)) -eq $auth.snapshot.avatar_preset_sha256 -and $old.slide -eq $row.number){
      if($old.stage -eq 'Avatar Ready' -and $old.validation.passed -and (Test-Path (Resolve-LFAvatarPath $dir $old.output_path)) -and (Get-FileHash (Resolve-LFAvatarPath $dir $old.output_path)).Hash -eq $old.output_sha256){$job.output_path=$old.output_path;$job.output_sha256=$old.output_sha256;$job.stage='Validating';break}
      if($old.provider_status -eq 'completed' -and $old.video_id){
       $job.video_id=$old.video_id;$job.submission_intent=$true;$job.provider_status='completed';$job.provider_terminal=$true;$job.audio_asset_id=$old.audio_asset_id;$job.stage='HeyGen Processing'
       if($old.raw_path -and (Test-Path (Resolve-LFAvatarPath $dir $old.raw_path)) -and (Get-FileHash (Resolve-LFAvatarPath $dir $old.raw_path)).Hash -eq $old.raw_sha256){$job.raw_path=$old.raw_path;$job.raw_sha256=$old.raw_sha256;$job.stage='Avatar Downloaded'}
       break
      }
      if($old.submission_intent -and -not $old.provider_terminal){$job.existing_submission_id=$old.id;$job.stage='Exception';$job.exception_kind='existing submission';$job.error='Existing provider submission '+$old.id+' must be reconciled before any replacement.';break}
     }
    }
    Write-LFJson $path $job
   };$jid
  })
  $queue.jobs=@($queue.jobs+$ids|Select-Object -Unique);$queue.batches+=@{authorization=$auth.id;jobs=$ids};Write-LFJson (Join-Path $dir 'avatar-queue.json') $queue
 }
}
function Get-LFAvatarStatus($Root,$Id){
 # Read-only, no provider requests, no media hashing on frequent dashboard reads.
 Invoke-LFNarrationLock $Root $Id {param($dir)
  $q=Read-LFAvatarQueue $dir;$jobs=@(foreach($jid in $q.jobs){Read-LFAvatarJob $dir $jid})
  $latest=$q.batches|Select-Object -Last 1;$active=@($jobs|Group-Object slide|ForEach-Object {$_.Group|Select-Object -Last 1}|Sort-Object slide)
  # Blocked rows follow their recovered original; only that original is processed.
  $active=@(foreach($j in $active){if($j.exception_kind -eq 'existing submission'){$source=Resolve-LFAvatarSubmission $dir $j;if($source.recovery){$source}else{$j}}else{$j}})
  $stale=$false;if($latest){$a=Read-LFApproval $dir $latest.authorization 'authorization';$p=Get-Content -Raw (Join-Path $dir 'project.json')|ConvertFrom-Json;$n=Read-LFNarrationFile $dir;$intent=if($n.PSObject.Properties.Name -contains 'intent_revision'){$n.intent_revision}else{0};$stale=($p.version -ne $a.snapshot.project_revision -or $intent -ne $a.snapshot.intent_revision)}
  @{stale=$stale;paused=$q.paused;authorization=if($latest){$latest.authorization}else{$null};jobs=$active;historical_jobs=@($jobs|Where-Object {$_.id -notin $active.id});authorized=$active.Count;ready=@($active|Where-Object stage -eq 'Avatar Ready').Count;exceptions=@($active|Where-Object stage -eq 'Exception').Count;queued=@($active|Where-Object stage -eq 'Queued').Count;provider_processing=@($active|Where-Object {$_.submission_intent -and -not $_.provider_terminal}).Count;provider_calls=($jobs|Measure-Object provider_calls -Sum).Sum}
 }
}
function Set-LFAvatarPause($Root,$Id,[bool]$Paused){Invoke-LFNarrationLock $Root $Id {param($dir) $q=Read-LFAvatarQueue $dir;$q.paused=$Paused;Write-LFJson (Join-Path $dir 'avatar-queue.json') $q}}
function Repair-LFAvatarJob($Root,$Id,$Job,$Action,$VideoId=$null){
 if($Action -eq 'reconcile'){throw 'Use Recover Existing HeyGen Video with the provider video URL.'}
 Invoke-LFNarrationLock $Root $Id {param($dir)
  $j=Read-LFAvatarJob $dir $Job;if($j.lease){throw 'Work is active. Wait for completion.'};if($j.stage -ne 'Exception'){throw 'Only an exception may be retried.'}
  if($j.recovery){Assert-LFAvatarRecoveryInputs $Root $Id $dir $j}else{$authorization=Assert-LFAvatarInputs $Root $Id $dir $j.authorization;Assert-LFAvatarJobBinding $authorization.snapshot $j}
  switch($Action){
   'resume'{if(-not $j.video_id -or $j.provider_status -eq 'failed'){throw 'No resumable provider job.'};$j.stage='HeyGen Processing'}
   'download'{if(-not $j.video_id -or $j.provider_status -ne 'completed'){throw 'No completed job to download.'};$j.stage='HeyGen Processing'}
   'convert'{if(-not $j.raw_path){throw 'No downloaded avatar.'};$j.stage='Avatar Downloaded'}
   'validate'{
    if(-not $j.output_path -and $j.row.reusable_avatar){
     $reuse=$j.row.reusable_avatar;$path=Resolve-LFReusableAsset $dir $reuse.path
     if((Get-FileHash $path).Hash -ne $reuse.sha256){throw 'Reusable avatar hash mismatch.'}
     $j.output_path=$reuse.path;$j.output_sha256=$reuse.sha256
    }
    if(-not $j.output_path){throw 'No white avatar to validate.'};$j.stage='Validating'
   }
   default{throw 'No paid retry action is supported. Review and explicitly authorize replacements.'}
  };$j.error=$null;$j.exception_kind=$null;$j.next_poll_utc=$null;Save-LFAvatarJob $dir $j
 }
}
function Set-LFAvatarException($Root,$Id,$Job,$Message,$Kind){Invoke-LFNarrationLock $Root $Id {param($dir) $j=Read-LFAvatarJob $dir $Job;$j.stage='Exception';$j.error=$Message;$j.exception_kind=$Kind;$j.lease=$null;Save-LFAvatarJob $dir $j}}
function Get-LFAvatarTasks($Root,$Owner,[int]$Capacity=3){
 $tasks=@();$occupied=0;$liveLeases=0
 # A known nonterminal job or an uncertain submission holds a provider slot, even across restart.
 foreach($p in Get-LFProjects $Root){$d=Get-LFProjectPath $Root $p.id;$q=Read-LFAvatarQueue $d;foreach($jid in $q.jobs){$j=Read-LFAvatarJob $d $jid;if($j.lease -and $j.lease.pid){$proc=Get-Process -Id $j.lease.pid -ErrorAction SilentlyContinue;if($proc -and $proc.StartTime.ToUniversalTime().Ticks.ToString() -eq $j.lease.started){$liveLeases++;if($j.stage -eq 'Queued' -and -not $j.row.reusable_avatar){$occupied++}}};if($j.submission_intent -and -not $j.provider_terminal){$occupied++}}}
 if($liveLeases -ge 3){return}
 foreach($p in Get-LFProjects $Root){
  if($tasks.Count -ge $Capacity){break}
  $selected=Invoke-LFNarrationLock $Root $p.id {param($dir)
   $q=Read-LFAvatarQueue $dir;if($q.paused){return};$c=$null
   foreach($jid in $q.jobs){
    $j=Read-LFAvatarJob $dir $jid
    if($j.stage -in @('Avatar Ready','Exception')){continue}
    if($j.lease){
     $alive=$false;if($j.lease.pid){$proc=Get-Process -Id $j.lease.pid -ErrorAction SilentlyContinue;$alive=$proc -and $proc.StartTime.ToUniversalTime().Ticks.ToString() -eq $j.lease.started}
     if($alive){continue}
     if($j.stage -eq 'Submitting to HeyGen' -and -not $j.video_id){$j.stage='Exception';$j.exception_kind='submission uncertain';$j.error='Submission interrupted without a job ID. Reconcile the existing job; never resubmit.';$j.lease=$null;Save-LFAvatarJob $dir $j;continue}
     if($j.stage -in @('White Avatar Building','Validating')){$j.stage='Exception';$j.exception_kind='local interrupted';$j.error='Local work interrupted. Existing outputs preserved; use targeted retry.';$j.lease=$null;Save-LFAvatarJob $dir $j;continue}
     $j.lease=$null
    }
    if($j.next_poll_utc -and [DateTimeOffset]::Parse($j.next_poll_utc).UtcDateTime -gt [DateTime]::UtcNow){continue}
    try{if($j.recovery){Assert-LFAvatarRecoveryInputs $Root $p.id $dir $j}else{if(-not $c){$c=Get-LFSelectionInputs $Root $p.id $dir};if($c.binding_sha256 -ne $j.authorization){throw 'Authorization stale. Return to Review Selections.'};Assert-LFAvatarJobBinding $c.snapshot $j}}catch{$j.stage='Exception';$j.exception_kind='stale authorization';$j.error=$_.Exception.Message;Save-LFAvatarJob $dir $j;continue}
    if($j.stage -eq 'Queued' -and -not $j.row.reusable_avatar -and $occupied -ge 3){continue}
    $j.lease=@{owner=$Owner;pid=$PID;started=(Get-Process -Id $PID).StartTime.ToUniversalTime().Ticks.ToString()};Save-LFAvatarJob $dir $j
    return @{id=$p.id;job=$jid;new_provider=($j.stage -eq 'Queued' -and -not $j.row.reusable_avatar)}
   }
  }
  if($selected){$tasks+=,$selected;if($selected.new_provider){$occupied++}}
 }
 # One per project per scheduling pass; the main worker immediately repeats up to its capacity.
 $tasks
}
Export-ModuleMember -Function *-LF*
