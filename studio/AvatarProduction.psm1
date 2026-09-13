$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Authorization.psm1')
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
 if(-not $c.snapshot.can_authorize -or $c.binding_sha256 -ne $a.binding_sha256){throw 'Authorization stale. Return to Review Selections.'};$a
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
      if($old.submission_intent -and -not $old.provider_terminal){$job.stage='Exception';$job.exception_kind='existing submission';$job.error='Existing provider submission '+$old.id+' must be reconciled before any replacement.';break}
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
  $latest=$q.batches|Select-Object -Last 1;$active=@($jobs|Where-Object {$latest -and $_.id -in $latest.jobs})
  $stale=$false;if($latest){$a=Read-LFApproval $dir $latest.authorization 'authorization';$p=Get-Content -Raw (Join-Path $dir 'project.json')|ConvertFrom-Json;$n=Read-LFNarrationFile $dir;$intent=if($n.PSObject.Properties.Name -contains 'intent_revision'){$n.intent_revision}else{0};$stale=($p.version -ne $a.snapshot.project_revision -or $intent -ne $a.snapshot.intent_revision)}
  @{stale=$stale;paused=$q.paused;authorization=if($latest){$latest.authorization}else{$null};jobs=$active;historical_jobs=@($jobs|Where-Object {$_.id -notin $active.id});authorized=$active.Count;ready=@($active|Where-Object stage -eq 'Avatar Ready').Count;exceptions=@($active|Where-Object stage -eq 'Exception').Count;queued=@($active|Where-Object stage -eq 'Queued').Count;provider_processing=@($active|Where-Object {$_.submission_intent -and -not $_.provider_terminal}).Count;provider_calls=($jobs|Measure-Object provider_calls -Sum).Sum}
 }
}
function Set-LFAvatarPause($Root,$Id,[bool]$Paused){Invoke-LFNarrationLock $Root $Id {param($dir) $q=Read-LFAvatarQueue $dir;$q.paused=$Paused;Write-LFJson (Join-Path $dir 'avatar-queue.json') $q}}
function Repair-LFAvatarJob($Root,$Id,$Job,$Action,$VideoId=$null){
 Invoke-LFNarrationLock $Root $Id {param($dir)
  $j=Read-LFAvatarJob $dir $Job;if($j.lease){throw 'Work is active. Wait for completion.'};if($j.stage -ne 'Exception'){throw 'Only an exception may be retried.'}
  $authorization=Assert-LFAvatarInputs $Root $Id $dir $j.authorization
  Assert-LFAvatarJobBinding $authorization.snapshot $j
  switch($Action){
   'resume'{if(-not $j.video_id -or $j.provider_status -eq 'failed'){throw 'No resumable provider job.'};$j.stage='HeyGen Processing'}
   'reconcile'{if(-not $j.submission_intent -or $j.video_id -or $VideoId -notmatch '^[a-zA-Z0-9_-]{8,100}$'){throw 'Enter the existing provider job ID; this never submits a replacement.'};$j.video_id=$VideoId;$j.stage='HeyGen Processing'}
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
    try{if(-not $c){$c=Get-LFSelectionInputs $Root $p.id $dir};if($c.binding_sha256 -ne $j.authorization){throw 'Authorization stale. Return to Review Selections.'};Assert-LFAvatarJobBinding $c.snapshot $j}catch{$j.stage='Exception';$j.exception_kind='stale authorization';$j.error=$_.Exception.Message;Save-LFAvatarJob $dir $j;continue}
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
