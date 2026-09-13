$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Narration.psm1')

function ConvertTo-LFCanonicalJson($Value){
 if($null -eq $Value){return 'null'}
 if($Value -is [Collections.IDictionary]){
  $parts=@(foreach($key in @($Value.Keys|Sort-Object)){(ConvertTo-Json ([string]$key) -Compress)+':'+(ConvertTo-LFCanonicalJson $Value[$key])});return '{'+($parts -join ',')+'}'
 }
 if($Value.GetType().FullName -eq 'System.Management.Automation.PSCustomObject'){
  $parts=@(foreach($key in @($Value.PSObject.Properties.Name|Sort-Object)){(ConvertTo-Json ([string]$key) -Compress)+':'+(ConvertTo-LFCanonicalJson $Value.$key)});return '{'+($parts -join ',')+'}'
 }
 if($Value -is [Collections.IEnumerable] -and $Value -isnot [string]){return '['+(@(foreach($item in $Value){ConvertTo-LFCanonicalJson $item}) -join ',')+']'}
 ConvertTo-Json -InputObject $Value -Compress
}
function Get-LFAvatarPreset($Project){
 [ordered]@{id=$Project.preset.id;version=$Project.preset.version;avatar=$Project.preset.avatar;encoding=$Project.preset.encoding;authoritative_audio=$Project.preset.recording.authoritative_selected_narration}
}
function Resolve-LFReusableAsset([string]$Dir,[string]$Relative){
 if($Relative -notmatch '^avatars/[a-zA-Z0-9_-]{1,40}\.mp4$'){throw 'Invalid reusable avatar path.'}
 $path=[IO.Path]::GetFullPath((Join-Path $Dir $Relative))
 if($path.Length -ge 250 -or -not $path.StartsWith(([IO.Path]::GetFullPath($Dir).TrimEnd('\')+'\'),[StringComparison]::OrdinalIgnoreCase)){throw 'Unsafe reusable avatar path.'}
 $path
}
function Get-LFSelectionInputs([string]$Root,[string]$Id,[string]$Dir){
 $p=Read-LFProject $Root $Id;$n=Read-LFNarrationFile $Dir
 $preset=Get-LFAvatarPreset $p;$presetHash=Get-LFTextHash (ConvertTo-LFCanonicalJson $preset)
 $registryFile=Join-Path $Dir 'avatars.json';$registry=@();$registryHash=$null
 if(Test-Path $registryFile){$registryHash=(Get-FileHash $registryFile).Hash;$registry=@((Get-Content -Raw -Encoding UTF8 $registryFile|ConvertFrom-Json).assets)}
 $rows=@(foreach($s in $p.slides|Where-Object enabled){
  $r=Get-LFCurrentRevision $n $p $s
  $old=@($n.revisions|Where-Object {$_.slide_number -eq $s.number -and $_.selected_take})|Select-Object -Last 1
  $status='missing narration selection';$take=$null;$asset=$null;$reuse=$null;$preparation=$null
  if(-not $r -and $old){$status=if($old.script_sha256 -ne (Get-LFTextHash $s.script)){'script changed after selection'}else{'stale narration selection'}}
  if($r -and $r.selected_take){
   $take=$r.selected_take;$t=@($r.takes|Where-Object number -eq $take)[0];$status='selected asset missing/invalid'
   try{
    if(-not $t -or -not (Test-LFTake $Dir $t) -or -not $t.asset.decode_ok -or [double]$t.asset.duration_seconds -le 0 -or [double]::IsNaN([double]$t.asset.duration_seconds) -or [double]::IsInfinity([double]$t.asset.duration_seconds)){throw 'Invalid take.'}
    if($r.script_sha256 -ne (Get-LFTextHash $s.script)){throw 'Script hash mismatch.'}
    if([double]$t.asset.post_speech_silence_seconds -ne [double]$r.pause_seconds){throw 'Prepared pause mismatch.'}
    if($r.pause_seconds -gt 0 -and ([long]$t.asset.pause_samples -ne [long][math]::Round($r.pause_seconds*$t.asset.pause_sample_rate))){throw 'Pause samples mismatch.'}
    $asset=[ordered]@{id=$r.id+'/'+$take+'/'+$t.attempts[-1].id;path=$t.asset.path;sha256=$t.asset.sha256;duration_seconds=$t.asset.duration_seconds;decode_ok=$true}
    $preparation=[ordered]@{post_speech_silence_seconds=$r.pause_seconds;pause_samples=$t.asset.pause_samples;sample_rate=$t.asset.pause_sample_rate}
    $status='narration selected'
    foreach($candidate in $registry){
     if($candidate.slide_number -ne $s.number -or $candidate.status -ne 'approved' -or $candidate.rejected -eq $true -or $candidate.narration_sha256 -ne $asset.sha256 -or $candidate.preset_sha256 -ne $presetHash -or $candidate.validation.passed -ne $true -or $candidate.validation.authoritative_audio_verified -ne $true){continue}
     try{
      $path=Resolve-LFReusableAsset $Dir $candidate.path
      if((Test-Path $path) -and (Get-FileHash $path).Hash -eq $candidate.sha256){$reuse=[ordered]@{id=$candidate.id;path=$candidate.path;sha256=$candidate.sha256;narration_sha256=$candidate.narration_sha256;preset_sha256=$candidate.preset_sha256};break}
     }catch{ } # Invalid historical assets cannot earn reuse credit.
    }
   }catch{$asset=$null;$preparation=$null;$status='selected asset missing/invalid'}
  }
  [ordered]@{slide_id=$Id+'/'+$s.number;number=$s.number;title=$s.title;selected_take=$take;status=$status;narration=$asset;script_revision=if($r){$r.script_revision}else{$null};script_sha256=Get-LFTextHash $s.script;narration_revision=if($r){$r.id}else{$null};preparation=$preparation;placement_override=$s.placement_override;reusable_avatar=$reuse;new_jobs=if($asset -and -not $reuse){1}else{0}}
 })
 $valid=$rows.Count -gt 0 -and @($rows|Where-Object {$_.status -ne 'narration selected'}).Count -eq 0
 $input=[ordered]@{project_id=$Id;project_revision=$p.version;intent_revision=if($n.PSObject.Properties.Name -contains 'intent_revision'){$n.intent_revision}else{0};source_sha256=$p.source.sha256;enabled_slide_set=@($rows.number);excluded_slides=@($p.slides|Where-Object {-not $_.enabled}|ForEach-Object number);avatar_preset=$preset;avatar_preset_sha256=$presetHash;default_placement=$p.preset.placement;recording_policy=$p.preset.recording;registry_sha256=$registryHash;slides=$rows;can_authorize=$valid;expected_new_provider_jobs=@($rows|Where-Object {$_.new_jobs -eq 1}).Count;reusable_avatars=@($rows|Where-Object reusable_avatar).Count}
 @{snapshot=$input;binding_sha256=Get-LFTextHash (ConvertTo-LFCanonicalJson $input)}
}
function Write-LFImmutableJson([string]$Path,$Value){
 $text=ConvertTo-Json -InputObject $Value -Depth 30
 $temp=$Path+'.'+[guid]::NewGuid().ToString('N').Substring(0,8)+'.tmp'
 [IO.File]::WriteAllText($temp,$text,[Text.UTF8Encoding]::new($false))
 # Move is atomic and fails if the destination already exists. Never overwrite approvals.
 [IO.File]::Move($temp,$Path)
}
function Get-LFApprovalPath([string]$Dir,[string]$Id,[string]$Kind){
 if($Id -notmatch '^[A-F0-9]{64}$' -or $Kind -notin @('review','authorization','consumed')){throw 'Invalid approval identifier.'}
 Join-Path $Dir "approvals/$Id.$Kind.json"
}
function Read-LFApproval([string]$Dir,[string]$Id,[string]$Kind){
 $path=Get-LFApprovalPath $Dir $Id $Kind
 $record=Get-Content -Raw -Encoding UTF8 $path|ConvertFrom-Json
 if($record.id -ne $Id -or $record.binding_sha256 -ne (Get-LFTextHash (ConvertTo-LFCanonicalJson $record.snapshot)) -or $record.binding_sha256 -ne $Id){throw 'Approval integrity check failed.'}
 $record
}
function Get-LFAuthorizationStatus([string]$Dir,[string]$Id,$Current){
 $path=Get-LFApprovalPath $Dir $Id 'authorization'
 if(-not (Test-Path $path)){return 'not authorized'}
 $null=Read-LFApproval $Dir $Id 'authorization'
 if(Test-Path (Get-LFApprovalPath $Dir $Id 'consumed')){return 'consumed'}
 if($Id -ne $Current.binding_sha256 -or -not $Current.snapshot.can_authorize){return 'stale'}
 'authorized'
}
function Get-LFSelectionReview([string]$Root,[string]$Id){
 Invoke-LFNarrationLock $Root $Id {param($dir)
  $current=Get-LFSelectionInputs $Root $Id $dir;$saved=$null;$review=$null;$authorization=$null;$status='not authorized'
  $file=Join-Path $dir 'selection-review.json'
  if(Test-Path $file){
   $saved=Get-Content -Raw -Encoding UTF8 $file|ConvertFrom-Json
   if($saved.review_id){
    $review=Read-LFApproval $dir $saved.review_id 'review'
    $status=Get-LFAuthorizationStatus $dir $saved.review_id $current
    if(Test-Path (Get-LFApprovalPath $dir $saved.review_id 'authorization')){$authorization=Read-LFApproval $dir $saved.review_id 'authorization'}
   }
  }
  @{current=$current;saved=$saved;review=$review;authorization=$authorization;authorization_status=$status;review_stale=($review -and $review.id -ne $current.binding_sha256);provider_submission_enabled=$false}
 }
}
function Save-LFSelectionReview([string]$Root,[string]$Id,[string]$Expected,[string]$Phase='review'){
 Invoke-LFNarrationLock $Root $Id {param($dir)
  if($Phase -notin @('review','confirmation')){throw 'Invalid review phase.'}
  $current=Get-LFSelectionInputs $Root $Id $dir
  if($current.binding_sha256 -ne $Expected){throw 'Selections changed in another window. Review again.'}
  if($Phase -eq 'confirmation' -and -not $current.snapshot.can_authorize){throw 'Resolve missing or stale selections first.'}
  [void][IO.Directory]::CreateDirectory((Join-Path $dir 'approvals'))
  $rid=$current.binding_sha256;$file=Get-LFApprovalPath $dir $rid 'review'
  if(-not (Test-Path $file)){Write-LFImmutableJson $file @{schema='lectureforge-selection-review-1';id=$rid;timestamp=[DateTime]::UtcNow.ToString('o');binding_sha256=$rid;snapshot=$current.snapshot}}
  Write-LFJson (Join-Path $dir 'selection-review.json') @{review_id=$rid;phase=$Phase;slide=$null}
  Read-LFApproval $dir $rid 'review'
 }
}
function Save-LFReviewNavigation([string]$Root,[string]$Id,[int]$Slide){
 Invoke-LFNarrationLock $Root $Id {param($dir)
  $p=Read-LFProject $Root $Id
  if($Slide -notin @($p.slides.number)){throw 'Unknown slide.'}
  $file=Join-Path $dir 'selection-review.json';$saved=@{review_id=$null;phase='narration';slide=$Slide}
  if(Test-Path $file){$saved=Get-Content -Raw $file|ConvertFrom-Json;$saved.phase='narration';$saved.slide=$Slide}
  Write-LFJson $file $saved
 }
}
function Confirm-LFAvatarAuthorization([string]$Root,[string]$Id,[string]$ReviewId,[string]$Expected){
 Invoke-LFNarrationLock $Root $Id {param($dir)
  $current=Get-LFSelectionInputs $Root $Id $dir
  if($Expected -ne $ReviewId -or $current.binding_sha256 -ne $ReviewId -or -not $current.snapshot.can_authorize){throw 'Stale or incomplete summary. Review selections again.'}
  $review=Read-LFApproval $dir $ReviewId 'review'
  $saved=Get-Content -Raw (Join-Path $dir 'selection-review.json')|ConvertFrom-Json
  if($saved.review_id -ne $ReviewId -or $saved.phase -notin @('confirmation','authorized')){throw 'Open the final confirmation summary first.'}
  if(Test-Path (Get-LFApprovalPath $dir $ReviewId 'consumed')){throw 'Authorization already consumed. Resubmission is blocked.'}
  $path=Get-LFApprovalPath $dir $ReviewId 'authorization'
  if(-not (Test-Path $path)){
   Write-LFImmutableJson $path @{schema='lectureforge-avatar-authorization-1';id=$ReviewId;timestamp=[DateTime]::UtcNow.ToString('o');project_id=$Id;project_revision=$current.snapshot.project_revision;binding_sha256=$ReviewId;snapshot=$review.snapshot;authorization='paid avatar generation for these exact selections';provider_jobs_submitted=0;consumer_milestone='D'}
  }
  $record=Read-LFApproval $dir $ReviewId 'authorization'
  Write-LFJson (Join-Path $dir 'selection-review.json') @{review_id=$ReviewId;phase='authorized';slide=$null}
  @{record=$record;file_sha256=(Get-FileHash $path).Hash;status='authorized';provider_jobs_submitted=0}
 }
}
# Milestone D must validate exact inputs and atomically create its consumed receipt
# under the same project mutex BEFORE any request. This milestone has no consumer.
Export-ModuleMember -Function *-LF*
