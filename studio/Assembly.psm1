$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'AvatarProduction.psm1')
function Read-LFAssembly($Dir){$file=Join-Path $Dir 'assembly.json';if(Test-Path $file){Read-LFSharedJson $file}else{[pscustomobject]@{schema='lectureforge-assembly-1';current=$null;last_success=$null;history=@()}}}
function Add-LFAssemblyEvent($Attempt,[string]$Stage,[int]$Slide=0,[string]$Message=''){
 if(-not ($Attempt.PSObject.Properties.Name -contains 'events')){$Attempt|Add-Member -NotePropertyName events -NotePropertyValue @()}
 $Attempt.events+=@{utc=[DateTime]::UtcNow.ToString('o');stage=$Stage;slide=$Slide;message=$Message}
}
function Read-LFPlacement($Dir){$file=Join-Path $Dir 'placement.json';if(Test-Path $file){Read-LFSharedJson $file}else{[pscustomobject]@{version=0;overrides=@()}}}
function Assert-LFPlacement($Placement,$Project){
 foreach($key in 'left','top','width','height'){if($null -eq $Placement.$key){throw "Missing placement $key"};$v=[double]$Placement.$key;if([double]::IsNaN($v) -or [double]::IsInfinity($v)){throw 'Placement must be finite.'}}
 if($Placement.left -lt 0 -or $Placement.top -lt 0 -or $Placement.width -le 0 -or $Placement.height -le 0 -or $Placement.width -gt .3 -or $Placement.height -gt .3 -or $Placement.left+$Placement.width -gt 1 -or $Placement.top+$Placement.height -gt 1){throw 'Avatar must be compact and inside the slide (maximum 30% width/height).'}
 $ratio=$Placement.width*$Project.source.width_emu/($Placement.height*$Project.source.height_emu)
 if([math]::Abs($ratio-640.0/560.0) -gt .002){throw 'Preserve the 640:560 avatar aspect ratio.'}
}
function Set-LFAssemblyPlacement($Root,$Id,[int]$Slide,$Placement,[int]$Expected){
 Invoke-LFNarrationLock $Root $Id {param($dir)
  $p=Read-LFProject $Root $Id;$v=Read-LFPlacement $dir
  if($v.version -ne $Expected){throw 'Placement changed in another window. Reload.'}
  if($Slide -notin $p.slides.number){throw 'Unknown slide.'}
  if($null -ne $Placement){Assert-LFPlacement $Placement $p}
  $v.overrides=@($v.overrides|Where-Object slide -ne $Slide)
  if($null -ne $Placement){$v.overrides+=@{slide=$Slide;placement=$Placement}}
  $v.version++;Write-LFJson (Join-Path $dir 'placement.json') $v
 }
}
function Get-LFAssemblyInputs($Root,$Id,$Dir){
 $p=Read-LFProject $Root $Id;$selection=Get-LFSelectionInputs $Root $Id $Dir;$queue=Read-LFAvatarQueue $Dir;$placement=Read-LFPlacement $Dir
 $errors=@();$rows=@();$batch=$queue.batches|Select-Object -Last 1
 if(-not $batch){$errors+='No confirmed production batch.'}
 foreach($s in $selection.snapshot.slides){
  try{
   if($s.status -ne 'narration selected'){throw $s.status}
   $matches=@(foreach($jid in $queue.jobs){$j=Read-LFAvatarJob $Dir $jid;if($j.slide -eq $s.number -and $j.stage -eq 'Avatar Ready' -and -not $j.error -and -not $j.lease -and $j.validation.passed){$j}})
   if(-not $matches.Count){throw 'avatar not ready or unresolved exception'};$j=$matches[-1]
   $null=Assert-LFAvatarArtifactAuthorization $Dir $j
   if((ConvertTo-LFCanonicalJson $j.preset) -ne (ConvertTo-LFCanonicalJson $selection.snapshot.avatar_preset)){throw 'stale avatar generation binding'}
   if($j.row.narration.sha256 -ne $s.narration.sha256 -or $j.validation.narration_sha256 -ne $s.narration.sha256 -or -not $j.validation.authoritative_audio_verified){throw 'stale avatar or audio binding'}
   if((ConvertTo-LFCanonicalJson $j.row.preparation) -ne (ConvertTo-LFCanonicalJson $s.preparation) -or -not $j.validation.pause_verified -or [double]$j.validation.prepared_pause_seconds -ne [double]$s.preparation.post_speech_silence_seconds){throw 'prepared pause binding mismatch'}
   $path=Resolve-LFAvatarPath $Dir $j.output_path
   if(-not (Test-Path $path) -or (Get-FileHash $path).Hash -ne $j.output_sha256 -or $j.validation.output_sha256 -ne $j.output_sha256){throw 'avatar missing or hash mismatch'}
   $override=@($placement.overrides|Where-Object slide -eq $s.number)|Select-Object -First 1
   $pos=if($override){$override.placement}elseif($s.placement_override){$s.placement_override}else{$p.preset.placement};Assert-LFPlacement $pos $p
   $rows+=@{slide=$s.number;take=$s.selected_take;job=$j.id;authorization=$j.authorization;path=$path;sha256=$j.output_sha256;narration=$s.narration;preparation=$s.preparation;placement=$pos;validation=$j.validation}
  }catch{$errors+="Slide $($s.number) - $($_.Exception.Message)"}
 }
 if(-not $selection.snapshot.slides.Count){$errors+='No enabled slides.'}
 $snapshot=@{project=$Id;source_sha256=$p.source.sha256;authorization=$selection.binding_sha256;placement_version=$placement.version;slides=$rows}
 @{ready=($errors.Count -eq 0);errors=$errors;binding=Get-LFTextHash (ConvertTo-LFCanonicalJson $snapshot);snapshot=$snapshot;placement=$placement;default_placement=$p.preset.placement;project_name=$p.name}
}
function Get-LFAssemblyReadiness($Root,$Id){Invoke-LFNarrationLock $Root $Id {param($dir) Get-LFAssemblyInputs $Root $Id $dir}}
function Get-LFAssemblyStatus($Root,$Id){Invoke-LFNarrationLock $Root $Id {param($dir) Read-LFAssembly $dir}}
function Test-LFAssemblyOwner($Lease){if(-not $Lease){return $false};$p=Get-Process -Id $Lease.pid -ErrorAction SilentlyContinue;return ($p -and $p.StartTime.ToUniversalTime().Ticks.ToString() -eq $Lease.started)}
function Request-LFAssembly($Root,$Id,$Expected,$Folder){
 Invoke-LFNarrationLock $Root $Id {param($dir)
  $a=Read-LFAssembly $dir
  if($a.current -and $a.current.stage -notin 'Complete','Exception'){return $a}
  $input=Get-LFAssemblyInputs $Root $Id $dir
  if(-not $input.ready){throw ($input.errors -join '; ')};if($Expected -ne $input.binding){throw 'Assembly inputs changed. Check readiness again.'}
  if(-not $Folder){$Folder=Join-Path $dir 'output'}
  if(-not [IO.Path]::IsPathRooted($Folder) -or $Folder.Contains('"') -or $Folder -match '[\r\n]'){throw 'Choose an absolute output folder.'}
  $folderPath=[IO.Path]::GetFullPath($Folder).TrimEnd('\');if($folderPath.Length -gt 160){throw 'Choose a shorter output folder (160 characters maximum).'}
  if($a.current -and $a.current.stage -eq 'Complete' -and $a.current.binding -eq $Expected -and $a.current.folder -eq $folderPath -and (Test-Path $a.current.output) -and (Get-FileHash $a.current.output).Hash -eq $a.current.output_sha256){return $a}
  $attempt=[guid]::NewGuid().ToString('N').Substring(0,8);$attemptDir=Join-Path $dir "r/$attempt";[void][IO.Directory]::CreateDirectory($attemptDir)
  Write-LFImmutableJson (Join-Path $attemptDir 'input.json') $input.snapshot
  if($a.current){$a.history+=,$a.current}
  $name=[regex]::Replace($input.project_name,'[^\p{L}\p{N} _-]','_').Trim();if(-not $name){$name='Lecture'};if($name.Length -gt 55){$name=$name.Substring(0,55)}
  $a.current=[pscustomobject]@{id=$attempt;binding=$Expected;stage='Queued';slide=0;total=$input.snapshot.slides.Count;folder=$folderPath;filename=($name+'_RECORDING.pptx');output=$null;output_sha256=$null;error=$null;lease=$null;events=@();created_utc=[DateTime]::UtcNow.ToString('o')}
  Add-LFAssemblyEvent $a.current 'Readiness validated' 0 ($input.snapshot.slides.Count.ToString()+' slides ready')
  Add-LFAssemblyEvent $a.current 'Assembly request accepted'
  Write-LFJson (Join-Path $dir 'assembly.json') $a;$a
 }
}
function Start-LFAssemblyTask($Root){
 foreach($p in Get-LFProjects $Root){
  $task=Invoke-LFNarrationLock $Root $p.id {param($dir)
   $a=Read-LFAssembly $dir;$c=$a.current;if(-not $c -or $c.stage -in 'Complete','Exception'){return}
   if($c.lease){if(Test-LFAssemblyOwner $c.lease){return};$c.stage='Exception';$c.error='Local assembly interrupted. Retry Assembly creates a fresh derivative; prior outputs are preserved.';$c.lease=$null;Write-LFJson (Join-Path $dir 'assembly.json') $a;return}
   if($c.stage -ne 'Queued'){return};$c.lease=@{pid=$PID;started=(Get-Process -Id $PID).StartTime.ToUniversalTime().Ticks.ToString()};Write-LFJson (Join-Path $dir 'assembly.json') $a
   @{id=$p.id;attempt=$c.id}
  }
  if($task){return $task}
 }
}
function Open-LFRecording($Root,$Id,$Kind){
 $state=Get-LFAssemblyStatus $Root $Id;$s=$state.last_success
 if(-not $s -or -not (Test-Path $s.output) -or (Get-FileHash $s.output).Hash -ne $s.output_sha256){throw 'No intact validated Recording PowerPoint is available.'}
 if($Kind -eq 'deck'){Start-Process -FilePath $s.output}elseif($Kind -eq 'folder'){Start-Process explorer.exe -ArgumentList ('"'+(Split-Path $s.output)+'"')}else{throw 'Unknown open action.'}
 @{opened=$true;path=$s.output}
}
Export-ModuleMember -Function *-LF*
