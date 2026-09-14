$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Production.psm1')
function Get-LFTextHash([string]$Text){
 $h=[Security.Cryptography.SHA256]::Create();try{([BitConverter]::ToString($h.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-','')}finally{$h.Dispose()}
}
function Get-LFBinding($Project,$Slide){
 Get-LFTextHash ($Slide.script+"`n"+($Slide.post_speech_silence_seconds.ToString([Globalization.CultureInfo]::InvariantCulture))+"`n"+($Project.preset.narration|ConvertTo-Json -Compress))
}
function Invoke-LFNarrationLock([string]$Root,[string]$Id,[scriptblock]$Action){
 $dir=Get-LFProjectPath $Root $Id
 $mutex=[Threading.Mutex]::new($false,('Local\LFN-'+(Get-LFTextHash $dir).Substring(0,24)))
 $held=$false
 # Full-lecture asset verification can hold this lock beyond ten seconds on local disks.
 # Allow bounded contention between the three workers and UI; never retry provider calls.
 try{try{$held=$mutex.WaitOne(60000)}catch [Threading.AbandonedMutexException]{$held=$true};if(-not $held){throw 'Narration state busy. Try again.'}; & $Action $dir}finally{if($held){$mutex.ReleaseMutex()};$mutex.Dispose()}
}
function Read-LFNarrationFile([string]$Dir){
 $file=Join-Path $Dir 'narration.json'
 $state=if(Test-Path $file){Get-Content -Raw -Encoding UTF8 $file|ConvertFrom-Json}else{[pscustomobject]@{schema='lectureforge-narration-1';revisions=@()}}
 foreach($r in $state.revisions){
  # Read-time migration only. Import never rewrites narration history or queues work.
  $raw=$r.takes
  $items=@(if($raw -is [array]){$raw}elseif($null -ne $raw){
   if($raw.PSObject.Properties.Name -contains 'number'){$raw}
   elseif($raw -is [pscustomobject]){$raw.PSObject.Properties.Value}
   else{$raw}
  })
  $valid=@();$legacy=@()
  foreach($t in $items){
   if($null -ne $t -and $t -is [pscustomobject] -and $t.number -in 1,2,3 -and $t.state -is [string]){$valid+= $t}
   elseif($null -ne $t){$legacy+= $t}
  }
  # Preserve unrecognized data for recovery, including on later narration writes.
  if($legacy.Count){
   $r|Add-Member -NotePropertyName takes_legacy -NotePropertyValue (@($r.takes_legacy)+@($raw) | Where-Object {$null -ne $_}) -Force
  }
  $r|Add-Member -NotePropertyName takes -NotePropertyValue $valid -Force
 }
 $state
}
function Get-LFCurrentRevision($State,$Project,$Slide){
 $binding=Get-LFBinding $Project $Slide
 @($State.revisions|Where-Object {$_.slide_number -eq $Slide.number -and $_.binding -eq $binding})|Select-Object -Last 1
}
function Resolve-LFNarrationAsset([string]$Dir,[string]$Relative){
 if($Relative -notmatch '^n/[a-f0-9]{12}/[a-f0-9]{12}/(raw\.mp3|take\.(mp3|wav))$'){throw 'Invalid narration asset path.'}
 $path=[IO.Path]::GetFullPath((Join-Path $Dir $Relative))
 if(-not $path.StartsWith(([IO.Path]::GetFullPath($Dir).TrimEnd('\')+'\'),[StringComparison]::OrdinalIgnoreCase) -or $path.Length -ge 250){throw 'Unsafe narration path.'};$path
}
function Test-LFTake($Dir,$Take){
 if($Take.state -ne 'ready'){return $false}
 $path=Resolve-LFNarrationAsset $Dir $Take.asset.path
 (Test-Path $path) -and (Get-FileHash $path -Algorithm SHA256).Hash -eq $Take.asset.sha256
}
function Get-LFNarration([string]$Root,[string]$Id){
 Invoke-LFNarrationLock $Root $Id {param($dir)
  $p=Read-LFProject $Root $Id;$state=Read-LFNarrationFile $dir
  $slides=@(foreach($s in $p.slides){
   $r=Get-LFCurrentRevision $state $p $s
   # The array subexpression must surround the conditional: PowerShell otherwise
   # unwraps zero/one items (Windows PowerShell serializes the empty result as {}).
   [pscustomobject]@{number=$s.number;enabled=$s.enabled;revision=if($r){$r.id}else{$null};selected_take=if($r){$r.selected_take}else{$null};takes=@(if($r){$r.takes});binding=Get-LFBinding $p $s}
  })
  @{slides=$slides;project_version=$p.version;provider_calls=@($state.revisions.takes.attempts|Where-Object {$_.submitted_utc}).Count;heygen_calls=0}
 }
}
function Get-LFNarrationPlan([string]$Root,[string]$Id){
 Invoke-LFNarrationLock $Root $Id {param($dir)
  $p=Read-LFProject $Root $Id;$state=Read-LFNarrationFile $dir;$count=0;$slides=0
  foreach($s in $p.slides|Where-Object enabled){
   if([string]::IsNullOrWhiteSpace($s.script)){throw "Slide $($s.number) needs a saved script."}
   $r=Get-LFCurrentRevision $state $p $s
   if(-not $r){$count+=3;$slides++}else{
    foreach($t in $r.takes|Where-Object state -eq 'ready'){if(-not (Test-LFTake $dir $t)){throw "Slide $($s.number) Take $($t.number) asset missing or changed. No regeneration authorized."}}
   }
  }
  @{slides=$slides;generations=$count;version=$p.version}
 }
}
function New-LFAttempt($Revision,$Take){
 $aid=[guid]::NewGuid().ToString('N').Substring(0,12)
 [pscustomobject]@{id=$aid;state='queued';queued_utc=[DateTime]::UtcNow.ToString('o');submitted_utc=$null;finished_utc=$null;owner_pid=0;owner_start=$null;error=$null;raw_path="n/$($Revision.id)/$aid/raw.mp3";provider='ElevenLabs'}
}
function Add-LFNarrationBatch([string]$Root,[string]$Id,[int]$Version){
 Invoke-LFNarrationLock $Root $Id {param($dir)
  $p=Read-LFProject $Root $Id;if($p.version -ne $Version){throw 'Scripts changed. Review generation totals again.'}
  $state=Read-LFNarrationFile $dir
  foreach($s in $p.slides|Where-Object enabled){if([string]::IsNullOrWhiteSpace($s.script)){throw "Slide $($s.number) needs a script."}}
  foreach($s in $p.slides|Where-Object enabled){
   if(Get-LFCurrentRevision $state $p $s){continue}
   $r=[pscustomobject]@{id=[guid]::NewGuid().ToString('N').Substring(0,12);slide_id="$Id/$($s.number)";slide_number=$s.number;script=$s.script;script_sha256=Get-LFTextHash $s.script;script_revision=$p.version;binding=Get-LFBinding $p $s;pause_seconds=$s.post_speech_silence_seconds;settings=$p.preset.narration;created_utc=[DateTime]::UtcNow.ToString('o');selected_take=$null;takes=@()}
   foreach($n in 1..3){$t=[pscustomobject]@{number=$n;state='queued';asset=$null;attempts=@()};$t.attempts=@(New-LFAttempt $r $t);$r.takes+= $t}
   $state.revisions+= $r
  }
  Write-LFJson (Join-Path $dir 'narration.json') $state
 }
}
function Set-LFNarrationSelection([string]$Root,[string]$Id,[int]$Slide,[string]$Revision,[int]$Take){
 Invoke-LFNarrationLock $Root $Id {param($dir)
  $p=Read-LFProject $Root $Id;$state=Read-LFNarrationFile $dir;$s=@($p.slides|Where-Object number -eq $Slide)[0];$r=Get-LFCurrentRevision $state $p $s
  if(-not $r -or $r.id -ne $Revision -or $Take -notin 0,1,2,3){throw 'Narration revision changed. Reload this slide.'}
  if($Take -ne 0 -and -not (Test-LFTake $dir (@($r.takes|Where-Object number -eq $Take)|Select-Object -First 1))){throw 'Take is not verified and ready.'}
  $selected=if($Take){$Take}else{$null}
  if($r.selected_take -ne $selected){
   if($state.PSObject.Properties.Name -notcontains 'intent_revision'){$state|Add-Member -NotePropertyName intent_revision -NotePropertyValue 0}
   $state.intent_revision++
  }
  $r.selected_take=$selected;Write-LFJson (Join-Path $dir 'narration.json') $state
 }
}
function Retry-LFNarration([string]$Root,[string]$Id,[string]$Revision,[int]$Take){
 Invoke-LFNarrationLock $Root $Id {param($dir)
  $state=Read-LFNarrationFile $dir;$p=Read-LFProject $Root $Id;$r=@($state.revisions|Where-Object id -eq $Revision)[0]
  if(-not $r -or $Take -notin 1,2,3){throw 'Unknown take.'}
  $s=@($p.slides|Where-Object number -eq $r.slide_number)[0]
  if((Get-LFCurrentRevision $state $p $s).id -ne $Revision){throw 'Cannot retry an obsolete narration revision.'}
  $t=@($r.takes|Where-Object number -eq $Take)|Select-Object -First 1
  if($t.state -ne 'failed'){throw 'Only a confirmed rejected take can be retried. Uncertain attempts require investigation.'}
  $t.attempts+= (New-LFAttempt $r $t);$t.state='queued';Write-LFJson (Join-Path $dir 'narration.json') $state
 }
}
function Get-LFAudioValidation([string]$Path){
 & ffmpeg -v error -i $Path -f null NUL 2>$null
 if($LASTEXITCODE -ne 0){throw 'Narration decode failed.'}
 $raw=(& ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 $Path) -join ''
 if($LASTEXITCODE -ne 0){throw 'Narration duration probe failed.'}
 $duration=[double]::Parse($raw,[Globalization.CultureInfo]::InvariantCulture)
 if($duration -le 0){throw 'Narration duration is zero.'}
 @{sha256=(Get-FileHash $Path -Algorithm SHA256).Hash;duration_seconds=$duration;decode_ok=$true}
}
function Complete-LFNarrationAudio([string]$Dir,$Revision,$Attempt){
 $raw=Resolve-LFNarrationAsset $Dir $Attempt.raw_path
 $null=Get-LFAudioValidation $raw
 $ext=if($Revision.pause_seconds -gt 0){'wav'}else{'mp3'}
 $relative=$Attempt.raw_path.Replace('raw.mp3',"take.$ext");$dest=Resolve-LFNarrationAsset $Dir $relative
 if(Test-Path $dest){throw 'Prepared candidate already exists; preserve and investigate.'}
 $samples=[long][math]::Round([double]$Revision.pause_seconds*44100)
 if($samples -gt 0){
  # PCM WAV avoids MP3 encoder padding. apad adds exactly N zero samples after decoded speech.
  & ffmpeg -v error -n -i $raw -ar 44100 -ac 1 -af "aresample=44100,apad=pad_len=$samples" -c:a pcm_s16le $dest 2>$null
  if($LASTEXITCODE -ne 0){throw 'Prepared silence conversion failed.'}
 }else{[IO.File]::Copy($raw,$dest,$false)}
 $asset=Get-LFAudioValidation $dest;$asset.path=$relative;$asset.pause_samples=$samples;$asset.pause_sample_rate=44100;$asset.post_speech_silence_seconds=$samples/44100.0
 $asset
}
function Get-LFNextNarration([string]$Root){
 foreach($project in Get-LFProjects $Root){
  $id=$project.id
  $task=Invoke-LFNarrationLock $Root $id {param($dir)
   $state=Read-LFNarrationFile $dir;$changed=$false;$active=$false;$p=Read-LFProject $Root $id
   foreach($r in $state.revisions){foreach($t in $r.takes){
    $a=$t.attempts[-1]
    if($t.state -in @('dispatched','generating')){
     $owner=Get-Process -Id $a.owner_pid -ErrorAction SilentlyContinue
     if($owner -and $a.owner_start -and $owner.StartTime.ToUniversalTime().ToString('o') -eq $a.owner_start){$active=$true}else{$t.state='uncertain';$a.state='uncertain';$a.error='Interrupted attempt. No automatic paid retry. Preserve all existing files.';$changed=$true}
    }
   }}
   if($changed){Write-LFJson (Join-Path $dir 'narration.json') $state}
   if($active){return @{busy=$true}}
   foreach($s in $p.slides|Where-Object enabled){
    $r=Get-LFCurrentRevision $state $p $s;if(-not $r){continue}
    foreach($t in $r.takes|Where-Object state -eq 'queued'){
     $a=$t.attempts[-1];$t.state='dispatched';$a.state='dispatched';$a.owner_pid=$PID;$a.owner_start=(Get-Process -Id $PID).StartTime.ToUniversalTime().ToString('o')
     Write-LFJson (Join-Path $dir 'narration.json') $state
     return @{id=$id;revision=$r.id;take=$t.number;attempt=$a.id;busy=$false}
    }
   }
  }
  if($task){return $task}
 }
}
Export-ModuleMember -Function *-LF*
