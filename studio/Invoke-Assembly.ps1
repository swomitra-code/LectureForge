param([string]$RuntimeRoot,[string]$Id,[string]$Attempt)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Assembly.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Recording.psm1') -Force
$dir=Get-LFProjectPath $RuntimeRoot $Id;$attemptDir=Join-Path $dir "r/$Attempt"
$mutex=[Threading.Mutex]::new($false,'Local\LectureForge-PowerPoint-Assembly');$held=$false
function Progress($Stage,$Slide){Invoke-LFNarrationLock $RuntimeRoot $Id {param($d) $a=Read-LFAssembly $d;if($a.current.id -ne $Attempt){throw 'Assembly attempt replaced.'};$a.current.stage=$Stage;$a.current.slide=$Slide;Add-LFAssemblyEvent $a.current $Stage $Slide;Write-LFJson (Join-Path $d 'assembly.json') $a}}
try{
 Invoke-LFNarrationLock $RuntimeRoot $Id {param($d) $a=Read-LFAssembly $d;$a.current.lease=@{pid=$PID;started=(Get-Process -Id $PID).StartTime.ToUniversalTime().Ticks.ToString()};Write-LFJson (Join-Path $d 'assembly.json') $a}
 Progress 'Waiting for PowerPoint' 0
 try{$held=$mutex.WaitOne(0)}catch [Threading.AbandonedMutexException]{$held=$true}
 if(-not $held){throw 'Another local PowerPoint assembly is active. Retry after it completes.'}
 Progress 'Preparing deck' 0
 $state=Get-LFAssemblyStatus $RuntimeRoot $Id;$manifest=Read-LFSharedJson (Join-Path $attemptDir 'input.json');$inputs=Get-LFAssemblyReadiness $RuntimeRoot $Id
 if(-not $inputs.ready -or $inputs.binding -ne $state.current.binding -or (Get-LFTextHash (ConvertTo-LFCanonicalJson $manifest)) -ne $inputs.binding){throw 'Assembly inputs stale. Check readiness again.'}
 Progress 'Launching PowerPoint COM' 0
 $result=New-V2RecordingBatch -Source (Join-Path $dir 'source.pptx') -SourceSha256 $manifest.source_sha256 -Output (Join-Path $attemptDir 'deck.pptx') -EvidenceDirectory (Join-Path $attemptDir 'v') -Avatars $manifest.slides -Progress {param($stage,$slide) Progress $stage $slide}
 if(-not $result.passed){throw 'Deck validation failed.'}
 # Recheck exact binding at publication. No long project lock during COM or decoding.
 Invoke-LFNarrationLock $RuntimeRoot $Id {param($d)
  $now=Get-LFAssemblyInputs $RuntimeRoot $Id $d;$a=Read-LFAssembly $d
  if(-not $now.ready -or $now.binding -ne $a.current.binding){throw 'Inputs changed during assembly. Validated working deck preserved; check readiness and retry.'}
  [void][IO.Directory]::CreateDirectory($a.current.folder)
  $dest=Join-Path $a.current.folder $a.current.filename
  if(Test-Path $dest){$dest=Join-Path $a.current.folder ([IO.Path]::GetFileNameWithoutExtension($a.current.filename)+'_'+$Attempt+'.pptx')}
  # Copy to a unique temporary file; publish atomically without replacing any output.
  $tmp=Join-Path $a.current.folder ($Attempt+'.partial')
  [IO.File]::Copy((Join-Path $attemptDir 'deck.pptx'),$tmp,$false)
  if((Get-FileHash $tmp).Hash -ne $result.output_sha256){throw 'Export copy hash mismatch.'}
  [IO.File]::Move($tmp,$dest)
  $a.current.output=$dest;$a.current.output_sha256=$result.output_sha256;$a.current.stage='Complete';$a.current.lease=$null;Add-LFAssemblyEvent $a.current 'Complete' 0 $dest
  $a.last_success=$a.current.PSObject.Copy();Write-LFJson (Join-Path $d 'assembly.json') $a
 }
}catch{
 $message=$_.Exception.Message;foreach($key in @($env:ELEVENLABS_API_KEY,$env:HEYGEN_API_KEY)){if($key){$message=$message.Replace($key,'[REDACTED]')}}
 Invoke-LFNarrationLock $RuntimeRoot $Id {param($d) $a=Read-LFAssembly $d;if($a.current.id -eq $Attempt){$a.current.error=$message;$a.current.stage='Exception';$a.current.lease=$null;Add-LFAssemblyEvent $a.current 'Exception' $a.current.slide $message;Write-LFJson (Join-Path $d 'assembly.json') $a}}
 exit 1
}finally{if($held){$mutex.ReleaseMutex()};$mutex.Dispose()}
