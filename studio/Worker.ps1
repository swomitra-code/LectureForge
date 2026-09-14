param([Parameter(Mandatory=$true)][string]$RuntimeRoot,[Parameter(Mandatory=$true)][string]$Instance)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Production.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Narration.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'AvatarProduction.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Assembly.psm1') -Force
$assemblyChild=$null
$root=Get-LFRoot $RuntimeRoot
$stop=Join-Path $root "$Instance.stop"
$heartbeat=Join-Path $root "$Instance.worker.json"
$avatarChildren=@()
$child=$null;$task=$null
try{
 while(-not (Test-Path -LiteralPath $stop)){
  $runtime=Join-Path $root 'runtime.json'
  if(Test-Path $runtime){
   $owner=Read-LFSharedJson $runtime
   if($owner.instance -eq $Instance -and -not (Get-Process -Id $owner.server_pid -ErrorAction SilentlyContinue)){break}
  }
  if($assemblyChild -and $assemblyChild.HasExited){$assemblyChild=$null}
  if(-not $assemblyChild -and -not $child){
   $at=Start-LFAssemblyTask $root
   if($at){
    $args='-NoProfile -ExecutionPolicy Bypass -File "'+(Join-Path $PSScriptRoot 'Invoke-Assembly.ps1')+'" -RuntimeRoot "'+$root+'" -Id '+$at.id+' -Attempt '+$at.attempt
    $log=Join-Path (Get-LFProjectPath $root $at.id) ('r/'+$at.attempt+'/executor')
    $assemblyChild=Start-Process (Join-Path $PSHOME 'powershell.exe') -PassThru -RedirectStandardOutput ($log+'.log') -RedirectStandardError ($log+'.err') -ArgumentList $args
    Invoke-LFNarrationLock $root $at.id {param($d) $a=Read-LFAssembly $d;if($a.current.lease){$a.current.lease=@{pid=$assemblyChild.Id;started=$assemblyChild.StartTime.ToUniversalTime().Ticks.ToString()};Write-LFJson (Join-Path $d 'assembly.json') $a}}
   }
  }
  $avatarChildren=@($avatarChildren|Where-Object {-not $_.HasExited})
  foreach($project in Get-LFProjects $root){try{Start-LFAvatarBatch $root $project.id}catch{[Console]::Error.WriteLine('Avatar authorization exception for project '+$project.id+': '+$_.Exception.Message)}}
  for($slot=$avatarChildren.Count;$slot -lt 3;$slot++){
   $next=@(Get-LFAvatarTasks $root $Instance 1)|Select-Object -First 1
   if(-not $next){break}
   $args='-NoProfile -ExecutionPolicy Bypass -File "'+(Join-Path $PSScriptRoot 'Invoke-AvatarStep.ps1')+'" -RuntimeRoot "'+$root+'" -Id '+$next.id+' -Job '+$next.job
   $log=Join-Path (Get-LFProjectPath $root $next.id) ('a/'+$next.job+'/run-'+[guid]::NewGuid().ToString('N').Substring(0,8))
   $av=Start-Process (Join-Path $PSHOME 'powershell.exe') -PassThru -RedirectStandardOutput ($log+'.log') -RedirectStandardError ($log+'.err') -ArgumentList $args
   Invoke-LFNarrationLock $root $next.id {param($dir) $j=Read-LFAvatarJob $dir $next.job;if($j.lease){$j.lease.pid=$av.Id;$j.lease.started=$av.StartTime.ToUniversalTime().Ticks.ToString();Save-LFAvatarJob $dir $j}}
   $avatarChildren+=,$av
  }
  if($child -and $child.HasExited){
   if($task -and $task.ContainsKey('attempt')){
    Invoke-LFNarrationLock $root $task.id {param($dir)
     $state=Read-LFNarrationFile $dir;$r=@($state.revisions|Where-Object id -eq $task.revision)[0];$t=@($r.takes|Where-Object number -eq $task.take)|Select-Object -First 1
     if($t.state -in @('dispatched','generating')){$t.state='uncertain';$t.attempts[-1].state='uncertain';$t.attempts[-1].error='Executor exited without a final result. No automatic retry.';Write-LFJson (Join-Path $dir 'narration.json') $state}
    }
   }
   $child=$null;$task=$null
  }
  if(-not $child -and -not $assemblyChild){
   $task=Get-LFNextNarration $root
   if($task -and -not $task.busy){
    $args='-NoProfile -ExecutionPolicy Bypass -File "'+(Join-Path $PSScriptRoot 'Invoke-NarrationTake.ps1')+'" -RuntimeRoot "'+$root+'" -Id '+$task.id+' -Revision '+$task.revision+' -Take '+$task.take+' -Attempt '+$task.attempt
    $child=Start-Process (Join-Path $PSHOME 'powershell.exe') -PassThru -ArgumentList $args
   }elseif(-not $task){
    foreach($p in Get-LFProjects $root){
     $dir=Get-LFProjectPath $root $p.id;$request=Join-Path $dir 'thumbs/request.json'
     if(Test-Path $request){
      $status=Join-Path $dir 'thumbs/status.json'
      if(-not (Test-Path $status)){
       Write-LFJson $status @{state='exporting'}
       $child=Start-Process (Join-Path $PSHOME 'powershell.exe') -PassThru -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "'+(Join-Path $PSScriptRoot 'Export-Thumbnails.ps1')+'" -RuntimeRoot "'+$root+'" -Id '+$p.id)
       break
      }
     }
    }
   }
  }
  Write-LFJson $heartbeat @{instance=$Instance;pid=$PID;stage=if($child -or $assemblyChild -or $avatarChildren.Count){'working'}else{'idle'};milestone='E';provider_actions_enabled=$true;utc=[DateTime]::UtcNow.ToString('o')}
  Start-Sleep -Milliseconds 500
 }
}finally{
 # Finish the one authorized in-flight request; never kill or repeat a paid request on clean stop.
 if($child){$child.WaitForExit()}
 if($assemblyChild){$assemblyChild.WaitForExit()}
 foreach($av in $avatarChildren){$av.WaitForExit()}
 Write-LFJson $heartbeat @{instance=$Instance;pid=$PID;stage='stopped';provider_actions_enabled=$false;utc=[DateTime]::UtcNow.ToString('o')}
}
