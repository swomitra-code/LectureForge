param([Parameter(Mandatory=$true)][string]$RuntimeRoot,[Parameter(Mandatory=$true)][string]$Instance)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Production.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Narration.psm1') -Force
$root=Get-LFRoot $RuntimeRoot
$stop=Join-Path $root "$Instance.stop"
$heartbeat=Join-Path $root "$Instance.worker.json"
$child=$null;$task=$null
try{
 while(-not (Test-Path -LiteralPath $stop)){
  $runtime=Join-Path $root 'runtime.json'
  if(Test-Path $runtime){
   $owner=Get-Content -Encoding UTF8 -Raw $runtime|ConvertFrom-Json
   if($owner.instance -eq $Instance -and -not (Get-Process -Id $owner.server_pid -ErrorAction SilentlyContinue)){break}
  }
  if($child -and $child.HasExited){
   if($task -and $task.ContainsKey('attempt')){
    Invoke-LFNarrationLock $root $task.id {param($dir)
     $state=Read-LFNarrationFile $dir;$r=@($state.revisions|Where-Object id -eq $task.revision)[0];$t=$r.takes[$task.take-1]
     if($t.state -in @('dispatched','generating')){$t.state='uncertain';$t.attempts[-1].state='uncertain';$t.attempts[-1].error='Executor exited without a final result. No automatic retry.';Write-LFJson (Join-Path $dir 'narration.json') $state}
    }
   }
   $child=$null;$task=$null
  }
  if(-not $child){
   $task=Get-LFNextNarration $root
   if($task -and -not $task.busy){
    $args='-NoProfile -ExecutionPolicy Bypass -File "'+(Join-Path $PSScriptRoot 'Invoke-NarrationTake.ps1')+'" -RuntimeRoot "'+$root+'" -Id '+$task.id+' -Revision '+$task.revision+' -Take '+$task.take+' -Attempt '+$task.attempt
    $child=Start-Process (Join-Path $PSHOME 'powershell.exe') -WindowStyle Hidden -PassThru -ArgumentList $args
   }elseif(-not $task){
    foreach($p in Get-LFProjects $root){
     $dir=Get-LFProjectPath $root $p.id;$request=Join-Path $dir 'thumbs/request.json'
     if(Test-Path $request){
      $status=Join-Path $dir 'thumbs/status.json'
      if(-not (Test-Path $status)){
       Write-LFJson $status @{state='exporting'}
       $child=Start-Process (Join-Path $PSHOME 'powershell.exe') -WindowStyle Hidden -PassThru -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "'+(Join-Path $PSScriptRoot 'Export-Thumbnails.ps1')+'" -RuntimeRoot "'+$root+'" -Id '+$p.id)
       break
      }
     }
    }
   }
  }
  Write-LFJson $heartbeat @{instance=$Instance;pid=$PID;stage=if($child){'working'}else{'idle'};milestone='B';provider_actions_enabled=$true;utc=[DateTime]::UtcNow.ToString('o')}
  Start-Sleep -Milliseconds 500
 }
}finally{
 # Finish the one authorized in-flight request; never kill or repeat a paid request on clean stop.
 if($child){$child.WaitForExit()}
 Write-LFJson $heartbeat @{instance=$Instance;pid=$PID;stage='stopped';provider_actions_enabled=$false;utc=[DateTime]::UtcNow.ToString('o')}
}
