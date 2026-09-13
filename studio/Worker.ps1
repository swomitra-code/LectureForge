param([Parameter(Mandatory=$true)][string]$RuntimeRoot,[Parameter(Mandatory=$true)][string]$Instance)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Production.psm1') -Force
$root=Get-LFRoot $RuntimeRoot
$stop=Join-Path $root "$Instance.stop"
$heartbeat=Join-Path $root "$Instance.worker.json"
try{
 while(-not (Test-Path -LiteralPath $stop)){
  $runtime=Join-Path $root 'runtime.json'
  if(Test-Path $runtime){
   $owner=Get-Content -Encoding UTF8 -Raw $runtime|ConvertFrom-Json
   if($owner.instance -eq $Instance -and -not (Get-Process -Id $owner.server_pid -ErrorAction SilentlyContinue)){break}
  }
  Write-LFJson $heartbeat @{instance=$Instance;pid=$PID;stage='idle';milestone='A';provider_actions_enabled=$false;utc=[DateTime]::UtcNow.ToString('o')}
  Start-Sleep -Milliseconds 500
 }
}finally{Write-LFJson $heartbeat @{instance=$Instance;pid=$PID;stage='stopped';provider_actions_enabled=$false;utc=[DateTime]::UtcNow.ToString('o')}}
