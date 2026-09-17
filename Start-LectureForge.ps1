param([string]$RuntimeRoot=(Join-Path $env:LOCALAPPDATA 'LectureForge'),[ValidateRange(1024,65535)][int]$Port=8770,[switch]$NoBrowser)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'studio/Production.psm1') -Force
$root=Get-LFRoot $RuntimeRoot
# Loopback control traffic must not use a machine-configured outbound proxy.
[Net.WebRequest]::DefaultWebProxy=$null
function Check($Condition,[string]$Label){if(-not $Condition){throw "$Label unavailable. Correct the prerequisite and run Start-LectureForge.ps1 again."};Write-Host ([char]0x2713+' '+$Label) -ForegroundColor Green}
try{
 Check (-not [string]::IsNullOrWhiteSpace($env:ELEVENLABS_API_KEY)) 'ElevenLabs API key found'
 Check (-not [string]::IsNullOrWhiteSpace($env:HEYGEN_API_KEY)) 'HeyGen API key found'
 Check ($null -ne [Type]::GetTypeFromProgID('PowerPoint.Application')) 'PowerPoint found'
 Check ($null -ne (Get-Command ffmpeg -ErrorAction SilentlyContinue) -and $null -ne (Get-Command ffprobe -ErrorAction SilentlyContinue)) 'FFmpeg found'
 [void][IO.Directory]::CreateDirectory((Join-Path $root 'Projects'))
 $hash=[Security.Cryptography.SHA256]::Create();try{$name=([BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($root)))).Replace('-','').Substring(0,16)}finally{$hash.Dispose()}
 $mutex=[Threading.Mutex]::new($false,"Local\LectureForge-$name")
 if(-not $mutex.WaitOne(10000)){throw 'Another LectureForge launcher is starting.'}
 try{
  $running=$false;$runtime=Join-Path $root 'runtime.json'
  if(Test-Path $runtime){
   $old=Read-LFSharedJson $runtime
   if($old.status -eq 'running'){
    try{$health=Invoke-RestMethod -UseBasicParsing -Uri ($old.url+'api/health') -TimeoutSec 2;$running=($health.instance -eq $old.instance -and $health.ready -and $health.milestone -eq 'E' -and $health.api_version -eq 2)}catch{}
    if($running){$url=$old.url}else{
     # Signal only the old instance; never terminate unrelated/user-owned processes.
     [IO.File]::WriteAllText((Join-Path $root ($old.instance+'.stop')),'stop')
     Start-Sleep -Milliseconds 750
    }
   }
  }
  if(-not $running){
   $probe=[Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback,$Port);try{$probe.Start()}finally{$probe.Stop()}
   $instance=[guid]::NewGuid().ToString('N').Substring(0,12)
   $exe=Join-Path $PSHOME 'powershell.exe'
   $common='-NoProfile -ExecutionPolicy Bypass -File '
   $worker=Start-Process $exe -PassThru -RedirectStandardError (Join-Path $root "$instance.worker-error.log") -RedirectStandardOutput (Join-Path $root "$instance.worker.log") -ArgumentList ($common+'"'+(Join-Path $PSScriptRoot 'studio/Worker.ps1')+'" -RuntimeRoot "'+$root+'" -Instance '+$instance)
   $server=Start-Process $exe -PassThru -RedirectStandardError (Join-Path $root "$instance.server-error.log") -RedirectStandardOutput (Join-Path $root "$instance.server.log") -ArgumentList ($common+'"'+(Join-Path $PSScriptRoot 'studio/Server.ps1')+'" -RuntimeRoot "'+$root+'" -Instance '+$instance+' -Port '+$Port+' -WorkerPid '+$worker.Id)
   $url="http://127.0.0.1:$Port/";$ready=$false
   for($i=0;$i -lt 40;$i++){
    Start-Sleep -Milliseconds 250
    try{$health=Invoke-RestMethod -UseBasicParsing -Uri ($url+'api/health') -TimeoutSec 1;if($health.instance -eq $instance -and $health.ready -and $health.api_version -eq 2){$ready=$true;break}}catch{}
    if($server.HasExited){break}
   }
   if(-not $ready){[IO.File]::WriteAllText((Join-Path $root "$instance.stop"),'stop');throw 'LectureForge could not start. Check the runtime folder and port availability.'}
  }
  if(-not $NoBrowser){Start-Process $url}
  Write-Host ([char]0x2713+' LectureForge ready') -ForegroundColor Green
 }finally{$mutex.ReleaseMutex();$mutex.Dispose()}
}catch{Write-Host $_.Exception.Message -ForegroundColor Red;exit 1}
