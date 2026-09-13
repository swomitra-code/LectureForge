param([string]$RuntimeRoot=(Join-Path $env:LOCALAPPDATA 'LectureForgeD2'),[string]$Id='709e1df2cb7a',[int]$Port=8778)
$ErrorActionPreference='Stop';[Net.WebRequest]::DefaultWebProxy=$null
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'));Import-Module (Join-Path $repo 'studio/AvatarProduction.psm1') -Force
$root=Get-LFRoot $RuntimeRoot;$dir=Get-LFProjectPath $root $Id;$checks=[Collections.Generic.List[object]]::new()
function Check($Name,[bool]$OK){$checks.Add(@{name=$Name;passed=$OK});Write-LFJson (Join-Path $root 'recovery-results.json') @{checks=@($checks);project_id=$Id;url="http://127.0.0.1:$Port/";elevenlabs_calls=0;heygen_calls=0};if(-not $OK){throw "FAIL: $Name"};"PASS: $Name"}
function State {Get-LFAvatarStatus $root $Id}
$s=State;$j8=$s.jobs|Where-Object slide -eq 8;$hash=(Get-FileHash (Resolve-LFAvatarPath $dir $j8.output_path)).Hash
if($j8.stage -eq 'Exception'){Repair-LFAvatarJob $root $Id $j8.id validate}
Check 'Targeted retry reuses the existing selected Slide 8 MP4' ((Get-FileHash (Resolve-LFAvatarPath $dir $j8.output_path)).Hash -eq $hash -and (State).provider_calls -eq 0)
& (Join-Path $repo 'Start-LectureForge.ps1') -RuntimeRoot $root -Port $Port -NoBrowser
$url="http://127.0.0.1:$Port/";function Get($Path){Invoke-RestMethod ($url+$Path) -TimeoutSec 30};$token=(Get 'api/bootstrap').token
function Post($Path,$Body){Invoke-RestMethod -Method Post -Uri ($url+$Path) -Headers @{'X-LF-Token'=$script:token} -ContentType application/json -Body ($Body|ConvertTo-Json -Depth 15) -TimeoutSec 30}
for($i=0;$i -lt 90;$i++){Start-Sleep -Seconds 2;$s=Get "api/avatar-status?id=$Id";if($s.ready -eq 13){break};if($s.exceptions){$s.jobs|Where-Object stage -eq Exception|Select-Object slide,error|ConvertTo-Json|Write-Output;throw 'Retry failed'}}
Check 'All 13 avatars reach Avatar Ready after isolated local retry' ($s.ready -eq 13 -and $s.exceptions -eq 0)
Check 'Every ready avatar has verified authoritative narration audio' (@($s.jobs|Where-Object {-not $_.validation.authoritative_audio_verified}).Count -eq 0)
$j8=$s.jobs|Where-Object slide -eq 8
Check 'Slide 8 prepared hash and exact 6.000-second silence pass local decode' ($j8.validation.prepared_pause_seconds -eq 6 -and $j8.validation.pause_verified -and $j8.validation.narration_sha256 -eq 'E234676D7F92B42D20E98A3CB6CADF5BAA89595C2DCA906005EB7913607D69A5')
Check 'Real shared converter preserves motion and authoritative audio' ('White Avatar Building' -in $s.jobs[0].history -and $s.jobs[0].validation.motion.motion_present)
function Hashes { @((Get-ChildItem (Join-Path $dir 'a') -Recurse -Filter job.json)|Sort-Object FullName|ForEach-Object {(Get-FileHash $_.FullName).Hash}) -join ',' }
$before=Hashes;$approval=(Get-FileHash (Get-LFApprovalPath $dir $s.authorization authorization)).Hash
Add-Type -AssemblyName System.Net.Http;$client=[Net.Http.HttpClient]::new();$client.Timeout=[TimeSpan]::FromSeconds(30)
try{$requests=@(1..8|ForEach-Object {$client.GetStringAsync($url+"api/avatar-status?id=$Id")});$views=@(foreach($r in $requests){$r.GetAwaiter().GetResult()|ConvertFrom-Json});Check 'Eight concurrent dashboard polls return accurate durable state' (@($views|Where-Object {$_.ready -ne 13 -or $_.provider_calls -ne 0}).Count -eq 0)}finally{$client.Dispose()}
Check 'Status reads preserve state and make zero provider calls' ((Hashes) -eq $before -and (State).provider_calls -eq 0)
$page=Invoke-WebRequest -UseBasicParsing $url
Check 'Refresh resources expose dashboard and durable ready state' ($page.Content.Contains('avatar-dashboard') -and (Get "api/avatar-status?id=$Id").ready -eq 13)
$null=Post "api/avatar-pause?id=$Id" @{paused=$true};Check 'Pause Queue persists through HTTP' (State).paused
$null=Post 'api/stop' @{}
for($i=0;$i -lt 40;$i++){Start-Sleep -Milliseconds 250;$rt=Read-LFSharedJson (Join-Path $root 'runtime.json');if($rt.status -eq 'stopped'){break}}
& (Join-Path $repo 'Start-LectureForge.ps1') -RuntimeRoot $root -Port $Port -NoBrowser;$script:token=(Get 'api/bootstrap').token
Check 'Server restart preserves paused queue, jobs and authorization bytes' ((Hashes) -eq $before -and (State).paused -and (Get-FileHash (Get-LFApprovalPath $dir $s.authorization authorization)).Hash -eq $approval)
$rt=Read-LFSharedJson (Join-Path $root 'runtime.json');$process=Get-CimInstance Win32_Process -Filter "ProcessId=$($rt.worker_pid)"
if($process.CommandLine -notlike '*studio\Worker.ps1*' -or $process.CommandLine -notlike "*$root*"){throw 'Refusing to interrupt an unrelated process.'}
Stop-Process -Id $rt.worker_pid -Force
for($i=0;$i -lt 60;$i++){Start-Sleep -Milliseconds 500;$after=Read-LFSharedJson (Join-Path $root 'runtime.json');if($after.worker_pid -ne $rt.worker_pid -and (Get 'api/health').ready){break}}
Check 'Server recovers stopped worker without duplicate work' ($after.worker_pid -ne $rt.worker_pid -and (Hashes) -eq $before -and (State).provider_calls -eq 0)
$null=Post "api/avatar-pause?id=$Id" @{paused=$false};Check 'Resume Queue keeps completed assets intact' ((Get "api/avatar-status?id=$Id").ready -eq 13 -and (Hashes) -eq $before)
Check 'Zero narration and avatar provider calls during recovery' ((Get-LFNarration $root $Id).provider_calls -eq 0 -and (State).provider_calls -eq 0)
"RESULTS=$root/recovery-results.json"
