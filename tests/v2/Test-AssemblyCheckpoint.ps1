param([string]$RuntimeRoot=(Join-Path $env:LOCALAPPDATA 'LectureForgeE'),[int]$Port=8788)
$ErrorActionPreference='Stop'
[Net.WebRequest]::DefaultWebProxy=$null
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Import-Module (Join-Path $repo 'studio/Assembly.psm1') -Force
$fixture=Read-LFSharedJson (Join-Path $RuntimeRoot 'results.json')
$id=$fixture.project_id;$dir=$fixture.project_root
$stateBefore=(Get-FileHash (Join-Path $dir 'assembly.json')).Hash
$url="http://127.0.0.1:$Port/"
$bootstrap=Invoke-RestMethod ($url+'api/bootstrap')
$null=Invoke-RestMethod -Method Post -Uri ($url+'api/stop') -Headers @{'X-LF-Token'=$bootstrap.token} -Body '{}' -ContentType application/json
for($i=0;$i -lt 60;$i++){
 Start-Sleep -Milliseconds 250
 if((Read-LFSharedJson (Join-Path $RuntimeRoot 'runtime.json')).status -eq 'stopped'){break}
}
& (Join-Path $repo 'Start-LectureForge.ps1') -RuntimeRoot $RuntimeRoot -Port $Port -NoBrowser
$checks=@()
function Check($Name,[bool]$OK){
 if(-not $OK){throw "FAIL: $Name"}
 $script:checks+=@{name=$Name;passed=$true}
 "PASS: $Name"
}
$health=Invoke-RestMethod ($url+'api/health')
Check 'Final server/worker startup runs Milestone E' ($health.ready -and $health.milestone -eq 'E')
Check 'Final restart preserves byte-identical completed assembly state' ((Get-FileHash (Join-Path $dir 'assembly.json')).Hash -eq $stateBefore)
$status=Get-LFAssemblyStatus $RuntimeRoot $id
$ready=Get-LFAssemblyReadiness $RuntimeRoot $id
Check 'Final Recording and all 14 bindings remain valid' ($ready.ready -and $ready.snapshot.slides.Count -eq 14 -and $status.current.binding -eq $ready.binding -and (Get-FileHash $status.current.output).Hash -eq $status.current.output_sha256)
$sourceHash=(Get-FileHash (Join-Path $repo 'source/RE-M2-Solar-Thermal-Overview.pptx')).Hash
Check 'Final source integrity and zero provider calls' ($sourceHash -eq '6340C137FD85021957721341298D640C8FAF59D36507D904DF09948DCB576C01' -and (Get-LFAvatarStatus $RuntimeRoot $id).provider_calls -eq 0 -and (Get-LFNarration $RuntimeRoot $id).provider_calls -eq 0)
Write-LFJson (Join-Path $RuntimeRoot 'checkpoint-results.json') @{
 checks=$checks;output=$status.current.output;output_sha256=$status.current.output_sha256
 source_sha256=$sourceHash;elevenlabs_calls=0;heygen_calls=0;url=$url
}
Write-Output ('Recording: '+$status.current.output)
