param([string]$RuntimeRoot=(Join-Path $env:LOCALAPPDATA 'LectureForgeC'),[int]$Port=8772)
$ErrorActionPreference='Stop';[Net.WebRequest]::DefaultWebProxy=$null
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'));Import-Module (Join-Path $repo 'studio/Authorization.psm1') -Force
$result=Get-Content -Raw (Join-Path $RuntimeRoot 'results.json')|ConvertFrom-Json;$id=$result.project_id;$url="http://127.0.0.1:$Port/";$token=(Invoke-RestMethod ($url+'api/bootstrap')).token
function Post($Path,$Body){Invoke-RestMethod -Method Post -Uri ($url+$Path) -Headers @{'X-LF-Token'=$token} -ContentType application/json -Body ($Body|ConvertTo-Json -Depth 20)}
function View {Invoke-RestMethod ($url+"api/selections?id=$id")}
$checks=@()
function Check($Name,[bool]$Ok){if(-not $Ok){throw "FAIL: $Name"};$script:checks+=@{name=$Name;passed=$true};"PASS: $Name"}
$tabA=View;$old=$tabA.current.binding_sha256;$n=Get-LFNarration $RuntimeRoot $id;$rev=$n.slides[4].revision
$null=Post "api/narration-select?id=$id" @{slide=5;revision=$rev;take=1}
$rejected=$false;try{$null=Post "api/avatar-authorize?id=$id" @{confirm=$true;review_id=$old;expected=$old}}catch{$rejected=$_.Exception.Response.StatusCode.value__ -eq 400}
Check 'Two HTTP clients reject tab A after tab B changes narration' $rejected
$null=Post "api/narration-select?id=$id" @{slide=5;revision=$rev;take=3}
$c=View;$null=Post "api/selection-review?id=$id" @{expected=$c.current.binding_sha256;phase='review'}
$rejected=$false;try{$null=Post "api/avatar-authorize?id=$id" @{confirm=$true;review_id=$c.current.binding_sha256;expected=$c.current.binding_sha256}}catch{$rejected=$_.Exception.Response.StatusCode.value__ -eq 400}
Check 'Review screen alone cannot authorize without final confirmation phase' $rejected
$null=Post "api/selection-change?id=$id" @{slide=4};$after=View
Check 'CHANGE route persists Slide 4 narration navigation without changing intent' ($after.saved.phase -eq 'narration' -and $after.saved.slide -eq 4 -and $after.current.snapshot.slides[2].selected_take -eq 2)
Check 'No provider calls or submissions through either HTTP client' ((Get-LFNarration $RuntimeRoot $id).provider_calls -eq 0 -and -not $after.provider_submission_enabled)
Write-LFJson (Join-Path $RuntimeRoot 'http-results.json') @{checks=$checks;elevenlabs_calls=0;heygen_calls=0}
