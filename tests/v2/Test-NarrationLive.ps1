param([string]$RuntimeRoot=(Join-Path $env:LOCALAPPDATA 'LectureForgeB'),[int]$Port=8771)
# Opt-in live smoke test: at most three new ElevenLabs requests. Never retry uncertain attempts.
$ErrorActionPreference='Stop'
[Net.WebRequest]::DefaultWebProxy=$null
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Import-Module (Join-Path $repo 'studio/Narration.psm1') -Force
$root=Get-LFRoot $RuntimeRoot;$checkpoint=Join-Path $root 'live-narration-test.json';$url="http://127.0.0.1:$Port/"
function GetJson($Path){Invoke-RestMethod -UseBasicParsing -Uri ($url+$Path) -TimeoutSec 20}
function PostJson($Path,$Data){Invoke-RestMethod -UseBasicParsing -Method Post -Uri ($url+$Path) -Headers @{'X-LF-Token'=$token} -ContentType 'application/json' -Body ([Text.Encoding]::UTF8.GetBytes(($Data|ConvertTo-Json -Depth 20))) -TimeoutSec 20}
$token=(GetJson 'api/bootstrap').token
if(Test-Path $checkpoint){$record=Get-Content -Raw $checkpoint|ConvertFrom-Json;$id=$record.project_id}else{
 $source=Join-Path $repo 'source/RE-M2-Solar-Thermal-Overview.pptx'
 $p=New-LFProject $root 'Solar Thermal - LIVE three-take test COPY' 'Solar-Thermal-copy.pptx' ([IO.File]::ReadAllBytes($source))
 $scripts=Get-Content -Raw -Encoding UTF8 (Join-Path $repo 'work/solar-thermal-slides2-14-production-r001/scripts.json')|ConvertFrom-Json
 foreach($s in $p.slides){$s.enabled=$s.number -eq 8;if($s.enabled){$s.script=$scripts.'8';$s.post_speech_silence_seconds=6}}
 $p=Update-LFProject $root $p.id ([pscustomobject]@{version=$p.version;slides=$p.slides});$id=$p.id
 Write-LFJson $checkpoint @{project_id=$id;authorized_calls=3;created_utc=[DateTime]::UtcNow.ToString('o')}
}
$plan=GetJson "api/narration-plan?id=$id"
if($plan.generations -eq 3 -and $plan.slides -eq 1){$null=PostJson "api/narration-generate?id=$id" @{confirm=$true;version=$plan.version}}elseif($plan.generations -ne 0){throw 'Unexpected generation count. No requests authorized.'}
for($i=0;$i -lt 180;$i++){
 $state=GetJson "api/narration?id=$id";$s=$state.slides[7]
 Write-Output ((@($s.takes|ForEach-Object {"Take $($_.number): $($_.state)"})) -join ' | ')
 if(@($s.takes|Where-Object state -in @('failed','uncertain','local exception')).Count){throw 'Live test exception. Preserve attempts; do not repeat paid requests.'}
 if(@($s.takes|Where-Object state -eq 'ready').Count -eq 3){break}
 Start-Sleep -Seconds 2
}
if(@($s.takes|Where-Object state -eq 'ready').Count -ne 3){throw 'Timed out waiting. Poll existing attempts only.'}
if($state.provider_calls -ne 3){throw 'Expected exactly three submitted requests.'}
$null=PostJson "api/narration-select?id=$id" @{slide=8;revision=$s.revision;take=1}
$null=PostJson "api/narration-select?id=$id" @{slide=8;revision=$s.revision;take=2}
$state=GetJson "api/narration?id=$id"
Write-LFJson (Join-Path $root 'live-narration-results.json') @{passed=$true;project_id=$id;elevenlabs_calls=$state.provider_calls;heygen_calls=0;selected_take=$state.slides[7].selected_take;takes=$state.slides[7].takes}
Write-Output "PASS: three independent ElevenLabs takes decode; selected Take 2; HeyGen calls 0; PROJECT=$id"
