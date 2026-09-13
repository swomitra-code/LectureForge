param([string]$RuntimeRoot=(Join-Path $env:LOCALAPPDATA 'LectureForgeB'),[int]$Port=8771,[ValidateRange(1,3)][int]$ExpectedTake=2)
$ErrorActionPreference='Stop'
[Net.WebRequest]::DefaultWebProxy=$null
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Import-Module (Join-Path $repo 'studio/Narration.psm1') -Force
Import-Module (Join-Path $repo 'studio/NarrationProvider.psm1') -Force
$root=Get-LFRoot $RuntimeRoot;$url="http://127.0.0.1:$Port/"
$live=Get-Content -Raw (Join-Path $root 'live-narration-results.json')|ConvertFrom-Json;$id=$live.project_id;$dir=Get-LFProjectPath $root $id
$saved=(Get-FileHash (Join-Path $dir 'narration.json')).Hash
$before=Get-LFNarration $root $id
if($before.slides[7].selected_take -ne $ExpectedTake -or $before.provider_calls -ne 3){throw 'Expected reviewed live take and exactly three provider attempts.'}
$token=(Invoke-RestMethod -UseBasicParsing -Uri ($url+'api/bootstrap')).token
$null=Invoke-RestMethod -UseBasicParsing -Method Post -Uri ($url+'api/stop') -Headers @{'X-LF-Token'=$token} -ContentType application/json -Body '{}'
for($i=0;$i -lt 60;$i++){Start-Sleep -Milliseconds 250;$runtime=Get-Content -Raw (Join-Path $root 'runtime.json')|ConvertFrom-Json;if($runtime.status -eq 'stopped'){break}}
if($runtime.status -ne 'stopped'){throw 'Clean stop did not finish.'}
& (Join-Path $repo 'Start-LectureForge.ps1') -RuntimeRoot $root -Port $Port -NoBrowser
$after=Invoke-RestMethod -UseBasicParsing -Uri ($url+"api/narration?id=$id")
for($i=0;$i -lt 100;$i++){
 $health=Invoke-RestMethod -UseBasicParsing -Uri ($url+'api/health') -TimeoutSec 5
 if(-not $health.ready){throw ('Worker heartbeat unavailable at read '+$i+': '+($health|ConvertTo-Json -Compress))}
}
if((Get-FileHash (Join-Path $dir 'narration.json')).Hash -ne $saved -or $after.provider_calls -ne 3 -or $after.slides[7].selected_take -ne $ExpectedTake){throw 'Restart changed narration state.'}
foreach($t in $after.slides[7].takes){if(-not (Test-LFTake $dir $t)){throw 'Take hash changed.'}}
$audio=$url+"api/audio?id=$id&revision=$($after.slides[7].revision)&take=$ExpectedTake"
$request=[Net.HttpWebRequest]::Create($audio);$request.AddRange(1000,1999);$response=$request.GetResponse()
try{if([int]$response.StatusCode -ne 206 -or $response.ContentLength -ne 1000){throw 'Seeking failed after restart.'}}finally{$response.Close()}
$redacted=Protect-NarrationDiagnostic "xi-api-key: secret-fixture`nmessage secret-fixture" @('secret-fixture')
if($redacted.Contains('secret-fixture') -or $redacted.Contains('xi-api-key')){throw 'Diagnostic redaction failed.'}
Write-LFJson (Join-Path $root "resume-narration-take-$ExpectedTake-results.json") @{passed=$true;project_id=$id;state_sha256=$saved;selection=$ExpectedTake;elevenlabs_calls_total=3;elevenlabs_calls_during_restart=0;heygen_calls=0;hashes_unchanged=$true;audio_seek_after_restart=$true;diagnostics_redacted=$true;health_reads_passed=100}
Write-Output "PASS: clean stop; identical live state after restart; Take $ExpectedTake retained; audio seeking works; 0 new calls; diagnostics redacted."
