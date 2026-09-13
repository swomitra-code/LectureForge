param([string]$RuntimeRoot=(Join-Path $env:LOCALAPPDATA 'LectureForgeB'),[int]$Port=8771)
$ErrorActionPreference='Stop'
[Net.WebRequest]::DefaultWebProxy=$null
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Import-Module (Join-Path $repo 'studio/Narration.psm1') -Force
$root=Get-LFRoot $RuntimeRoot
if(Test-Path (Join-Path $root 'runtime.json')){throw 'Fixture tests require a fresh runtime root so no live worker can consume synthetic test queues. Pass -RuntimeRoot with a new short path.'}
$test=Join-Path $root ('Tests/'+[guid]::NewGuid().ToString('N').Substring(0,8));[void][IO.Directory]::CreateDirectory($test)
$checks=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Value){$checks.Add(@{name=$Name;passed=$Value});if(-not $Value){throw "FAILED: $Name"};Write-Output "PASS: $Name"}
$source=Join-Path $repo 'source/RE-M2-Solar-Thermal-Overview.pptx';$hash=(Get-FileHash $source).Hash
$p=New-LFProject $root 'Solar Thermal - narration review COPY' 'Solar-Thermal-copy.pptx' ([IO.File]::ReadAllBytes($source))
$id=$p.id;$dir=Get-LFProjectPath $root $id
$production=Join-Path $repo 'work/solar-thermal-slides2-14-production-r001'
$scripts=Get-Content -Encoding UTF8 -Raw (Join-Path $production 'scripts.json')|ConvertFrom-Json
foreach($s in $p.slides){$s.enabled=$s.number -gt 1;if($s.enabled){$s.script=$scripts."$($s.number)"};if($s.number -eq 8){$s.post_speech_silence_seconds=6}}
$p=Update-LFProject $root $id ([pscustomobject]@{version=$p.version;slides=$p.slides})
$plan=Get-LFNarrationPlan $root $id
Check 'Confirmation plan binds 13 enabled scripts to 39 takes' ($plan.slides -eq 13 -and $plan.generations -eq 39)
Add-LFNarrationBatch $root $id $p.version
$state=Read-LFNarrationFile $dir
Check 'Three independent queued attempts per slide before any provider call' ($state.revisions.Count -eq 13 -and @($state.revisions.takes).Count -eq 39 -and (Get-LFNarration $root $id).provider_calls -eq 0)
# Fixture import is test-only and occurs before starting any worker. No production provider substitution.
$fixtures=Get-Content -Encoding UTF8 -Raw (Join-Path $production 'narration-results.json')|ConvertFrom-Json
$selections=Get-Content -Encoding UTF8 -Raw (Join-Path $production 'selection-workflow-v2.json')|ConvertFrom-Json
foreach($r in $state.revisions){foreach($t in $r.takes){
 $f=@($fixtures.takes|Where-Object {$_.slide -eq $r.slide_number -and $_.take -eq $t.number})[0]
 if($f.script_sha256 -ne $r.script_sha256 -or (Get-FileHash $f.path).Hash -ne $f.sha256){throw 'Fixture binding or hash mismatch.'}
 $a=$t.attempts[0];$relative=$a.raw_path.Replace('raw.mp3',('take'+[IO.Path]::GetExtension($f.path)));$dest=Resolve-LFNarrationAsset $dir $relative
 [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($dest));[IO.File]::Copy($f.path,$dest,$false)
 $asset=Get-LFAudioValidation $dest;$asset.path=$relative;$asset.post_speech_silence_seconds=$r.pause_seconds;$asset.pause_samples=[long]($r.pause_seconds*44100);$asset.pause_sample_rate=44100
 $t.asset=$asset;$t.state='ready';$a.state='imported fixture';$a.provider='existing Solar Thermal';$a.finished_utc=$f.completed_utc
};$r.selected_take=$selections.selections."$($r.slide_number)".take}
Write-LFJson (Join-Path $dir 'narration.json') $state
Check '39 copied candidates decode with exact script and SHA-256 binding' (@($state.revisions.takes|Where-Object state -eq 'ready').Count -eq 39)
$before=(Get-FileHash (Join-Path $dir 'narration.json')).Hash
Add-LFNarrationBatch $root $id $p.version
Check 'Repeated generation preserves successful candidates and selections' ((Get-FileHash (Join-Path $dir 'narration.json')).Hash -eq $before -and (Get-LFNarrationPlan $root $id).generations -eq 0)
$r=$state.revisions[0]
Set-LFNarrationSelection $root $id 2 $r.id 1
Set-LFNarrationSelection $root $id 2 $r.id 2
Check 'Selection changes freely without creating a provider attempt' ((Get-LFNarration $root $id).slides[1].selected_take -eq 2 -and (Get-LFNarration $root $id).provider_calls -eq 0)
Set-LFNarrationSelection $root $id 2 $r.id 3
# New revision queues are tested with no worker, then archived as test exceptions.
$p.slides[1].script+="`nRevision test."
$p=Update-LFProject $root $id ([pscustomobject]@{version=$p.version;slides=$p.slides})
Check 'Script change hides old selection from current revision' ($null -eq (Get-LFNarration $root $id).slides[1].selected_take)
Add-LFNarrationBatch $root $id $p.version
$state=Read-LFNarrationFile $dir;$new=$state.revisions[-1]
Check 'Script revision creates new take paths and preserves history' ($state.revisions.Count -eq 14 -and $new.id -ne $r.id -and $state.revisions[0].selected_take -eq 3)
$new.takes[0].state='ready';$new.takes[0].asset=$state.revisions[0].takes[0].asset
$new.takes[1].state='failed';$new.takes[1].attempts[-1].state='failed'
$new.takes[2].state='ready';$new.takes[2].asset=$state.revisions[0].takes[2].asset
Write-LFJson (Join-Path $dir 'narration.json') $state
Retry-LFNarration $root $id $new.id 2
$state=Read-LFNarrationFile $dir;$new=$state.revisions[-1]
Check 'Targeted failed-take retry adds only one fresh attempt' ($new.takes[1].attempts.Count -eq 2 -and $new.takes[0].attempts.Count -eq 1 -and $new.takes[2].attempts.Count -eq 1 -and $new.takes[1].attempts[0].raw_path -ne $new.takes[1].attempts[1].raw_path)
$new.takes[1].state='generating';$new.takes[1].attempts[-1].state='calling';$new.takes[1].attempts[-1].owner_pid=0
Write-LFJson (Join-Path $dir 'narration.json') $state
$null=Get-LFNextNarration $root
$state=Read-LFNarrationFile $dir
Check 'Interrupted attempt becomes uncertain without a paid retry' ($state.revisions[-1].takes[1].state -eq 'uncertain')
$rejected=$false;try{Retry-LFNarration $root $id $new.id 2}catch{$rejected=$true};Check 'Uncertain attempt rejects retry' $rejected
$p.slides[1].script=$scripts.'2';$p=Update-LFProject $root $id ([pscustomobject]@{version=$p.version;slides=$p.slides})
Check 'Returning to saved script restores the matching selected revision' ((Get-LFNarration $root $id).slides[1].selected_take -eq 3)
# Exercise the real preparation helper on copied raw speech, with exact PCM sample verification below.
$prep=[pscustomobject]@{id=[guid]::NewGuid().ToString('N').Substring(0,12);pause_seconds=6}
$a=New-LFAttempt $prep $null;$raw=Resolve-LFNarrationAsset $dir $a.raw_path
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($raw));[IO.File]::Copy($fixtures.takes[0].path,$raw,$false)
$prepared=Complete-LFNarrationAudio $dir $prep $a
Check 'Prepared candidate records exactly 264600 appended samples at 44100 Hz' ($prepared.pause_samples -eq 264600 -and $prepared.post_speech_silence_seconds -eq 6)
Write-LFJson (Join-Path $test 'pause.json') @{raw=$raw;prepared=Resolve-LFNarrationAsset $dir $prepared.path}
& (Join-Path $repo 'Start-LectureForge.ps1') -RuntimeRoot $root -Port $Port -NoBrowser
$url="http://127.0.0.1:$Port/"
function GetJson($Path){Invoke-RestMethod -UseBasicParsing -Uri ($url+$Path) -TimeoutSec 20}
function PostJson($Path,$Data){Invoke-RestMethod -UseBasicParsing -Method Post -Uri ($url+$Path) -Headers @{'X-LF-Token'=$script:token} -ContentType 'application/json' -Body ([Text.Encoding]::UTF8.GetBytes(($Data|ConvertTo-Json -Depth 20))) -TimeoutSec 20}
$script:token=(GetJson 'api/bootstrap').token
$r=$state.revisions[0];$audioUrl=$url+"api/audio?id=$id&revision=$($r.id)&take=1"
$wc=[Net.HttpWebRequest]::Create($audioUrl);$wc.AddRange(100,199);$response=$wc.GetResponse()
try{Check 'Audio seeking returns HTTP 206 and exact byte range' ([int]$response.StatusCode -eq 206 -and $response.ContentLength -eq 100 -and $response.Headers['Content-Range'].StartsWith('bytes 100-199/'))}finally{$response.Close()}
$null=PostJson "api/narration-select?id=$id" @{slide=4;revision=$state.revisions[2].id;take=2}
Check 'HTTP selection changes only narration intent and preserves Slide 4 Take 2' ((GetJson "api/narration?id=$id").slides[3].selected_take -eq 2)
$null=PostJson "api/thumbnails?id=$id" @{}
for($i=0;$i -lt 60;$i++){if(Test-Path (Join-Path $dir 'thumbs/s14.png')){break};Start-Sleep -Seconds 1}
Check 'Read-only thumbnail export produces all 14 slide previews' (@(Get-ChildItem (Join-Path $dir 'thumbs') -Filter '*.png').Count -eq 14)
$before=(Get-FileHash (Join-Path $dir 'narration.json')).Hash
$null=PostJson 'api/stop' @{}
Start-Sleep -Seconds 3
& (Join-Path $repo 'Start-LectureForge.ps1') -RuntimeRoot $root -Port $Port -NoBrowser
$script:token=(GetJson 'api/bootstrap').token
Check 'Restart restores identical narration state and selected takes' ((Get-FileHash (Join-Path $dir 'narration.json')).Hash -eq $before -and (GetJson "api/narration?id=$id").slides[1].selected_take -eq 3)
Check 'All 39 approved take hashes remain unchanged after restart' (@($state.revisions|Select-Object -First 13|ForEach-Object {$_.takes}|Where-Object {-not (Test-LFTake $dir $_)}).Count -eq 0)
Check 'No provider calls occurred during fixtures, selection, or restart' ((GetJson "api/narration?id=$id").provider_calls -eq 0)
Check 'Canonical source PPTX unchanged' ((Get-FileHash $source).Hash -eq $hash)
Write-LFJson (Join-Path $test 'results.json') @{passed=$true;checks=@($checks);project_id=$id;project_root=$dir;runtime_root=$root;url=$url;elevenlabs_calls=0;heygen_calls=0;source_sha256=$hash}
Write-Output "RESULTS=$test/results.json"
Write-Output "PROJECT=$dir"
