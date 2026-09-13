param([string]$RuntimeRoot=(Join-Path $env:LOCALAPPDATA 'LectureForgeC'),[int]$Port=8772)
$ErrorActionPreference='Stop';[Net.WebRequest]::DefaultWebProxy=$null
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Import-Module (Join-Path $repo 'studio/Authorization.psm1') -Force
$root=Get-LFRoot $RuntimeRoot
if(Test-Path (Join-Path $root 'runtime.json')){throw 'Use a fresh isolated test root.'}
$id=[guid]::NewGuid().ToString('N').Substring(0,12);$dir=Get-LFProjectPath $root $id
$template=Join-Path $env:LOCALAPPDATA 'LectureForgeB/Projects/4b9d067085b3'
[void][IO.Directory]::CreateDirectory((Split-Path $dir))
Copy-Item -LiteralPath $template -Destination $dir -Recurse
$p=Get-Content -Raw -Encoding UTF8 (Join-Path $dir 'project.json')|ConvertFrom-Json
$p.id=$id;$p.name='Solar Thermal - Milestone C authorization COPY'
Write-LFJson (Join-Path $dir 'project.json') $p
# Test C independently of the D consumer: pause this fixture queue before startup.
Write-LFJson (Join-Path $dir 'avatar-queue.json') @{schema='lectureforge-avatar-queue-1';paused=$true;batches=@();jobs=@()}
Write-LFJson (Join-Path $root 'avatar-policy.json') @{provider_calls_enabled=$false}
$checks=[Collections.Generic.List[object]]::new()
function Check($Name,[bool]$Passed){$checks.Add(@{name=$Name;passed=$Passed});if(-not $Passed){throw "FAILED: $Name"};Write-Output "PASS: $Name"}
function Current {Get-LFSelectionReview $root $id}
function Change([int]$Slide,[int]$Take){$n=Get-LFNarration $root $id;Set-LFNarrationSelection $root $id $Slide $n.slides[$Slide-1].revision $Take}
function Prepare { $c=Current;$null=Save-LFSelectionReview $root $id $c.current.binding_sha256 'confirmation';Current }
function Reject([scriptblock]$Action){try{& $Action|Out-Null;return $false}catch{return $true}}
$baseline=Current
Check 'All 13 enabled selections present with 39 reused narration fixtures' ($baseline.current.snapshot.can_authorize -and $baseline.current.snapshot.slides.Count -eq 13 -and @((Read-LFNarrationFile $dir).revisions|Select-Object -First 13|ForEach-Object {$_.takes}).Count -eq 39)
Check 'No catalog means 13 expected new jobs, zero actual calls' ($baseline.current.snapshot.expected_new_provider_jobs -eq 13 -and (Get-LFNarration $root $id).provider_calls -eq 0)
$assets=Join-Path $dir 'avatars';[void][IO.Directory]::CreateDirectory($assets)
$production=Join-Path $repo 'work/solar-thermal-slides2-14-production-r001'
$registry=@{assets=@()}
foreach($n in 2,4){
 $job=Get-Content -Raw (Join-Path $production "jobs/slide-$n/job.json")|ConvertFrom-Json
 if($job.status -ne 'complete' -or -not $job.automated_checks_passed){throw 'Historical job validation missing.'}
 $rel="avatars/s$n.mp4";$source=Join-Path $production "jobs/slide-$n/avatar-white.mp4"
 Copy-Item -LiteralPath $source -Destination (Join-Path $dir $rel)
 $registry.assets+=@{id="historical-s$n";slide_number=$n;status=if($n -eq 4){'rejected'}else{'approved'};rejected=($n -eq 4);narration_sha256=$job.narration_sha256;preset_sha256=$baseline.current.snapshot.avatar_preset_sha256;path=$rel;sha256=(Get-FileHash $source).Hash;validation=@{passed=$true;authoritative_audio_verified=$true}}
}
Write-LFJson (Join-Path $dir 'avatars.json') $registry
$c=Current
Check 'Exact approved avatar reuse reduces expected jobs to 12' ($c.current.snapshot.reusable_avatars -eq 1 -and $c.current.snapshot.expected_new_provider_jobs -eq 12 -and $c.current.snapshot.slides[0].new_jobs -eq 0)
Check 'Rejected Slide 4 Take 1 does not match approved Take 2' ($c.current.snapshot.slides[2].selected_take -eq 2 -and -not $c.current.snapshot.slides[2].reusable_avatar)
Change 4 1;$c=Current
Check 'Even exact-hash rejected history receives no reuse credit' (-not $c.current.snapshot.slides[2].reusable_avatar)
Change 4 2
Change 3 0;$c=Current
Check 'Missing selection is visibly blocked' (-not $c.current.snapshot.can_authorize -and $c.current.snapshot.slides[1].status -eq 'missing narration selection')
Check 'Missing selection cannot open final confirmation' (Reject {Save-LFSelectionReview $root $id $c.current.binding_sha256 'confirmation'})
Change 3 1
$p=Read-LFProject $root $id;$original=$p.slides[2].script;$p.slides[2].script+="`nChanged for stale test."
$p=Update-LFProject $root $id ([pscustomobject]@{version=$p.version;slides=$p.slides});$c=Current
Check 'Script changed after selection is clearly flagged' ($c.current.snapshot.slides[1].status -eq 'script changed after selection' -and -not $c.current.snapshot.can_authorize)
$p.slides[2].script=$original;$p=Update-LFProject $root $id ([pscustomobject]@{version=$p.version;slides=$p.slides})
$p.slides[7].post_speech_silence_seconds=5;$p=Update-LFProject $root $id ([pscustomobject]@{version=$p.version;slides=$p.slides});$c=Current
Check 'Changed preparation makes narration selection stale' ($c.current.snapshot.slides[6].status -eq 'stale narration selection')
$p.slides[7].post_speech_silence_seconds=6;$p=Update-LFProject $root $id ([pscustomobject]@{version=$p.version;slides=$p.slides})
$state=Read-LFNarrationFile $dir;$r=Get-LFCurrentRevision $state $p $p.slides[1];$t=$r.takes[$r.selected_take-1];$originalPath=$t.asset.path
$t.asset.path=$t.asset.path.Replace('.mp3','.wav');Write-LFJson (Join-Path $dir 'narration.json') $state;$c=Current
Check 'Missing selected asset blocks authorization without regenerating' ($c.current.snapshot.slides[0].status -eq 'selected asset missing/invalid' -and -not $c.current.snapshot.can_authorize)
$t.asset.path=$originalPath;Write-LFJson (Join-Path $dir 'narration.json') $state
$c=Prepare;$review=$c.review
$slide8=$review.snapshot.slides[6]
Check 'Slide 8 snapshot binds prepared hash and exact six-second pause' ($slide8.preparation.post_speech_silence_seconds -eq 6 -and $slide8.preparation.pause_samples -eq 264600 -and $slide8.narration.sha256 -eq 'E234676D7F92B42D20E98A3CB6CADF5BAA89595C2DCA906005EB7913607D69A5')
Check 'Confirmation summary is not paid authorization' ($c.saved.phase -eq 'confirmation' -and $c.authorization_status -eq 'not authorized' -and @(Get-ChildItem (Join-Path $dir 'approvals') -Filter '*.authorization.json').Count -eq 0)
Change 5 1
Check 'Old tab summary cannot authorize changed selections' (Reject {Confirm-LFAvatarAuthorization $root $id $review.id $review.id})
Change 5 3
Check 'Changing back still invalidates the old summary' ((Current).current.binding_sha256 -ne $review.id)
$c=Prepare;$review=$c.review
$first=Confirm-LFAvatarAuthorization $root $id $review.id $review.id
$second=Confirm-LFAvatarAuthorization $root $id $review.id $review.id
Check 'Rapid duplicate confirms return one identical immutable authorization' ($first.record.id -eq $second.record.id -and $first.file_sha256 -eq $second.file_sha256 -and @(Get-ChildItem (Join-Path $dir 'approvals') -Filter '*.authorization.json').Count -eq 1)
Check 'Snapshot contains exact project, preset, narration and script bindings' ($first.record.project_id -eq $id -and $first.record.snapshot.avatar_preset.version -eq 1 -and $first.record.snapshot.expected_new_provider_jobs -eq 12 -and $first.record.snapshot.slides[0].narration.id -and $first.record.snapshot.slides[0].script_revision)
$authPath=Get-LFApprovalPath $dir $review.id 'authorization';$authHash=(Get-FileHash $authPath).Hash
Change 6 1
Check 'Changing narration invalidates unsubmitted authorization' ((Current).authorization_status -eq 'stale')
Check 'Stale authorization file remains byte-identical' ((Get-FileHash $authPath).Hash -eq $authHash)
Change 6 2;$c=Prepare;$review=$c.review;$confirmed=Confirm-LFAvatarAuthorization $root $id $review.id $review.id
# A receipt fixture exercises the future D duplicate guard without consuming or submitting any real job.
Write-LFImmutableJson (Get-LFApprovalPath $dir $review.id 'consumed') @{test_fixture=$true;authorization_sha256=$confirmed.file_sha256;provider_calls=0}
Check 'Consumed authorization cannot be resubmitted' (Reject {Confirm-LFAvatarAuthorization $root $id $review.id $review.id})
Change 7 1;Change 7 2;$c=Prepare;$draft=$c.review;$draftHash=(Get-FileHash (Get-LFApprovalPath $dir $draft.id 'review')).Hash
& (Join-Path $repo 'Start-LectureForge.ps1') -RuntimeRoot $root -Port $Port -NoBrowser
$url="http://127.0.0.1:$Port/"
function GetJson($Path){Invoke-RestMethod -UseBasicParsing -Uri ($url+$Path) -TimeoutSec 30}
function PostJson($Path,$Data){Invoke-RestMethod -UseBasicParsing -Method Post -Uri ($url+$Path) -Headers @{'X-LF-Token'=$script:token} -ContentType application/json -Body ([Text.Encoding]::UTF8.GetBytes(($Data|ConvertTo-Json -Depth 20))) -TimeoutSec 30}
function Restart {
 $null=PostJson 'api/stop' @{}
 for($i=0;$i -lt 60;$i++){Start-Sleep -Milliseconds 250;$runtime=Read-LFSharedJson (Join-Path $root 'runtime.json');if($runtime.status -eq 'stopped'){break}}
 if($runtime.status -ne 'stopped'){throw 'Server did not stop.'}
 & (Join-Path $repo 'Start-LectureForge.ps1') -RuntimeRoot $root -Port $Port -NoBrowser
 $script:token=(GetJson 'api/bootstrap').token
}
$script:token=(GetJson 'api/bootstrap').token
Check 'HTTP Review Selections exposes exact mapping, counts and pause' ((GetJson "api/selections?id=$id").current.snapshot.slides[6].preparation.post_speech_silence_seconds -eq 6)
Restart;$after=GetJson "api/selections?id=$id"
Check 'Unconfirmed confirmation state survives restart unchanged' ($after.saved.phase -eq 'confirmation' -and $after.review.id -eq $draft.id -and (Get-FileHash (Get-LFApprovalPath $dir $draft.id 'review')).Hash -eq $draftHash)
$confirmed=PostJson "api/avatar-authorize?id=$id" @{confirm=$true;review_id=$draft.id;expected=$draft.id}
$repeat=PostJson "api/avatar-authorize?id=$id" @{confirm=$true;review_id=$draft.id;expected=$draft.id}
Check 'HTTP double-click is idempotent without a second authorization' ($confirmed.file_sha256 -eq $repeat.file_sha256)
$savedState=(Get-FileHash (Join-Path $dir 'narration.json')).Hash
Restart;$after=GetJson "api/selections?id=$id"
Check 'Confirmed authorization and narration survive restart exactly' ($after.authorization_status -eq 'authorized' -and (Get-FileHash (Get-LFApprovalPath $dir $draft.id 'authorization')).Hash -eq $confirmed.file_sha256 -and (Get-FileHash (Join-Path $dir 'narration.json')).Hash -eq $savedState)
Check 'No provider calls, submissions or avatar worker consumption' ((Get-LFNarration $root $id).provider_calls -eq 0 -and $after.authorization.provider_jobs_submitted -eq 0 -and -not $after.provider_submission_enabled -and -not (Test-Path (Get-LFApprovalPath $dir $draft.id 'consumed')))
Check 'Native source copy remains unchanged' ((Get-FileHash (Join-Path $dir 'source.pptx')).Hash -eq $p.source.sha256)
Write-LFJson (Join-Path $root 'results.json') @{passed=$true;checks=@($checks);project_id=$id;project_root=$dir;url=$url;authorization_id=$draft.id;authorization_sha256=$confirmed.file_sha256;expected_new_jobs=12;reusable_avatars=1;elevenlabs_calls=0;heygen_calls=0}
Write-Output "RESULTS=$root/results.json"
Write-Output "PROJECT=$dir"
