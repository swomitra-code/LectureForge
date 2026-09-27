param([string]$RuntimeRoot=('C:/LectureForge/work/reconcile-'+[guid]::NewGuid().ToString('N').Substring(0,8)))
$ErrorActionPreference='Stop'
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Import-Module (Join-Path $repo 'studio/AvatarProduction.psm1') -Force
$root=Get-LFRoot $RuntimeRoot;$id='abcdef123456';$dir=Get-LFProjectPath $root $id
[void][IO.Directory]::CreateDirectory((Join-Path $dir 'approvals'))
[void][IO.Directory]::CreateDirectory((Join-Path $dir 'n/aaaaaaaaaaaa/bbbbbbbbbbbb'))
$audio='n/aaaaaaaaaaaa/bbbbbbbbbbbb/take.mp3';[IO.File]::WriteAllText((Join-Path $dir $audio),'fixture narration')
$row=[pscustomobject]@{number=16;selected_take=2;narration=@{path=$audio;sha256=(Get-FileHash (Join-Path $dir $audio)).Hash};placement_override=$null;reusable_avatar=$null}
$preset=Get-LFAvatarPreset ([pscustomobject]@{preset=Get-LFPreset})
$snapshot=@{slides=@($row);avatar_preset=$preset;avatar_preset_sha256=(Get-LFTextHash (ConvertTo-LFCanonicalJson $preset));default_placement=@{left=.8};project_revision=1;intent_revision=0;can_authorize=$true}
$auth=Get-LFTextHash (ConvertTo-LFCanonicalJson $snapshot)
Write-LFJson (Get-LFApprovalPath $dir $auth authorization) @{id=$auth;binding_sha256=$auth;snapshot=$snapshot}
Write-LFJson (Get-LFApprovalPath $dir $auth consumed) @{consumer='lectureforge-avatar-worker-1';authorization=$auth;authorization_sha256=(Get-FileHash (Get-LFApprovalPath $dir $auth authorization)).Hash}
$global:lfRecoverySelection=[pscustomobject]@{snapshot=$snapshot;binding_sha256=$auth}
# Mock selection discovery only; immutable authorization, receipt and narration checks stay real.
& (Get-Module AvatarProduction) {function script:Get-LFSelectionInputs {param($Root,$Id,$Dir) $global:lfRecoverySelection}}
$global:lfRecoveryCalls=[Collections.Generic.List[string]]::new();$global:lfRecoveryMode='completed'
$global:lfRecoveryVideo='1234567890abcdef1234567890abcdef'
$global:lfRecoveryTitle='LectureForge abcdef123456 Slide 16 Take 2'
if(-not $env:HEYGEN_API_KEY){$env:HEYGEN_API_KEY='isolated-test-key'}
function global:Invoke-RestMethod {
 param($Method,$Uri,$Headers,$TimeoutSec,$UseBasicParsing,$ContentType,$Body)
 $global:lfRecoveryCalls.Add("$Method $Uri")
 if($Method -ne 'GET' -or $Uri -notmatch '^https://api\.heygen\.com/v3/videos/[a-f0-9]{32}$'){throw 'SAFETY FAILURE: creation or unexpected provider endpoint called'}
 if($global:lfRecoveryMode -eq 'network'){throw 'transient network failure'}
 if($global:lfRecoveryMode -eq 'missing'){return @{data=$null}}
 @{data=[pscustomobject]@{id=$global:lfRecoveryVideo;title=$global:lfRecoveryTitle;status=$global:lfRecoveryMode;video_page_url=('https://app.heygen.com/videos/'+$global:lfRecoveryVideo);video_url='https://cdn.example.test/existing.webm'}}
}
$jid='1111111111111111';$alias='2222222222222222';$url='https://app.heygen.com/videos/lectureforge-abcdef123456-slide-16-take-2-'+$global:lfRecoveryVideo
function InvokeFixtureStep {
 $code=[IO.File]::ReadAllText((Join-Path $repo 'studio/Invoke-AvatarStep.ps1'))
 $code=$code.Replace("Import-Module (Join-Path `$PSScriptRoot 'AvatarProduction.psm1') -Force",'')
 $code=$code.Replace('$PSScriptRoot',("'"+(Join-Path $repo 'studio')+"'"))
 & ([scriptblock]::Create($code)) -RuntimeRoot $root -Id $id -Job $jid
}
function Check($Name,[bool]$OK){if(-not $OK){throw "FAIL: $Name"};Write-Output "PASS: $Name"}
function ResetJob {
 $j=[pscustomobject]@{id=$jid;authorization=$auth;slide=16;take=2;row=$row;preset=$preset;placement=@{left=.8};stage='Exception';history=@('Authorized','Queued','Submitting to HeyGen');error='Submission uncertain';exception_kind='submission uncertain';lease=$null;submission_intent=$true;video_id=$null;audio_asset_id='audio-original';provider_status=$null;provider_terminal=$false;provider_calls=1;next_poll_utc=$null;download_url=$null;raw_path=$null;raw_sha256=$null;output_path=$null;output_sha256=$null;validation=$null;attempt=0;updated_utc=$null}
 [void][IO.Directory]::CreateDirectory((Split-Path (Get-LFAvatarJobPath $dir $jid)));Save-LFAvatarJob $dir $j
 Write-LFJson (Join-Path $dir 'avatar-queue.json') @{schema='lectureforge-avatar-queue-1';paused=$false;batches=@(@{authorization=$auth;jobs=@($jid)});jobs=@($jid)}
}
function Reject($Name,$TestUrl=$url){$before=(Get-FileHash (Get-LFAvatarJobPath $dir $jid)).Hash;$failed=$false;try{$null=Recover-LFAvatarVideo $root $id $jid $TestUrl}catch{$failed=$true};Check $Name ($failed -and (Get-FileHash (Get-LFAvatarJobPath $dir $jid)).Hash -eq $before)}
ResetJob
foreach($bad in @('hello',('https://evil.test/videos/'+$global:lfRecoveryVideo),('http://app.heygen.com/videos/'+$global:lfRecoveryVideo),('https://app.heygen.com@evil.test/videos/'+$global:lfRecoveryVideo))){Reject 'Malformed or unsafe URL leaves original untouched' $bad}
Check 'Malformed URLs make zero provider calls' ($global:lfRecoveryCalls.Count -eq 0)
$global:lfRecoveryTitle='LectureForge abcdef123456 Slide 15 Take 2';Reject 'Mismatched slide rejected'
$global:lfRecoveryTitle='LectureForge abcdef123456 Slide 16 Take 1';Reject 'Mismatched take rejected'
$global:lfRecoveryTitle='LectureForge abcdef123456 Slide 16 Take 2'
$global:lfRecoveryMode='missing';Reject 'Video not found leaves submission unchanged'
$global:lfRecoveryMode='network';Reject 'Transient network error leaves submission unchanged'
$global:lfRecoveryMode='completed';$before=Read-LFAvatarJob $dir $jid;$null=Recover-LFAvatarVideo $root $id $jid $url;$after=Read-LFAvatarJob $dir $jid
Check 'Completed recovery resumes normal download stage' ($after.stage -eq 'Downloading Avatar' -and $after.video_id -eq $global:lfRecoveryVideo -and $after.download_url -and $after.provider_terminal)
Check 'Original authorization, narration, take, placement and lineage preserved' ($after.authorization -eq $before.authorization -and (ConvertTo-LFCanonicalJson $before.row) -eq (ConvertTo-LFCanonicalJson $after.row) -and (ConvertTo-LFCanonicalJson $before.placement) -eq (ConvertTo-LFCanonicalJson $after.placement) -and $after.audio_asset_id -eq $before.audio_asset_id -and $after.recovery.submission -eq $jid -and 'ProviderJobRecovered' -in $after.history -and (Read-LFAvatarQueue $dir).jobs.Count -eq 1)
$hash=(Get-FileHash (Get-LFAvatarJobPath $dir $jid)).Hash;$calls=$global:lfRecoveryCalls.Count;$again=Recover-LFAvatarVideo $root $id $jid $url
Check 'Repeated recovery is byte-idempotent and makes no API calls' ($again.already_reconciled -and $global:lfRecoveryCalls.Count -eq $calls -and (Get-FileHash (Get-LFAvatarJobPath $dir $jid)).Hash -eq $hash)
Reject 'Already reconciled submission rejects a different asset' ('https://app.heygen.com/videos/'+('f'*32))
ResetJob;$global:lfRecoveryMode='processing';$null=Recover-LFAvatarVideo $root $id $jid $url
$j=Read-LFAvatarJob $dir $jid
Check 'Processing asset resumes ordinary polling' ($j.stage -eq 'HeyGen Processing' -and -not $j.provider_terminal -and -not $j.next_poll_utc)
# Exercise the real worker status branch with a GET-only transport double.
InvokeFixtureStep
$j=Read-LFAvatarJob $dir $jid
Check 'Real worker polls recovered processing job without generating' ($j.stage -eq 'HeyGen Processing' -and $j.next_poll_utc -and -not $j.lease)
$global:lfRecoveryMode='completed'
InvokeFixtureStep
Check 'Real worker advances recovered completed video to download' ((Read-LFAvatarJob $dir $jid).stage -eq 'Downloading Avatar')
ResetJob;$global:lfRecoveryMode='failed';$null=Recover-LFAvatarVideo $root $id $jid $url;$j=Read-LFAvatarJob $dir $jid
Check 'Failed provider asset recorded without replacement' ($j.stage -eq 'Exception' -and $j.provider_status -eq 'failed' -and $j.video_id -and $j.provider_terminal)
ResetJob;$j=Read-LFAvatarJob $dir $jid;$j.video_id=$global:lfRecoveryVideo;$j.exception_kind='stale authorization';Save-LFAvatarJob $dir $j
$blocked=Read-LFAvatarJob $dir $jid;$blocked.id=$alias;$blocked.video_id=$null;$blocked.submission_intent=$false;$blocked.exception_kind='existing submission';$blocked.error="Existing provider submission $jid must be reconciled before any replacement."
[void][IO.Directory]::CreateDirectory((Split-Path (Get-LFAvatarJobPath $dir $alias)));Save-LFAvatarJob $dir $blocked
$q=Read-LFAvatarQueue $dir;$q.jobs+=,$alias;Write-LFJson (Join-Path $dir 'avatar-queue.json') $q
$global:lfRecoveryMode='completed';$null=Recover-LFAvatarVideo $root $id $alias $url
Check 'Blocked newer row recovers original already-known submission' ((Read-LFAvatarJob $dir $jid).stage -eq 'Downloading Avatar' -and -not (Read-LFAvatarJob $dir $alias).video_id -and (Read-LFAvatarQueue $dir).jobs.Count -eq 2)
Write-LFJson (Join-Path $dir 'project.json') @{id=$id;name='recovery fixture';version=1;source=@{slide_count=16}}
Check 'Dashboard follows recovered original without duplicate processing' ((Get-LFAvatarStatus $root $id).jobs[0].id -eq $jid)
$global:lfRecoverySelection.binding_sha256='later-authorization'
$tasks=@(Get-LFAvatarTasks $root 'recovery-test' 1)
Check 'Scheduler resumes original despite later authorization and never schedules blocked alias' ($tasks.Count -eq 1 -and $tasks[0].job -eq $jid -and -not $tasks[0].new_provider)
$j=Read-LFAvatarJob $dir $jid;$j.lease=$null;Save-LFAvatarJob $dir $j
$j.stage='Queued';$rejected=$false;try{Assert-LFAvatarRecoveryInputs $root $id $dir $j}catch{$rejected=$true};Check 'Recovery authorization cannot enter a paid generation stage' $rejected
$global:lfRecoverySelection.binding_sha256=$auth
$global:lfRecoverySelection.snapshot.slides[0].selected_take=3
$rejected=$false;try{Assert-LFAvatarRecoveryInputs $root $id $dir (Read-LFAvatarJob $dir $jid)}catch{$rejected=$true};Check 'Changed selected take blocks recovered processing' $rejected
$global:lfRecoverySelection.snapshot.slides[0].selected_take=2
# Exercise the original duplicate guard through a newly consumed authorization.
ResetJob
$snapshot.project_revision=2;$nextAuth=Get-LFTextHash (ConvertTo-LFCanonicalJson $snapshot)
Write-LFJson (Get-LFApprovalPath $dir $nextAuth authorization) @{id=$nextAuth;binding_sha256=$nextAuth;snapshot=$snapshot}
$global:lfRecoverySelection.binding_sha256=$nextAuth
Write-LFJson (Join-Path $dir 'selection-review.json') @{phase='authorized';review_id=$nextAuth}
Start-LFAvatarBatch $root $id
$q=Read-LFAvatarQueue $dir;$new=Read-LFAvatarJob $dir $q.jobs[-1]
Check 'Original duplicate-submission safeguard still blocks replacement' ($q.jobs.Count -eq 2 -and $new.exception_kind -eq 'existing submission' -and $new.error -like '*must be reconciled before any replacement*' -and -not $new.submission_intent)
Check 'SAFETY: recovery and resumed polling NEVER call CREATE/generate or upload' (@($global:lfRecoveryCalls|Where-Object {$_ -notmatch '^GET https://api\.heygen\.com/v3/videos/[a-f0-9]{32}$'}).Count -eq 0)
Write-Output "RESULTS: all reconciliation checks passed; $($global:lfRecoveryCalls.Count) mocked GET calls; zero create calls."
