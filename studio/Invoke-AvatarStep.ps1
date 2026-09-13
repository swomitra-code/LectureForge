param([Parameter(Mandatory=$true)][string]$RuntimeRoot,[Parameter(Mandatory=$true)][string]$Id,[Parameter(Mandatory=$true)][string]$Job)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'AvatarProduction.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'HeyGenProvider.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'WhiteAvatar.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Render.psm1') -Force
$root=Get-LFRoot $RuntimeRoot;$dir=Get-LFProjectPath $root $Id
function Update($Values){Invoke-LFNarrationLock $root $Id {param($d) $j=Read-LFAvatarJob $d $Job;foreach($key in $Values.Keys){$j.$key=$Values[$key]};if($Values.ContainsKey('stage')){$j.history+=,$Values.stage};Save-LFAvatarJob $d $j};$script:j=Read-LFAvatarJob $dir $Job}
function ProviderAllowed {
 $policy=Join-Path $root 'avatar-policy.json'
 if(Test-Path $policy){$v=Get-Content -Raw $policy|ConvertFrom-Json;if($v.provider_calls_enabled -eq $false){throw 'Provider calls disabled for this isolated fixture runtime.'}}
}
function Request($Method,$Path,$Body=$null){ProviderAllowed;Update @{provider_calls=($j.provider_calls+1)};Invoke-LFHeyGen $Method $Path $Body}
function CheckInputs { $approved=Assert-LFAvatarInputs $root $Id $dir $j.authorization;Assert-LFAvatarJobBinding $approved.snapshot $j;$receipt=Get-Content -Raw (Get-LFApprovalPath $dir $j.authorization consumed)|ConvertFrom-Json;if($receipt.consumer -ne 'lectureforge-avatar-worker-1' -or $receipt.authorization_sha256 -ne (Get-FileHash (Get-LFApprovalPath $dir $j.authorization authorization)).Hash){throw 'Incompatible authorization consumption receipt.'};if((Get-FileHash (Resolve-LFNarrationAsset $dir $j.row.narration.path)).Hash -ne $j.row.narration.sha256){throw 'Authoritative narration changed.'} }
function ValidateWhite($Path,$Evidence){
 [void][IO.Directory]::CreateDirectory($Evidence)
 $n=Resolve-LFNarrationAsset $dir $j.row.narration.path;$np=Get-V2Probe $n;$probe=Get-V2Probe $Path;Assert-V2WhiteAvatarProfile $probe
 $duration=[double]$np.format.duration
 if([math]::Abs([double]$probe.format.duration-$duration) -gt .08){throw 'Full narration duration, including prepared silence, was not preserved.'}
 & ffmpeg -v error -xerror -i $Path -map 0:v:0 -map 0:a:0 -f null NUL
 if($LASTEXITCODE){throw 'White-avatar decode failed.'}
 $time=$duration.ToString('0.000000',[Globalization.CultureInfo]::InvariantCulture);$ref=Join-Path $Evidence 'ref.m4a'
 & ffmpeg -n -v error -i $n -map 0:a:0 -c:a aac -b:a 192k -ar 48000 -ac 2 -t $time $ref
 if($LASTEXITCODE){throw 'Authoritative audio reference failed.'}
 $expected=(& ffmpeg -v error -i $ref -map 0:a:0 -c:a copy -f hash -hash sha256 -) -join ''
 if($LASTEXITCODE){throw 'Reference audio hashing failed.'}
 $actual=(& ffmpeg -v error -i $Path -map 0:a:0 -c:a copy -f hash -hash sha256 -) -join ''
 if($LASTEXITCODE -or $actual -ne $expected){throw 'Selected narration audio packet identity mismatch.'}
 $pause=$j.row.preparation.post_speech_silence_seconds
 if($pause -gt 0){
  $rate=[int]$j.row.preparation.sample_rate;$samples=[long]$j.row.preparation.pause_samples
  if($rate -le 0 -or $samples -ne [long][math]::Round($pause*$rate)){throw 'Prepared pause metadata invalid.'}
  $pcm=Join-Path $Evidence 'n.pcm'; & ffmpeg -n -v error -i $n -map 0:a:0 -ar $rate -ac 1 -f s16le $pcm
  if($LASTEXITCODE){throw 'Prepared narration decode failed.'}
  $stream=[IO.File]::OpenRead($pcm);try{$length=$samples*2;if($stream.Length -lt $length){throw 'Prepared narration too short.'};$null=$stream.Seek(-$length,[IO.SeekOrigin]::End);$bytes=New-Object byte[] $length;$read=$stream.Read($bytes,0,$bytes.Length);if($read -ne $bytes.Length -or @($bytes|Where-Object {$_ -ne 0}).Count){throw 'Prepared trailing silence missing.'}}finally{$stream.Dispose()}
 }
 $motion=Get-V2MotionEvidence $Path (Join-Path $Evidence 'motion.md5')
 @{passed=$true;authoritative_audio_verified=$true;narration_sha256=$j.row.narration.sha256;audio_packet_hash=$actual;duration_seconds=$duration;prepared_pause_seconds=$pause;pause_verified=$true;motion=$motion;output_sha256=(Get-FileHash $Path).Hash}
}
try{
 $j=Read-LFAvatarJob $dir $Job
 Update @{lease=@{owner=$j.lease.owner;pid=$PID;started=(Get-Process -Id $PID).StartTime.ToUniversalTime().Ticks.ToString()}}
 CheckInputs
 if($j.stage -eq 'Queued'){
  if($j.row.reusable_avatar){$r=$j.row.reusable_avatar;$path=Resolve-LFReusableAsset $dir $r.path;if((Get-FileHash $path).Hash -ne $r.sha256){throw 'Reusable avatar hash mismatch.'};Update @{output_path=$r.path;output_sha256=$r.sha256;stage='Validating'}}
  else{
   # Every paid boundary rechecks exact authorization, before persisting intent.
   ProviderAllowed
   $look=Request GET ('/v3/avatars/looks/'+$j.preset.avatar.id)
   if($look.id -ne $j.preset.avatar.id -or $look.status -ne 'completed'){throw 'Approved avatar is unavailable.'}
   if(-not $j.audio_asset_id){Update @{provider_calls=($j.provider_calls+1)};$asset=Send-LFHeyGenAudio (Resolve-LFNarrationAsset $dir $j.row.narration.path);Update @{audio_asset_id=$asset}}
   CheckInputs
   if($j.submission_intent){throw 'Submission uncertain; reconcile existing job. Never resubmit.'}
   $body=@{type='avatar';avatar_id=$j.preset.avatar.id;title=('LectureForge '+$Id+' Slide '+$j.slide+' Take '+$j.take);resolution='1080p';aspect_ratio='16:9';remove_background=$true;output_format='webm';audio_asset_id=$j.audio_asset_id}
   Invoke-LFNarrationLock $root $Id {param($d) $null=Assert-LFAvatarInputs $root $Id $d $j.authorization;$live=Read-LFAvatarJob $d $Job;if($live.submission_intent){throw 'Submission already attempted.'};$live.submission_intent=$true;$live.stage='Submitting to HeyGen';$live.history+=,'Submitting to HeyGen';Save-LFAvatarJob $d $live}
   $j=Read-LFAvatarJob $dir $Job;$created=Request POST '/v3/videos' $body
   if(-not $created.video_id){throw 'No provider job ID returned; submission uncertain.'}
   # Persist ID before any polling or downstream work.
   Update @{video_id=[string]$created.video_id;stage='HeyGen Processing';provider_status='queued';next_poll_utc=[DateTime]::UtcNow.AddSeconds(30).ToString('o')}
  }
 }elseif($j.stage -eq 'HeyGen Processing'){
  if(-not $j.video_id){throw 'Missing persisted HeyGen job ID.'}
  $video=Request GET ('/v3/videos/'+$j.video_id)
  Update @{provider_status=$video.status;provider_terminal=($video.status -in @('completed','failed'));next_poll_utc=[DateTime]::UtcNow.AddSeconds(30).ToString('o')}
  if($video.status -eq 'failed'){throw 'HeyGen job failed. A replacement requires a new explicit authorization.'}
  if($video.status -eq 'completed'){
   if(-not $video.video_url){throw 'Completed job has no download URL.'}
   Update @{download_url=$video.video_url;stage='Downloading Avatar'}
  }
 }elseif($j.stage -eq 'Downloading Avatar'){
  ProviderAllowed;$uri=[uri]$j.download_url;if($uri.Scheme -ne 'https'){throw 'Invalid download URL.'}
  $rel="a/$Job/raw-$([guid]::NewGuid().ToString('N').Substring(0,8)).webm";$path=Join-Path $dir $rel
  Update @{provider_calls=($j.provider_calls+1)}
  Invoke-WebRequest -UseBasicParsing -Uri $uri -OutFile ($path+'.part') -TimeoutSec 180
  if((Get-Item ($path+'.part')).Length -le 0){throw 'Empty avatar download.'}
  [IO.File]::Move(($path+'.part'),$path);Update @{raw_path=$rel;raw_sha256=(Get-FileHash $path).Hash;stage='Avatar Downloaded'}
 }elseif($j.stage -eq 'Avatar Downloaded'){
  CheckInputs
  $suffix=[guid]::NewGuid().ToString('N').Substring(0,8);$out="a/$Job/w-$suffix.mp4";$ev=Join-Path $dir "a/$Job/e-$suffix"
  Update @{stage='White Avatar Building';attempt=($j.attempt+1)}
  $v=New-V2WhiteAvatar -Avatar (Resolve-LFAvatarPath $dir $j.raw_path) -AvatarSha256 $j.raw_sha256 -Narration (Resolve-LFNarrationAsset $dir $j.row.narration.path) -NarrationSha256 $j.row.narration.sha256 -Output (Join-Path $dir $out) -EvidenceDirectory $ev
  Update @{output_path=$out;output_sha256=$v.output_sha256;stage='Validating'}
 }elseif($j.stage -eq 'Validating'){
  $path=Resolve-LFAvatarPath $dir $j.output_path
  if((Get-FileHash $path).Hash -ne $j.output_sha256){throw 'White avatar changed.'}
  $v=ValidateWhite $path (Join-Path $dir "a/$Job/v-$([guid]::NewGuid().ToString('N').Substring(0,8))")
  CheckInputs;Update @{validation=$v;stage='Avatar Ready'}
 }
 Update @{lease=$null}
}catch{
 $message=$_.Exception.Message;foreach($key in @($env:HEYGEN_API_KEY,$env:ELEVENLABS_API_KEY)){if($key){$message=$message.Replace($key,'[REDACTED]')}}
 $j=Read-LFAvatarJob $dir $Job;$kind=if($j.submission_intent -and -not $j.video_id){'submission uncertain'}elseif($j.stage -match 'HeyGen|Downloading'){'provider recovery'}else{'local validation'}
 Set-LFAvatarException $root $Id $Job $message $kind
 exit 1
}
