param([string]$EvidenceDirectory=(Join-Path $PSScriptRoot ('../../work/recording-tests-'+[DateTime]::UtcNow.ToString('yyyyMMddHHmmss'))))
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$repo=(Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
Import-Module (Join-Path $repo 'studio/Project.psm1') -Force
Import-Module (Join-Path $repo 'studio/Render.psm1') -Force
Import-Module (Join-Path $repo 'studio/WhiteAvatar.psm1') -Force
Import-Module (Join-Path $repo 'studio/Recording.psm1') -Force
$dir=[IO.Path]::GetFullPath($EvidenceDirectory)
if(Test-Path $dir){throw 'Test evidence directory already exists.'}
New-Item -ItemType Directory -Path $dir | Out-Null
$checks=[Collections.Generic.List[object]]::new()
function Check($Name,$Pass){$checks.Add([pscustomobject]@{test=$Name;passed=[bool]$Pass});if(-not $Pass){throw "FAILED: $Name"}}
function Reject($Name,[scriptblock]$Action){$failed=$false;try{& $Action | Out-Null}catch{$failed=$true};Check $Name $failed}
$source=Join-Path $repo 'source/RE-M2-Solar-Thermal-Overview.pptx'
$avatar=Join-Path $repo 'media/energy-management/slide-1-SKM-BLUE-7-production-transparent.webm'
$narration=Join-Path $repo 'media/energy-management/slide-1-take-1.mp3'
$approved=Join-Path $repo 'work/avatar-white-proof-20260912/slide-1-avatar-white.mp4'
$acceptedDeck=Join-Path $repo 'output/RE-M2-Solar-Thermal-Overview_RECORDING_AVATAR_WHITE_TEST_20260912.pptx'
$before=@{};foreach($p in $source,$avatar,$narration,$approved,$acceptedDeck){$before[$p]=(Get-V2Asset $p).sha256}
Check 'Canonical source matches recorded Solar Thermal master' ($before[$source] -eq '6340C137FD85021957721341298D640C8FAF59D36507D904DF09948DCB576C01')
Check 'Approved white tile matches instructor-reviewed fixture' ($before[$approved] -eq '3D9420617F53F0172021B22665A81BCB4D0FA83683CF06C29C5954AE2FA8E2B8')
foreach($path in @(Get-ChildItem (Join-Path $repo 'studio') -File | Where-Object Extension -in '.ps1','.psm1')){
 $tokens=$null;$errors=$null;[void][Management.Automation.Language.Parser]::ParseFile($path.FullName,[ref]$tokens,[ref]$errors)
 Check ("Syntax: "+$path.Name) ($errors.Count -eq 0)
 Check ("Offline only: "+$path.Name) (-not [regex]::IsMatch([IO.File]::ReadAllText($path.FullName),'(?i)Invoke-WebRequest|Invoke-RestMethod|HttpClient|generate-avatar\.ps1|generate-narration\.ps1|curl(?:\.exe)?'))
}
$build=@{Avatar=$avatar;AvatarSha256=$before[$avatar];Narration=$narration;NarrationSha256=$before[$narration];Output=(Join-Path $dir 'new-authoritative-avatar.mp4');EvidenceDirectory=(Join-Path $dir 'white')}
$bad=$build.Clone();$bad.NarrationSha256='WRONG'
Reject 'Changed narration hash blocks composition' {New-V2WhiteAvatar @bad}
$bad=$build.Clone();$bad.Narration=$avatar;$bad.NarrationSha256=$before[$avatar]
Reject 'Avatar audio cannot silently replace authoritative narration' {New-V2WhiteAvatar @bad}
$white=New-V2WhiteAvatar @build
Check 'New white composition uses independent authoritative narration' ($white.audio_routing_verified -and -not $white.avatar_audio_mapped -and $white.narration_audio_map -eq '1:a:0')
Check 'White composition preserves narration duration without shortest' ($white.ffmpeg_arguments -notcontains '-shortest')
Reject 'Existing white output cannot be overwritten' {New-V2WhiteAvatar @build}
$sm=Get-V2MotionEvidence $avatar (Join-Path $dir 'source-mouth.framemd5') '130:100:1140:245' -TransparentSource
$wm=Get-V2MotionEvidence $approved (Join-Path $dir 'approved-mouth.framemd5')
Check 'Source and approved white tile retain changing mouth frames' ($sm.distinct_mouth_frames -gt 100 -and $wm.distinct_mouth_frames -gt 100 -and $white.motion.distinct_mouth_frames -gt 100)
# Inspect decoded white corners throughout the clip (not codec/color metadata).
Add-Type -AssemblyName System.Drawing
foreach($item in @(@{name='approved';path=$approved},@{name='new';path=$white.output})){
 foreach($time in 1,10,25){
  $png=Join-Path $dir ($item.name+"-$time.png")
  & ffmpeg -n -v error -ss $time -i $item.path -frames:v 1 $png
  if($LASTEXITCODE -ne 0){throw 'Review frame decode failed.'}
  $b=[Drawing.Bitmap]::new($png);$isWhite=$true
  try{foreach($pt in @(@(0,0),@(320,0),@(639,0),@(0,559),@(639,559))){$c=$b.GetPixel($pt[0],$pt[1]);if($c.R -ne 255 -or $c.G -ne 255 -or $c.B -ne 255){$isWhite=$false}}}finally{$b.Dispose()}
  Check ($item.name+" decoded white background at ${time}s") $isWhite
 }
}
$pack=@{Source=$source;ApprovedVideo=$approved;ApprovedVideoSha256=$before[$approved];Output=(Join-Path $dir 'SolarThermal_RECORDING_AVATAR_WHITE.pptx');EvidenceDirectory=(Join-Path $dir 'recording');Placement=@{left=438;top=337;width=110;height=96.25}}
$bad=$pack.Clone();$bad.ApprovedVideoSha256='WRONG'
Reject 'Unapproved video hash blocks recording' {New-V2RecordingProof @bad}
$bad=$pack.Clone();$bad.Output=$source
Reject 'Canonical source cannot be recording output' {New-V2RecordingProof @bad}
$bad=$pack.Clone();$bad.Placement=@{left=0;top=0;width=960;height=540}
Reject 'Full-slide avatar placement rejected' {New-V2RecordingProof @bad}
$composite=Get-V2Probe (Join-Path $repo 'work/v2-phase2a/solar-thermal/videos/slide_01_candidate_003.mp4')
Reject 'Full-slide composite profile rejected by recording path' {Assert-V2WhiteAvatarProfile $composite}
$records=@(New-V2RecordingProof @pack);$record=$records[-1]
Check 'Recording generated and reopened with all 14 native slides' ($record.automated_checks_passed -and $record.output_slide_count -eq 14 -and $record.slide1_native_content_preserved)
Check 'Slides 2-14 remain structurally and visually identical' (@($record.unrelated_slides | Where-Object {-not $_.structure_unchanged -or -not $_.render_identical}).Count -eq 0 -and $record.unrelated_slides.Count -eq 13)
Check 'Avatar remains independently movable/resizable with correct placement' ($record.independent_move_resize -and $record.avatar_geometry_matches -and $record.border_and_shadow_hidden)
Check 'Embedded media bytes/audio and motion preserved' ($record.embedded_video_sha256 -eq $before[$approved] -and $record.audio_stream_present -and $record.embedded_motion.distinct_mouth_frames -gt 100)
Check 'Autoplay persists and advances in desktop PowerPoint' ($record.autoplay_configured -and $record.play_on_entry -and $record.live_autoplay_advanced)
Reject 'Existing Recording deck cannot be overwritten' {New-V2RecordingProof @pack}
# Recheck the actual instructor-reviewed deck independently, without opening/saving it.
$accepted=Join-Path $dir 'accepted-package';New-Item -ItemType Directory -Path $accepted | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip=[IO.Compression.ZipFile]::OpenRead($acceptedDeck)
try{[IO.Compression.ZipFileExtensions]::ExtractToFile($zip.GetEntry('ppt/media/media1.mp4'),(Join-Path $accepted 'avatar.mp4'))}finally{$zip.Dispose()}
Check 'Instructor-reviewed deck still embeds exact accepted tile' ((Get-V2Asset (Join-Path $accepted 'avatar.mp4')).sha256 -eq $before[$approved])
foreach($p in $before.Keys){Check ('Immutable input: '+[IO.Path]::GetFileName($p)) ((Get-V2Asset $p).sha256 -eq $before[$p])}
$result=[ordered]@{passed=$true;count=$checks.Count;checks=@($checks);recording=$pack.Output;source_sha256=$before[$source];white_candidate=$white.output;provider_calls=0;manual_acceptance='Instructor approved original white workflow and visible motion; final real narration audible confirmation pending.'}
Write-V2Json (Join-Path $dir 'test-results.json') $result
$result | ConvertTo-Json -Depth 7
