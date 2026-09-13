Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Project.psm1')
Import-Module (Join-Path $PSScriptRoot 'Render.psm1')

function Assert-V2WhiteAvatarProfile($Probe) {
    $v=@($Probe.streams | Where-Object codec_type -eq 'video')
    $a=@($Probe.streams | Where-Object codec_type -eq 'audio')
    if($v.Count -ne 1 -or $a.Count -ne 1){throw 'Expected one avatar video and one narration stream.'}
    if($v[0].codec_name -ne 'h264' -or $v[0].pix_fmt -ne 'yuv420p' -or $v[0].width -ne 640 -or $v[0].height -ne 560 -or $a[0].codec_name -ne 'aac'){throw 'Expected the compact 640x560 H.264/AAC white-avatar profile; full-slide media is not accepted.'}
}

function Get-V2MotionEvidence([string]$Path,[string]$FrameHashPath,[string]$Crop='130:100:270:175',[switch]$TransparentSource) {
    if(Test-Path -LiteralPath $FrameHashPath){throw 'Motion evidence already exists.'}
    $args=@('-n','-v','error','-threads','2')
    if($TransparentSource){$args+=@('-c:v','libvpx-vp9')}
    $args+=@('-i',$Path,'-vf',"crop=$Crop",'-an','-f','framemd5',$FrameHashPath)
    & ffmpeg @args
    if($LASTEXITCODE -ne 0){throw 'Decoded motion analysis failed.'}
    $hashes=@(Get-Content -LiteralPath $FrameHashPath | Where-Object {$_ -notmatch '^#' -and $_.Trim()} | ForEach-Object {($_ -split ',')[-1].Trim()})
    $unique=@($hashes | Sort-Object -Unique).Count
    if($hashes.Count -lt 2 -or $unique -lt 2){throw 'Static or missing decoded presenter frames.'}
    [pscustomobject]@{decoded_frames=$hashes.Count;distinct_mouth_frames=$unique;motion_present=$true;note='Technical changing-pixel check; instructor confirms visible articulation and lip-sync.'}
}

function New-V2WhiteAvatar {
    param(
        [Parameter(Mandatory=$true)][string]$Avatar,
        [Parameter(Mandatory=$true)][string]$AvatarSha256,
        [Parameter(Mandatory=$true)][string]$Narration,
        [Parameter(Mandatory=$true)][string]$NarrationSha256,
        [Parameter(Mandatory=$true)][string]$Output,
        [Parameter(Mandatory=$true)][string]$EvidenceDirectory
    )
    $avatarAsset=Get-V2Asset $Avatar;$narrationAsset=Get-V2Asset $Narration
    if($avatarAsset.sha256 -ne $AvatarSha256 -or $narrationAsset.sha256 -ne $NarrationSha256){throw 'Explicit avatar/narration hash mismatch.'}
    if($avatarAsset.path -eq $narrationAsset.path){throw 'Use the separate authoritative narration asset; no automatic avatar-audio fallback.'}
    $out=[IO.Path]::GetFullPath($Output);$evidence=[IO.Path]::GetFullPath($EvidenceDirectory)
    if(Test-Path -LiteralPath $out){throw 'Output exists; refusing overwrite.'}
    if([IO.Path]::GetExtension($out) -ne '.mp4' -or (Test-Path -LiteralPath $evidence)){throw 'Use a new MP4 and evidence directory.'}
    $ap=Get-V2Probe $avatarAsset.path;$np=Get-V2Probe $narrationAsset.path
    $video=@($ap.streams | Where-Object codec_type -eq 'video')
    if($video.Count -ne 1 -or $video[0].codec_name -ne 'vp9' -or $video[0].width -ne 1920 -or $video[0].height -ne 1080 -or $video[0].tags.alpha_mode -ne '1'){throw 'Expected the local transparent VP9 fixture.'}
    $duration=[double]::Parse([string]$np.format.duration,[Globalization.CultureInfo]::InvariantCulture)
    if($duration -le 0 -or [math]::Abs([double]$ap.format.duration-$duration) -gt 1){throw 'Narration/avatar duration mismatch; do not stretch or truncate narration.'}
    New-Item -ItemType Directory -Path $evidence | Out-Null
    New-Item -ItemType Directory -Path (Split-Path -Parent $out) -Force | Out-Null
    $time=$duration.ToString('0.000000',[Globalization.CultureInfo]::InvariantCulture)
    # Avatar input 0 is VIDEO ONLY. Explicit authoritative narration is input 1.
    # Restore alpha after RGB despill, then composite the compact tile onto white.
    $filter='[0:v]crop=640:560:870:70,format=rgba,split[c][a];[a]alphaextract,geq=lum=p(X\,Y)*clip((559-Y)/70\,0\,1)[alpha];[c]despill=type=green:mix=0.35:expand=0:alpha=0[rgb];[rgb][alpha]alphamerge,tpad=stop_mode=clone:stop_duration=1[fg];color=c=white:s=640x560:r=25[bg];[bg][fg]overlay=format=auto,scale=in_range=pc:out_range=tv,format=yuv420p[v]'
    $args=@('-n','-v','error','-threads','2','-c:v','libvpx-vp9','-i',$avatarAsset.path,'-i',$narrationAsset.path,'-filter_complex_threads','1','-filter_complex',$filter,'-map','[v]','-map','1:a:0','-c:v','libx264','-threads','2','-preset','medium','-crf','18','-color_range','tv','-colorspace','bt709','-color_primaries','bt709','-color_trc','bt709','-c:a','aac','-b:a','192k','-ar','48000','-ac','2','-t',$time,'-movflags','+faststart',$out)
    & ffmpeg @args
    if($LASTEXITCODE -ne 0){throw 'White-avatar composition failed.'}
    $probe=Get-V2Probe $out;Assert-V2WhiteAvatarProfile $probe
    if([math]::Abs([double]$probe.format.duration-$duration) -gt .08){throw 'Narration duration was not preserved.'}
    & ffmpeg -v error -i $out -map 0:v:0 -map 0:a:0 -f null NUL
    if($LASTEXITCODE -ne 0){throw 'White avatar full decode failed.'}
    $ref=Join-Path $evidence 'narration-reference.m4a'
    & ffmpeg -n -v error -i $narrationAsset.path -map 0:a:0 -c:a aac -b:a 192k -ar 48000 -ac 2 -t $time $ref
    if($LASTEXITCODE -ne 0){throw 'Narration reference encode failed.'}
    $actual=(& ffmpeg -v error -i $out -map 0:a:0 -c:a copy -f hash -hash sha256 -) -join ''
    if($LASTEXITCODE -ne 0){throw 'Audio hashing failed.'}
    $expected=(& ffmpeg -v error -i $ref -map 0:a:0 -c:a copy -f hash -hash sha256 -) -join ''
    if($LASTEXITCODE -ne 0 -or $actual -ne $expected){throw 'Authoritative narration packet mismatch.'}
    $motion=Get-V2MotionEvidence $out (Join-Path $evidence 'mouth.framemd5')
    if((Get-V2Asset $avatarAsset.path).sha256 -ne $AvatarSha256 -or (Get-V2Asset $narrationAsset.path).sha256 -ne $NarrationSha256){throw 'Input asset changed.'}
    $report=[ordered]@{output=$out;output_sha256=(Get-V2Asset $out).sha256;avatar_sha256=$AvatarSha256;narration_sha256=$NarrationSha256;narration_audio_map='1:a:0';avatar_audio_mapped=$false;audio_routing_verified=$true;audio_packet_hash=$actual;ffmpeg_arguments=$args;duration=$duration;tile_width=640;tile_height=560;background_rgb=@(255,255,255);motion=$motion;provider_calls=0;production_audio_listening='pending instructor';framing_issue='Lower 70 pixels fade smoothly into white; head and shoulder alpha retained.'}
    Write-V2Json (Join-Path $evidence 'white-avatar-validation.json') $report
    [pscustomobject]$report
}

Export-ModuleMember -Function New-V2WhiteAvatar,Assert-V2WhiteAvatarProfile,Get-V2MotionEvidence
