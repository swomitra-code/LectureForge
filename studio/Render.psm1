Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'Project.psm1')

function Invoke-V2FFmpeg([string[]]$Arguments) {
    & ffmpeg @Arguments
    if ($LASTEXITCODE -ne 0) { throw "FFmpeg failed (exit $LASTEXITCODE)." }
}

function Get-V2Probe([string]$Path) {
    $result = & ffprobe -v error -show_streams -show_format -of json $Path
    if ($LASTEXITCODE -ne 0) { throw 'FFprobe failed.' }
    ($result -join "`n") | ConvertFrom-Json
}

function Export-V2Slide([string]$ProjectPath) {
    $project = Read-V2ProofProject $ProjectPath
    $source = Join-Path $project.root $project.record.source.path
    $original = $project.record.source.original_path
    $before = Get-V2Asset $source
    $originalBefore = Get-V2Asset $original
    if ($originalBefore.sha256 -ne $before.sha256) { throw 'Original source changed since import.' }
    $infoBefore = Get-V2DeckInfo $source
    $png = Join-Path $project.root 'renders/slide_01.png'
    if (Test-Path -LiteralPath $png) { throw 'Render exists; use recomposition without opening PowerPoint.' }
    # PowerPoint may reuse an existing application. Close only our imported presentation;
    # never quit the application when a user-owned session was present at entry.
    $applicationWasRunning = @((Get-Process POWERPNT -ErrorAction SilentlyContinue)).Count -gt 0
    $app = $null; $deck = $null; $slide = $null
    $count = 0; $afterCount = 0; $text = @()
    try {
        $app = New-Object -ComObject PowerPoint.Application
        $deck = $app.Presentations.Open($source, $true, $false, $false)
        $count = $deck.Slides.Count
        if ($count -ne 14 -or [math]::Abs($deck.PageSetup.SlideWidth / $deck.PageSetup.SlideHeight - 16.0/9.0) -gt 0.00001) { throw 'COM source geometry mismatch.' }
        $slide = $deck.Slides.Item(1)
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) { $text += $shape.TextFrame.TextRange.Text }
        }
        $slide.Export($png, 'PNG', 1920, 1080)
        $afterCount = $deck.Slides.Count
    } finally {
        if ($deck) { $deck.Close() }
        if ($app -and -not $applicationWasRunning) { $app.Quit() }
        foreach ($com in @($slide,$deck,$app)) { if ($null -ne $com) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com) } }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        if ((Get-V2Asset $source).sha256 -ne $before.sha256 -or (Get-V2Asset $original).sha256 -ne $originalBefore.sha256) { throw 'SOURCE IMMUTABILITY FAILURE.' }
    }
    $infoAfter = Get-V2DeckInfo $source
    $probe = Get-V2Probe $png
    if ($probe.streams[0].width -ne 1920 -or $probe.streams[0].height -ne 1080) { throw 'Unexpected PNG dimensions.' }
    if ($count -ne $afterCount -or $infoBefore.slide_count -ne $infoAfter.slide_count) { throw 'Source slide count changed.' }
    $report = [ordered]@{
        original_sha256_before=$originalBefore.sha256; original_sha256_after=(Get-V2Asset $original).sha256
        imported_sha256_before=$before.sha256; imported_sha256_after=(Get-V2Asset $source).sha256
        slide_count_before=$count; slide_count_after=$afterCount
        package_slide_count_after=$infoAfter.slide_count; opened_read_only=$true
        save_calls=0; embed_calls=0; slide_number=1; render_path=$png
        render_sha256=(Get-V2Asset $png).sha256; width=1920; height=1080; native_text=$text
        static_slide_only=$true; visual_review='pending'
    }
    Write-V2Json (Join-Path $project.root 'validation/source-render.json') $report
    [pscustomobject]$report
}

function Get-V2Placement($Placement, $Crop) {
    foreach ($key in 'x','y','width') { if ([double]::IsNaN([double]$Placement.$key) -or [double]::IsInfinity([double]$Placement.$key)) { throw 'Non-finite placement.' } }
    if ($Placement.x -lt 0 -or $Placement.y -lt 0 -or $Placement.width -le 0 -or $Placement.width -gt 1) { throw 'Invalid normalized placement.' }
    $w = [int][math]::Round(1920 * $Placement.width)
    $h = [int][math]::Round($w * $Crop.height / $Crop.width)
    $x = [int][math]::Round(1920 * $Placement.x); $y = [int][math]::Round(1080 * $Placement.y)
    if ($w -lt 2 -or $h -lt 2 -or $x+$w -gt 1920 -or $y+$h -gt 1080) { throw 'Presenter extends outside slide.' }
    [pscustomobject]@{x=$x;y=$y;width=$w;height=$h}
}

function Get-V2CompositionArguments([string]$SlideImage,[string]$Avatar,[string]$Narration,[string]$Output,$Crop,$Placement,[double]$Duration) {
    $p = Get-V2Placement $Placement $Crop
    $durationText = $Duration.ToString('0.000000',[Globalization.CultureInfo]::InvariantCulture)
    # Input 0 is the static slide, input 1 supplies VIDEO ONLY, input 2 is authoritative audio.
    $filter = "[1:v]crop=$($Crop.width):$($Crop.height):$($Crop.x):$($Crop.y),format=rgba,split=2[rgb0][a0];[a0]alphaextract[a];[rgb0]despill=type=green:mix=0.30:expand=0.05:alpha=0[rgb];[rgb][a]alphamerge,scale=$($p.width):$($p.height):flags=lanczos,tpad=stop_mode=clone:stop_duration=1[p];[0:v][p]overlay=x=$($p.x):y=$($p.y):format=auto:eof_action=repeat[composed];[composed]scale=out_range=tv,format=yuv420p[v]"
    @('-n','-hide_banner','-loglevel','error','-threads','2','-loop','1','-framerate','30','-i',$SlideImage,
      '-threads','2','-c:v','libvpx-vp9','-i',$Avatar,'-i',$Narration,'-filter_complex_threads','1','-filter_complex',$filter,
      '-map','[v]','-map','2:a:0','-c:v','libx264','-threads','2','-preset','medium','-crf','18','-pix_fmt','yuv420p','-color_range','tv','-r','30',
      '-c:a','aac','-b:a','192k','-ar','48000','-ac','2','-t',$durationText,'-movflags','+faststart',$Output)
}

function Assert-V2VideoProfile($Probe,[double]$NarrationDuration) {
    $video = @($Probe.streams | Where-Object codec_type -eq 'video')
    $audio = @($Probe.streams | Where-Object codec_type -eq 'audio')
    if ($video.Count -ne 1 -or $audio.Count -ne 1) { throw 'Expected exactly one video and one audio stream.' }
    $v=$video[0]; $a=$audio[0]
    if ($v.codec_name -ne 'h264' -or $v.width -ne 1920 -or $v.height -ne 1080 -or $v.pix_fmt -ne 'yuv420p' -or $v.avg_frame_rate -ne '30/1' -or $a.codec_name -ne 'aac') { throw 'Incorrect output media profile.' }
    if ([math]::Abs([double]$Probe.format.duration - $NarrationDuration) -gt 0.08) { throw 'Output duration differs from narration by more than 80 ms.' }
}

function Get-V2AudioPacketHash([string]$Path) {
    $hash = & ffmpeg -v error -i $Path -map '0:a:0' -c:a copy -f hash -hash sha256 -
    if ($LASTEXITCODE -ne 0) { throw 'Audio packet hashing failed.' }
    ($hash -join '').Trim()
}

function Build-V2Slide {
    param([string]$ProjectPath,[string]$CandidateName='slide_01_candidate_001',[object]$Placement=$null)
    if ($CandidateName -notmatch '^slide_01_[a-zA-Z0-9_]+$') { throw 'Invalid candidate name.' }
    $project = Read-V2ProofProject $ProjectPath
    $record = $project.record
    if ($null -eq $Placement) { $Placement = $record.candidate_placement }
    $pixels = Get-V2Placement $Placement $record.avatar_crop
    $image = Join-Path $project.root 'renders/slide_01.png'
    $render = Get-Content -Raw (Join-Path $project.root 'validation/source-render.json') | ConvertFrom-Json
    if ((Get-V2Asset $image).sha256 -ne $render.render_sha256) { throw 'Static slide render changed.' }
    $avatar = Join-Path $project.root $record.avatar.path
    $narration = Join-Path $project.root $record.narration.path
    $output = Join-Path $project.root "videos/$CandidateName.mp4"
    if (Test-Path -LiteralPath $output) { throw 'Candidate exists; refusing overwrite.' }
    $audioProbe = Get-V2Probe $narration; $avatarProbe = Get-V2Probe $avatar
    $duration = [double]::Parse([string]$audioProbe.format.duration,[Globalization.CultureInfo]::InvariantCulture)
    if ($duration -le 0 -or [math]::Abs([double]$avatarProbe.format.duration-$duration) -gt 1.0) { throw 'Technical fixtures must have durations within one second; no audio stretching allowed.' }
    $video = @($avatarProbe.streams | Where-Object codec_type -eq 'video')[0]
    if ($video.codec_name -ne 'vp9' -or $video.tags.alpha_mode -ne '1') { throw 'Expected a transparent VP9 fixture.' }
    $crop=$record.avatar_crop
    if ($crop.x -lt 0 -or $crop.y -lt 0 -or $crop.width -le 0 -or $crop.height -le 0 -or $crop.x+$crop.width -gt $video.width -or $crop.y+$crop.height -gt $video.height) { throw 'Invalid avatar crop.' }
    Invoke-V2FFmpeg @('-v','error','-i',$narration,'-f','null','NUL')
    Invoke-V2FFmpeg @('-v','error','-threads','2','-c:v','libvpx-vp9','-i',$avatar,'-f','null','NUL')
    $alphaPath = Join-Path $project.root "validation/$CandidateName-alpha.gray"
    Invoke-V2FFmpeg @('-n','-v','error','-threads','2','-c:v','libvpx-vp9','-i',$avatar,'-vf',"crop=$($crop.width):$($crop.height):$($crop.x):$($crop.y),alphaextract",'-frames:v','1','-an','-f','rawvideo',$alphaPath)
    $alpha=[IO.File]::ReadAllBytes($alphaPath)
    if ($alpha -notcontains 0 -or $alpha -notcontains 255) { throw 'Decoded crop lacks transparent and opaque pixels.' }
    $arguments = @(Get-V2CompositionArguments $image $avatar $narration $output $crop $Placement $duration)
    Invoke-V2FFmpeg $arguments
    $probe = Get-V2Probe $output
    Assert-V2VideoProfile $probe $duration
    Invoke-V2FFmpeg @('-v','error','-i',$output,'-map','0:v:0','-map','0:a:0','-f','null','NUL')
    # Independently encode the explicit narration with the same AAC settings.
    # Exact packet hashes demonstrate routing, not identity with the original MP3 bytes.
    $reference = Join-Path $project.root "validation/$CandidateName-narration-reference.m4a"
    Invoke-V2FFmpeg @('-n','-v','error','-i',$narration,'-map','0:a:0','-c:a','aac','-b:a','192k','-ar','48000','-ac','2','-t',$duration.ToString('0.000000',[Globalization.CultureInfo]::InvariantCulture),$reference)
    $actualHash=Get-V2AudioPacketHash $output; $referenceHash=Get-V2AudioPacketHash $reference
    if ($actualHash -ne $referenceHash) { throw 'Final AAC packets differ from independently encoded explicit narration.' }
    $frame = Join-Path $project.root "validation/$CandidateName-frame.png"
    Invoke-V2FFmpeg @('-n','-v','error','-i',$output,'-frames:v','1',$frame)
    $report = [ordered]@{
        candidate=$output; candidate_sha256=(Get-V2Asset $output).sha256; technical_proof_only=$true
        source_sha256=(Get-V2Asset (Join-Path $project.root $record.source.path)).sha256
        render_sha256=$render.render_sha256; narration_sha256=$record.narration.sha256; avatar_sha256=$record.avatar.sha256
        narration_path=$narration; avatar_path=$avatar; placement=$Placement; placement_pixels=$pixels
        project_default_approved=$false; provider_calls=0; render_calls=0
        ffmpeg_arguments=$arguments; narration_audio_map='2:a:0'; avatar_audio_mapped=$false
        expected_audio_packet_hash=$referenceHash; actual_audio_packet_hash=$actualHash; audio_routing_verified=$true
        narration_duration=$duration; output_duration=[double]$probe.format.duration
        decoded_alpha_has_zero_and_255=$true; output_profile_valid=$true; full_decode_ok=$true
        lip_sync_accepted=$false; visual_review='pending'; review_frame=$frame
    }
    Write-V2Json (Join-Path $project.root "validation/$CandidateName.json") $report
    [pscustomobject]$report
}

Export-ModuleMember -Function Export-V2Slide,Build-V2Slide,Get-V2CompositionArguments,Get-V2Placement,Assert-V2VideoProfile,Get-V2Probe
