Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Project.psm1')
Import-Module (Join-Path $PSScriptRoot 'Render.psm1')
Import-Module (Join-Path $PSScriptRoot 'WhiteAvatar.psm1')

function Get-RecordingShapeInfo($Shape) {
    $item=[ordered]@{id=$Shape.Id;name=$Shape.Name;type=$Shape.Type;left=[math]::Round($Shape.Left,4);top=[math]::Round($Shape.Top,4);width=[math]::Round($Shape.Width,4);height=[math]::Round($Shape.Height,4);rotation=[math]::Round($Shape.Rotation,4);visible=$Shape.Visible;text='';children=@();table=@()}
    if($Shape.HasTextFrame -eq -1 -and $Shape.TextFrame.HasText -eq -1){$item.text=$Shape.TextFrame.TextRange.Text}
    if($Shape.Type -eq 6){for($i=1;$i -le $Shape.GroupItems.Count;$i++){$item.children+=Get-RecordingShapeInfo $Shape.GroupItems.Item($i)}}
    if($Shape.HasTable -eq -1){for($r=1;$r -le $Shape.Table.Rows.Count;$r++){for($c=1;$c -le $Shape.Table.Columns.Count;$c++){$item.table+= $Shape.Table.Cell($r,$c).Shape.TextFrame.TextRange.Text}}}
    [pscustomobject]$item
}

function Get-RecordingSlideSignature($Slide,[int]$ExcludeShapeId=-1) {
    $items=@()
    for($i=1;$i -le $Slide.Shapes.Count;$i++){
        $shape=$Slide.Shapes.Item($i)
        if($shape.Id -ne $ExcludeShapeId){$items+=Get-RecordingShapeInfo $shape}
    }
    ConvertTo-Json -InputObject $items -Depth 20 -Compress
}

function Read-RecordingPackageXml($Archive,[string]$Name) {
    $entry=$Archive.GetEntry($Name)
    if($null -eq $entry){throw "Missing package part: $Name"}
    $reader=[IO.StreamReader]::new($entry.Open())
    try{[xml]$reader.ReadToEnd()}finally{$reader.Dispose()}
}

function Test-V2RecordingPackage([string]$OutputPath,[string]$ApprovedVideoSha256,[string]$EvidenceDirectory,[int]$SlideNumber=1) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive=[IO.Compression.ZipFile]::OpenRead($OutputPath)
    try {
        $pres=Read-RecordingPackageXml $archive 'ppt/presentation.xml'
        $presRels=Read-RecordingPackageXml $archive 'ppt/_rels/presentation.xml.rels'
        $node=@($pres.presentation.sldIdLst.sldId)[$SlideNumber-1]
        $rid=$node.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $target=@($presRels.Relationships.Relationship|Where-Object Id -eq $rid)[0].Target
        $slidePart=([Uri]::new([Uri]'http://package.invalid/ppt/presentation.xml',[string]$target)).AbsolutePath.TrimStart('/')
        $slideXml=Read-RecordingPackageXml $archive $slidePart
        $rels=Read-RecordingPackageXml $archive (([IO.Path]::GetDirectoryName($slidePart).Replace('\','/'))+'/_rels/'+[IO.Path]::GetFileName($slidePart)+'.rels')
        $ns=[Xml.XmlNamespaceManager]::new($slideXml.NameTable)
        $ns.AddNamespace('p','http://schemas.openxmlformats.org/presentationml/2006/main')
        $ns.AddNamespace('a','http://schemas.openxmlformats.org/drawingml/2006/main')
        $ns.AddNamespace('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $shapeName='LectureForge_White_Avatar_Slide'+$SlideNumber.ToString('00')
        $pictures=@($slideXml.SelectNodes("//p:pic[p:nvPicPr/p:cNvPr[@name='$shapeName']]",$ns))
        if($pictures.Count -ne 1){throw 'Expected one approved media picture node.'}
        $picture=$pictures[0]
        $crop=$picture.SelectSingleNode('.//a:srcRect',$ns)
        if($crop){foreach($attribute in $crop.Attributes){if([double]$attribute.Value -ne 0){throw 'Media has nonzero crop.'}}}
        $relationshipIds=@($picture.SelectNodes('.//@r:embed | .//@r:link',$ns) | ForEach-Object Value)
        $mediaTargets=@()
        foreach($relationship in $rels.DocumentElement.ChildNodes){
            if($relationshipIds -contains $relationship.GetAttribute('Id') -and $relationship.GetAttribute('Type') -match '/(video|media)$'){
                if($relationship.GetAttribute('TargetMode') -eq 'External'){throw 'Video relationship is external.'}
                $uri=[Uri]::new([Uri]('http://package.invalid/'+$slidePart),$relationship.GetAttribute('Target'))
                if($uri.Host -ne 'package.invalid'){throw 'Unexpected package target.'}
                $mediaTargets+=[Uri]::UnescapeDataString($uri.AbsolutePath.TrimStart('/'))
            }
        }
        $mediaTargets=@($mediaTargets | Sort-Object -Unique)
        if($mediaTargets.Count -ne 1){throw 'Expected one internally referenced video asset.'}
        $entry=$archive.GetEntry($mediaTargets[0])
        if($null -eq $entry){throw 'Referenced embedded media is missing.'}
        # Keep the leaf short: authorized attempt paths approach Windows MAX_PATH.
        $resolvedEvidence=[IO.Path]::GetFullPath($EvidenceDirectory)
        $embedded=[IO.Path]::GetFullPath((Join-Path $resolvedEvidence 'embedded.mp4'))
        if([IO.Path]::GetDirectoryName($embedded) -ne $resolvedEvidence.TrimEnd('\')){throw 'Extraction destination escaped its evidence directory.'}
        if($embedded.Length -ge 260){throw "Extraction path exceeds the Windows PowerShell limit; choose a shorter evidence directory: $embedded"}
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($embedded))
        if(Test-Path -LiteralPath $embedded){throw 'Extracted evidence exists; choose a new evidence directory.'}
        $inputStream=$entry.Open();$outputStream=$null
        try{$outputStream=[IO.File]::Open($embedded,[IO.FileMode]::CreateNew);$inputStream.CopyTo($outputStream)}finally{$inputStream.Dispose();if($outputStream){$outputStream.Dispose()}}
        $hash=(Get-V2Asset $embedded).sha256
        if($hash -ne $ApprovedVideoSha256){throw 'Embedded video bytes differ from approved MP4.'}
    }finally{$archive.Dispose()}
    $embeddedProbe=Get-V2Probe $embedded
    $audio=@($embeddedProbe.streams | Where-Object codec_type -eq 'audio')
    if($audio.Count -ne 1 -or $audio[0].codec_name -ne 'aac'){throw 'Expected approved AAC stream is missing.'}
    & ffmpeg -v error -i $embedded -map '0:v:0' -map '0:a:0' -f null NUL
    if($LASTEXITCODE -ne 0){throw 'Embedded video/audio decode failed.'}
    [pscustomobject]@{no_media_crop=$true;embedded_part=$mediaTargets[0];embedded_video_sha256=$hash;audio_stream_present=$true;embedded_full_decode_ok=$true}
}

function New-V2RecordingProof {
    param(
        [Parameter(Mandatory=$true)][string]$Source,
        [Parameter(Mandatory=$true)][string]$ApprovedVideo,
        [Parameter(Mandatory=$true)][string]$ApprovedVideoSha256,
        [Parameter(Mandatory=$true)][string]$Output,
        [string]$EvidenceDirectory=(Join-Path $PSScriptRoot '../work/v2-phase2b/recording'),
        [Parameter(Mandatory=$true)][object]$Placement
    )
    $sourceAsset=Get-V2Asset $Source
    $videoAsset=Get-V2Asset $ApprovedVideo
    if($videoAsset.sha256 -ne $ApprovedVideoSha256){throw 'Approved video hash mismatch.'}
    $info=Get-V2DeckInfo $sourceAsset.path
    if($info.slide_count -ne 14 -or -not $info.is_16_9){throw 'Expected the 14-slide 16:9 source.'}
    $probe=Get-V2Probe $videoAsset.path
    Assert-V2WhiteAvatarProfile $probe
    foreach($key in 'left','top','width','height'){
        $value=[double]$Placement.$key
        if([double]::IsNaN($value) -or [double]::IsInfinity($value)){throw 'Non-finite avatar placement.'}
    }
    $slideWidth=$info.width_emu/12700.0;$slideHeight=$info.height_emu/12700.0
    if($Placement.left -lt 0 -or $Placement.top -lt 0 -or $Placement.width -le 0 -or $Placement.height -le 0 -or $Placement.left+$Placement.width -gt $slideWidth -or $Placement.top+$Placement.height -gt $slideHeight -or $Placement.width -gt $slideWidth*.3 -or $Placement.height -gt $slideHeight*.3){throw 'Avatar must be compact and inside the reserved slide area.'}
    if([math]::Abs($Placement.width/$Placement.height-640.0/560.0) -gt .001){throw 'Avatar aspect ratio must be preserved.'}
    $outputPath=[IO.Path]::GetFullPath($Output)
    if($outputPath -eq $sourceAsset.path -or $outputPath -eq $videoAsset.path -or (Test-Path -LiteralPath $outputPath)){throw 'Recording output must be a distinct new file.'}
    if([IO.Path]::GetFileName($outputPath) -notmatch '_RECORDING_AVATAR_WHITE\.pptx$'){throw 'Use a clearly distinguished _RECORDING_AVATAR_WHITE.pptx output.'}
    $evidence=[IO.Path]::GetFullPath($EvidenceDirectory)
    if(Test-Path -LiteralPath $evidence){throw 'Evidence directory exists; preserve prior proof and choose a new directory.'}
    New-Item -ItemType Directory -Path $evidence | Out-Null
    New-Item -ItemType Directory -Path (Split-Path -Parent $outputPath) -Force | Out-Null
    # A reference snapshot avoids opening/closing any user-owned canonical presentation.
    $referencePath=Join-Path $evidence 'source-reference.pptx'
    Copy-Item -LiteralPath $sourceAsset.path -Destination $referencePath
    Copy-Item -LiteralPath $sourceAsset.path -Destination $outputPath
    if((Get-V2Asset $referencePath).sha256 -ne $sourceAsset.sha256 -or (Get-V2Asset $outputPath).sha256 -ne $sourceAsset.sha256){throw 'Initial copies are not byte-identical.'}
    $report=[ordered]@{
        source_path=$sourceAsset.path;source_sha256_before=$sourceAsset.sha256;source_sha256_after=$null
        approved_video_path=$videoAsset.path;approved_video_sha256=$videoAsset.sha256
        approval_scope='Instructor-approved white-avatar architecture, motion and technical placement; real narration audibility pending'
        final_narration_approved=$false;lip_sync_approved=$false;output_path=$outputPath
        source_slide_count=14;output_slide_count=0;modified_slides=@(1);unrelated_slides=@()
        reopened_successfully=$false;media_controls_setting_supported=$false;media_controls_hidden=$null
        provider_calls=0;canonical_save_calls=0;manual_cursor_laser_ink_acceptance='pending instructor test'
        automated_checks_passed=$false
    }
    $signatures=@{};$renderHashes=@{}
    $wasRunning=@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
    $app=$null;$reference=$null;$deck=$null;$show=$null
    try {
        $app=New-Object -ComObject PowerPoint.Application
        $report.powerpoint_version=[string]$app.Version
        $reference=$app.Presentations.Open($referencePath,$true,$false,$false)
        if($reference.Slides.Count -ne 14){throw 'Source reference count mismatch.'}
        for($n=1;$n -le 14;$n++){
            $slide=$reference.Slides.Item($n)
            $signatures[$n]=Get-RecordingSlideSignature $slide
            if($n -gt 1){
                $png=Join-Path $evidence "source-slide-$n.png"
                $slide.Export($png,'PNG',1920,1080)
                $renderHashes[$n]=(Get-V2Asset $png).sha256
            }elseif(@($slide.Shapes | Where-Object Type -eq 16).Count -ne 0){throw 'Source Slide 1 already has media; refusing ambiguous audio playback.'}
        }
        $reference.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reference);$reference=$null
        Write-Output 'Source reference rendered and structurally recorded (14 slides).'

        # The only editable presentation opened by this function is the new derivative.
        $deck=$app.Presentations.Open($outputPath,$false,$false,$false)
        $slide=$deck.Slides.Item(1)
        $media=$slide.Shapes.AddMediaObject2($videoAsset.path,$false,$true,$Placement.left,$Placement.top,$Placement.width,$Placement.height)
        $media.Name='LectureForge_White_Avatar_Slide01'
        $media.LockAspectRatio=0
        $media.Left=$Placement.left;$media.Top=$Placement.top
        $media.Width=$Placement.width;$media.Height=$Placement.height
        $media.Line.Visible=0;$media.Shadow.Visible=0
        $media.AnimationSettings.Animate=-1
        $media.AnimationSettings.PlaySettings.PlayOnEntry=-1
        $mediaId=[int]$media.Id
        for($i=1;$i -le $slide.TimeLine.MainSequence.Count;$i++){
            $effect=$slide.TimeLine.MainSequence.Item($i)
            if($effect.Shape.Id -eq $mediaId -and $effect.EffectType -eq 83){$effect.Timing.TriggerType=2}
        }
        try{$deck.SlideShowSettings.ShowMediaControls=0;$report.media_controls_setting_supported=$true}catch{$report.media_controls_setting_note='Installed PowerPoint did not expose ShowMediaControls; instructor must inspect controls.'}
        $deck.Save()
        $deck.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($deck);$deck=$null

        $deck=$app.Presentations.Open($outputPath,$true,$false,$false)
        $report.reopened_successfully=$true
        $report.output_slide_count=$deck.Slides.Count
        if($deck.Slides.Count -ne 14){throw 'Derivative slide count changed.'}
        $slide=$deck.Slides.Item(1)
        $mediaShapes=@($slide.Shapes | Where-Object Type -eq 16)
        if($mediaShapes.Count -ne 1){throw 'Expected one embedded media shape on Slide 1.'}
        $saved=$mediaShapes[0]
        if($saved.Name -ne 'LectureForge_White_Avatar_Slide01'){throw 'Unexpected media shape binding.'}
        $mediaId=[int]$saved.Id
        $linked=$false
        try{$linked=-not [string]::IsNullOrWhiteSpace([string]$saved.LinkFormat.SourceFullName)}catch{}
        $report.media_is_linked=$linked
        $report.avatar_geometry_matches=([math]::Abs($saved.Left-$Placement.left) -lt 0.001 -and [math]::Abs($saved.Top-$Placement.top) -lt 0.001 -and [math]::Abs($saved.Width-$Placement.width) -lt 0.001 -and [math]::Abs($saved.Height-$Placement.height) -lt 0.001)
        $report.media_geometry=@{left=$saved.Left;top=$saved.Top;width=$saved.Width;height=$saved.Height}
        $report.border_and_shadow_hidden=($saved.Line.Visible -eq 0 -and $saved.Shadow.Visible -eq 0)
        $report.play_on_entry=($saved.AnimationSettings.PlaySettings.PlayOnEntry -eq -1)
        $report.autoplay_configured=$false
        for($i=1;$i -le $slide.TimeLine.MainSequence.Count;$i++){
            $effect=$slide.TimeLine.MainSequence.Item($i)
            if($effect.Shape.Id -eq $mediaId -and $effect.EffectType -eq 83 -and $effect.Timing.TriggerType -eq 2){$report.autoplay_configured=$true}
        }
        $report.slide1_native_content_preserved=((Get-RecordingSlideSignature $slide $mediaId) -eq $signatures[1])
        if($report.media_controls_setting_supported){$report.media_controls_hidden=($deck.SlideShowSettings.ShowMediaControls -eq 0)}
        if($linked -or -not $report.avatar_geometry_matches -or -not $report.border_and_shadow_hidden -or -not $report.play_on_entry -or -not $report.autoplay_configured -or -not $report.slide1_native_content_preserved){throw 'Reopened Slide 1 structural checks failed.'}
        if($report.media_controls_setting_supported -and -not $report.media_controls_hidden){throw 'Media controls setting did not persist.'}
        # Change only the reopened validation copy in memory; never save these probes.
        $saved.Left=[single]($Placement.left+1);$saved.Width=[single]($Placement.width*.95)
        $report.independent_move_resize=([math]::Abs($saved.Left-($Placement.left+1)) -lt .001 -and [math]::Abs($saved.Width-$Placement.width*.95) -lt .001)
        $saved.Left=[single]$Placement.left;$saved.Width=[single]$Placement.width;$saved.Height=[single]$Placement.height
        if(-not $report.independent_move_resize){throw 'Independent avatar move/resize failed.'}
        $deck.Saved=-1
        $slide.Export((Join-Path $evidence 'recording-slide-1.png'),'PNG',1920,1080)
        for($n=2;$n -le 14;$n++){
            $slide=$deck.Slides.Item($n)
            $png=Join-Path $evidence "recording-slide-$n.png"
            $slide.Export($png,'PNG',1920,1080)
            $structural=((Get-RecordingSlideSignature $slide) -eq $signatures[$n])
            $hash=(Get-V2Asset $png).sha256
            $visual=($hash -eq $renderHashes[$n])
            $report.unrelated_slides+=@{slide=$n;structure_unchanged=$structural;render_identical=$visual;source_render_sha256=$renderHashes[$n];output_render_sha256=$hash}
            if(-not $structural -or -not $visual){throw "Unrelated Slide $n changed; stopping."}
        }
        Write-Output 'Reopened derivative: independent compact autoplay avatar verified; Slides 2-14 match structurally and visually.'

        # Optional live autoplay observation. Do not call Player.Play(), which could
        # conceal a broken automatic trigger. Manual recording acceptance is separate.
        $report.live_autoplay_advanced=$false
        try {
            $settings=$deck.SlideShowSettings
            $settings.RangeType=2;$settings.StartingSlide=1;$settings.EndingSlide=1;$settings.ShowType=2
            $show=$settings.Run()
            Start-Sleep -Milliseconds 750
            $player=$show.View.Player($mediaId)
            $before=[double]$player.CurrentPosition
            Start-Sleep -Seconds 2
            $after=[double]$player.CurrentPosition
            $report.live_autoplay_advanced=($after -gt $before)
            $report.live_position_before=$before;$report.live_position_after=$after
        }catch{$report.live_autoplay_note=$_.Exception.Message}
        finally{if($show){try{$show.View.Exit()}catch{};$show=$null}}
        # Exiting a windowed show can already close its presentation/application.
        # Cleanup is idempotent; all persisted structure is checked before the probe,
        # and independent package checks below do not rely on this COM object.
        try{$deck.Close()}catch{$report.post_show_cleanup_note=$_.Exception.Message}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($deck);$deck=$null

        $package = Test-V2RecordingPackage $outputPath $videoAsset.sha256 $evidence
        foreach($property in $package.PSObject.Properties){$report[$property.Name]=$property.Value}
        $report.embedded_motion=Get-V2MotionEvidence (Join-Path $evidence 'embedded.mp4') (Join-Path $evidence 'embedded-mouth.framemd5')
        if(-not $report.live_autoplay_advanced){throw 'Live autoplay did not advance; stop for desktop inspection.'}
        $report.source_sha256_after=(Get-V2Asset $sourceAsset.path).sha256
        if($report.source_sha256_after -ne $sourceAsset.sha256 -or (Get-V2Asset $videoAsset.path).sha256 -ne $videoAsset.sha256){throw 'Canonical source or approved MP4 changed.'}
        $report.output_sha256=(Get-V2Asset $outputPath).sha256
        $report.automated_checks_passed=$true
    }catch{$report.error=$_.Exception.Message;throw}
    finally {
        if($deck){try{$deck.Close()}catch{}}
        if($reference){try{$reference.Close()}catch{}}
        if($app -and -not $wasRunning){try{$app.Quit()}catch{}}
        foreach($com in @($deck,$reference,$app)){if($null -ne $com){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com)}}
        $report.source_sha256_after=(Get-V2Asset $sourceAsset.path).sha256
        if($report.source_sha256_after -ne $sourceAsset.sha256){$report.automated_checks_passed=$false;$report.error='CANONICAL SOURCE HASH CHANGED'}
        Write-V2Json (Join-Path $evidence 'recording-validation.json') $report
        [GC]::Collect();[GC]::WaitForPendingFinalizers()
    }
    [pscustomobject]$report
}

Export-ModuleMember -Function New-V2RecordingProof,Test-V2RecordingPackage
function New-V2RecordingBatch {
 param([string]$Source,[string]$SourceSha256,[string]$Output,[string]$EvidenceDirectory,[object[]]$Avatars,[scriptblock]$Progress)
 $ErrorActionPreference='Stop'
 $sourceAsset=Get-V2Asset $Source
 if($sourceAsset.sha256 -ne $SourceSha256){throw 'Source PPTX hash mismatch.'}
 $outputPath=[IO.Path]::GetFullPath($Output);$ev=[IO.Path]::GetFullPath($EvidenceDirectory)
 if($outputPath -eq $sourceAsset.path -or (Test-Path $outputPath) -or $outputPath.Length -ge 240 -or $ev.Length -ge 210){throw 'Use a distinct new derivative and short evidence paths.'}
 [void][IO.Directory]::CreateDirectory($ev);[void][IO.Directory]::CreateDirectory((Split-Path $outputPath))
 $info=Get-V2DeckInfo $Source;$sw=$info.width_emu/12700.0;$sh=$info.height_emu/12700.0
 $map=@{};foreach($a in $Avatars){if($map.ContainsKey([int]$a.slide) -or $a.slide -lt 1 -or $a.slide -gt $info.slide_count){throw 'Duplicate or invalid avatar slide.'};if((Get-V2Asset $a.path).sha256 -ne $a.sha256){throw "Slide $($a.slide) avatar hash mismatch."};$map[[int]$a.slide]=$a}
 [IO.File]::Copy($sourceAsset.path,$outputPath,$false)
 if((Get-V2Asset $outputPath).sha256 -ne $SourceSha256){throw 'Derivative copy hash mismatch.'}
 $app=$null;$deck=$null;$wasRunning=@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
 $report=[ordered]@{source_sha256_before=$SourceSha256;source_sha256_after=$null;slide_count=$info.slide_count;avatar_count=$Avatars.Count;slides=@();save_calls=0;reopened=$false;passed=$false;provider_calls=0}
 $signatures=@{};$renders=@{};$stage='Opening PowerPoint';$number=0
 try{
  & $Progress $stage 0
  $app=New-Object -ComObject PowerPoint.Application
  $deck=$app.Presentations.Open($outputPath,$false,$false,$false)
  for($n=1;$n -le $deck.Slides.Count;$n++){
   $number=$n;$slide=$deck.Slides.Item($n);$signatures[$n]=Get-RecordingSlideSignature $slide
   if($map.ContainsKey($n) -and @($slide.Shapes|Where-Object Type -eq 16).Count){throw "Slide $n source already contains media; ambiguous playback."}
   $png=Join-Path $ev "b$n.png";$slide.Export($png,'PNG',1920,1080);$renders[$n]=(Get-V2Asset $png).sha256
  }
  $done=0
  foreach($a in $Avatars){
   $number=[int]$a.slide;$stage='Inserting avatars';$done++;& $Progress $stage $number
   $p=$a.placement;$slide=$deck.Slides.Item($number)
   $m=$slide.Shapes.AddMediaObject2($a.path,$false,$true,($p.left*$sw),($p.top*$sh),($p.width*$sw),($p.height*$sh))
   $m.Name='LectureForge_White_Avatar_Slide'+$number.ToString('00');$m.LockAspectRatio=0
   $m.Left=$p.left*$sw;$m.Top=$p.top*$sh;$m.Width=$p.width*$sw;$m.Height=$p.height*$sh
   $m.Line.Visible=0;$m.Shadow.Visible=0;$m.AnimationSettings.Animate=-1;$m.AnimationSettings.PlaySettings.PlayOnEntry=-1
   for($i=1;$i -le $slide.TimeLine.MainSequence.Count;$i++){$effect=$slide.TimeLine.MainSequence.Item($i);if($effect.Shape.Id -eq $m.Id -and $effect.EffectType -eq 83){$effect.Timing.TriggerType=2}}
  }
  $number=0;$stage='Saving';& $Progress $stage 0
  $deck.Save();$report.save_calls++;$deck.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($deck);$deck=$null
  $stage='Validating';& $Progress $stage 0
  $deck=$app.Presentations.Open($outputPath,$true,$false,$false);$report.reopened=$true
  if($deck.Slides.Count -ne $info.slide_count){throw 'Slide count changed.'}
  for($n=1;$n -le $deck.Slides.Count;$n++){
   $number=$n;& $Progress $stage $n;$slide=$deck.Slides.Item($n);$exclude=-1
   $result=[ordered]@{slide=$n;produced=$map.ContainsKey($n);native_preserved=$false;native_render_identical=$false}
   if($map.ContainsKey($n)){
    $a=$map[$n];$p=$a.placement;$media=@($slide.Shapes|Where-Object Type -eq 16)
    if($media.Count -ne 1 -or $media[0].Name -ne ('LectureForge_White_Avatar_Slide'+$n.ToString('00'))){throw "Slide $n expected exactly one named media object."}
    $m=$media[0];$exclude=[int]$m.Id
    foreach($key in 'left','top','width','height'){$scale=if($key -in 'left','width'){$sw}else{$sh};if([math]::Abs($m.$key-$p.$key*$scale) -gt .002){throw "Slide $n placement mismatch: $key"}}
    $auto=$false;for($i=1;$i -le $slide.TimeLine.MainSequence.Count;$i++){$effect=$slide.TimeLine.MainSequence.Item($i);if($effect.Shape.Id -eq $m.Id -and $effect.EffectType -eq 83 -and $effect.Timing.TriggerType -eq 2){$auto=$true}}
    if(-not $auto -or $m.AnimationSettings.PlaySettings.PlayOnEntry -ne -1 -or $m.Line.Visible -ne 0 -or $m.Shadow.Visible -ne 0){throw "Slide $n autoplay/border/shadow mismatch."}
    $left=[single]$m.Left;$width=[single]$m.Width;$m.Left=[single]($left+1);$m.Width=[single]($width*.95)
    if([math]::Abs($m.Left-$left-1) -gt .002 -or [math]::Abs($m.Width-$width*.95) -gt .002){throw "Slide $n independent move/resize failed."}
    $m.Left=[single]$left;$m.Width=[single]$width;$m.Visible=0
    $result.independent_move_resize=$true;$result.autoplay=$true;$result.placement=$p
   }
   $result.native_preserved=((Get-RecordingSlideSignature $slide $exclude) -eq $signatures[$n])
   $png=Join-Path $ev "c$n.png";$slide.Export($png,'PNG',1920,1080);$result.native_render_identical=((Get-V2Asset $png).sha256 -eq $renders[$n])
   if($map.ContainsKey($n)){$m.Visible=-1}
   if(-not $result.native_preserved -or -not $result.native_render_identical){throw "Slide $n native content changed."}
   $report.slides+=,$result
  }
  # The movement/visibility probes are in-memory only; never save validation changes.
  $deck.Saved=-1;$deck.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($deck);$deck=$null
  foreach($a in $Avatars){$number=[int]$a.slide;& $Progress 'Validating embedded media' $number
   $package=Test-V2RecordingPackage $outputPath $a.sha256 (Join-Path $ev "s$number") $number
   $row=$report.slides[$number-1];$row.package=$package;$row.narration_sha256=$a.narration.sha256;$row.preparation=$a.preparation
  }
  $report.source_sha256_after=(Get-V2Asset $Source).sha256
  if($report.source_sha256_after -ne $SourceSha256){throw 'CANONICAL SOURCE HASH CHANGED'}
  $report.output_sha256=(Get-V2Asset $outputPath).sha256;$report.passed=$true
 }catch{$report.error="Stage $stage, Slide $number : $($_.Exception.Message)";$report.error_location=$_.ScriptStackTrace;throw $report.error}
 finally{
  if($deck){try{$deck.Saved=-1;$deck.Close()}catch{}}
  if($app -and -not $wasRunning){try{if($app.Presentations.Count -eq 0){$app.Quit()}}catch{}}
  foreach($com in @($deck,$app)){if($null -ne $com){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com)}}
  $report.source_sha256_after=(Get-V2Asset $Source).sha256;Write-V2Json (Join-Path $ev 'validation.json') $report
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
 }
 [pscustomobject]$report
}
Export-ModuleMember -Function New-V2RecordingBatch
