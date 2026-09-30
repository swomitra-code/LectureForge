param([string]$EvidenceDirectory=(Join-Path $PSScriptRoot ('../../work/avatar-alpha-'+[guid]::NewGuid().ToString('N').Substring(0,8))))
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Import-Module (Join-Path $repo 'studio/Render.psm1') -Force
$dir=[IO.Path]::GetFullPath($EvidenceDirectory)
if(Test-Path -LiteralPath $dir){throw 'Use a fresh evidence directory.'}
[void][IO.Directory]::CreateDirectory($dir)
function FF([string[]]$Arguments){& ffmpeg @Arguments;if($LASTEXITCODE){throw 'FFmpeg regression command failed.'}}
function Check([string]$Name,[bool]$Passed){if(-not $Passed){throw "FAIL: $Name"};Write-Output "PASS: $Name"}
# Read the actual production filter literal, avoiding a duplicated test implementation.
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'studio/WhiteAvatar.psm1'),[ref]$tokens,[ref]$errors)
Check 'White-avatar module parses' ($errors.Count -eq 0)
$assignment=@($ast.FindAll({param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$filter'},$true))
Check 'Exactly one production white-avatar filter' ($assignment.Count -eq 1)
$literal=$assignment[0].Right.Find({param($n) $n -is [Management.Automation.Language.StringConstantExpressionAst]},$true)
$filter=$literal.Value
# Three black strips extend through the cropped bottom: opaque, half-alpha, transparent.
# Hidden RGB is black too, so losing transparency cannot pass as a white background.
$rgba=New-Object byte[] (1920*1080*4)
for($y=70;$y -lt 630;$y++){
 for($x=1030;$x -lt 1190;$x++){$rgba[($y*1920+$x)*4+3]=255}
 for($x=1190;$x -lt 1350;$x++){$rgba[($y*1920+$x)*4+3]=128}
}
$raw=Join-Path $dir 'source.rgba';[IO.File]::WriteAllBytes($raw,$rgba)
$avatar=Join-Path $dir 'source.webm'
FF @('-v','error','-f','rawvideo','-pixel_format','rgba','-video_size','1920x1080','-framerate','25','-i',$raw,'-frames:v','1','-c:v','libvpx-vp9','-lossless','1','-pix_fmt','yuva420p',$avatar)
$slide=Join-Path $dir 'slide.png';$audio=Join-Path $dir 'audio.wav'
FF @('-v','error','-f','lavfi','-i','color=white:s=1920x1080','-frames:v','1',$slide)
FF @('-v','error','-f','lavfi','-i','anullsrc=r=48000:cl=stereo','-t','0.08',$audio)
$tile=Join-Path $dir 'tile.mp4'
FF @('-v','error','-c:v','libvpx-vp9','-i',$avatar,'-filter_complex_threads','1','-filter_complex',$filter,'-map','[v]','-frames:v','1','-c:v','libx264','-crf','18',$tile)
$full=Join-Path $dir 'full.mp4'
$crop=@{x=870;y=70;width=640;height=560}
# Native-size placement isolates alpha behavior from scaling interpolation.
$placement=@{x=0.5;y=0.25;width=(640.0/1920)}
$argsFull=@(Get-V2CompositionArguments $slide $avatar $audio $full $crop $placement .08)
FF $argsFull
foreach($case in @(@{name='Shared white tile';path=$tile;width=640;x=0;y=0},@{name='Full-slide renderer';path=$full;width=1920;x=960;y=270})){
 $pixels=Join-Path $dir ($case.width.ToString()+'.rgb')
 FF @('-v','error','-i',$case.path,'-frames:v','1','-pix_fmt','rgb24','-f','rawvideo',$pixels)
 $data=[IO.File]::ReadAllBytes($pixels)
 foreach($row in 480,500,530,558,559){
  foreach($sample in @(@{x=80;value=255;name='transparent background'},@{x=240;value=0;name='opaque torso'},@{x=400;value=127;name='natural partial alpha'})){
   $offset=(($case.y+$row)*$case.width+$case.x+$sample.x)*3
   $ok=$true;foreach($channel in 0,1,2){if([math]::Abs([int]$data[$offset+$channel]-$sample.value) -gt 6){$ok=$false}}
   Check ($case.name+': '+$sample.name+' preserved at row '+$row) $ok
  }
 }
}
Write-Output "PASS: bottom-edge alpha regression. Evidence: $dir"
