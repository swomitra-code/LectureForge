param([string]$Prefix='LectureForgeF',[ValidateSet('A','C')][string]$From='A')
# Sequential, zero-provider regression runner. All evidence stays in short local roots.
$ErrorActionPreference='Stop';[Net.WebRequest]::DefaultWebProxy=$null
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'));Set-Location $repo
function Run([string]$Script,[object[]]$Arguments){
 & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot $Script) @Arguments
 if($LASTEXITCODE){throw "$Script failed"}
}
function StopFixture([int]$Port){
 $url="http://127.0.0.1:$Port/";$b=Invoke-RestMethod ($url+'api/bootstrap') -TimeoutSec 60
 $null=Invoke-RestMethod -Method Post ($url+'api/stop') -Headers @{'X-LF-Token'=$b.token} -ContentType application/json -Body '{}' -TimeoutSec 60
}
Run 'Test-NarrationCues.ps1' @()
Run 'Test-ScriptImport.ps1' @('-RuntimeRoot',(Join-Path $repo 'work/script-import-tests'))
if($From -eq 'A'){
 $a=Join-Path $env:LOCALAPPDATA ($Prefix+'A');Run 'Test-ProductionStudio.ps1' @('-RuntimeRoot',$a,'-Port',8801);StopFixture 8801
 $b=Join-Path $env:LOCALAPPDATA ($Prefix+'B');Run 'Test-NarrationStudio.ps1' @('-RuntimeRoot',$b,'-Port',8802);StopFixture 8802
}
$c=Join-Path $env:LOCALAPPDATA ($Prefix+'C');Run 'Test-SelectionAuthorization.ps1' @('-RuntimeRoot',$c,'-Port',8803);StopFixture 8803
$g=Join-Path $env:LOCALAPPDATA ($Prefix+'G');Run 'Test-AvatarGuards.ps1' @('-RuntimeRoot',$g)
# Existing completed provider fixtures; no conversion or provider retry requested.
Run 'Test-AvatarRecovery.ps1' @('-RuntimeRoot',(Join-Path $env:LOCALAPPDATA 'LectureForgeDF'),'-Id','9a6884f7a78e','-Port',8784);StopFixture 8784
$e=Join-Path $env:LOCALAPPDATA ($Prefix+'E');Run 'Test-AssemblyStudio.ps1' @('-RuntimeRoot',$e,'-Port',8804)
Run 'Test-AssemblySupplemental.ps1' @('-RuntimeRoot',$e)
& node (Join-Path $PSScriptRoot 'Test-AssemblyBrowser.cjs') $e
if($LASTEXITCODE){throw 'E browser failed'}
Run 'Test-AssemblyCheckpoint.ps1' @('-RuntimeRoot',$e,'-Port',8804);StopFixture 8804
Run 'Test-ProductionPackaging.ps1' @()
Write-Output 'PASS: sequential A-E offline regressions completed.'
