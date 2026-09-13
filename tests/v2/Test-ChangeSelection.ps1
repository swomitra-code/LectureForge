param(
 [ValidateSet('Before','Cleared','Selected','Resumed')][string]$Phase,
 [string]$RuntimeRoot=(Join-Path $env:LOCALAPPDATA 'LectureForgeB'),
 [string]$ProjectId='31cd01c7e329'
)
# Observe the real UI test; this script never changes selections or calls providers.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '../../studio/Narration.psm1') -Force
$productionRoot=Join-Path $env:LOCALAPPDATA 'LectureForge'
if((Get-LFRoot '') -ne $productionRoot){throw 'Backend production runtime root drift.'}
$tokens=$null;$errors=$null
$launcher=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '../../Start-LectureForge.ps1'),[ref]$tokens,[ref]$errors)
$default=($launcher.ParamBlock.Parameters|Where-Object {$_.Name.VariablePath.UserPath -eq 'RuntimeRoot'}).DefaultValue.Extent.Text
if($errors.Count -or $default -ne "(Join-Path `$env:LOCALAPPDATA 'LectureForge')"){throw 'Launcher production runtime root drift.'}
$dir=Get-LFProjectPath $RuntimeRoot $ProjectId
$file=Join-Path $RuntimeRoot 'change-selection-results.json'
$state=Read-LFNarrationFile $dir;$r=@($state.revisions|Where-Object slide_number -eq 8)[-1]
$takes=$r.takes|ConvertTo-Json -Depth 20 -Compress
foreach($t in $r.takes){if(-not (Test-LFTake $dir $t)){throw 'Existing take missing or hash changed.'}}
if(@($r.takes).Count -ne 3){throw 'Expected three existing takes.'}
$calls=@($state.revisions.takes.attempts|Where-Object submitted_utc).Count
if($Phase -eq 'Before'){
 if($r.selected_take -ne 2 -or $calls -ne 3){throw 'Start with Take 2 and the three original live-proof attempts.'}
 if(Test-Path $file){throw 'Evidence exists; preserve it and use a separate test run.'}
 $result=@{project_id=$ProjectId;initial_calls=$calls;takes_sha256=Get-LFTextHash $takes;checks=@('Before');elevenlabs_calls_during_checks=0;heygen_calls=0}
}else{
 $result=Get-Content -Raw $file|ConvertFrom-Json
 if($result.project_id -ne $ProjectId -or (Get-LFTextHash $takes) -ne $result.takes_sha256 -or $calls -ne $result.initial_calls){throw 'Take metadata or provider attempts changed.'}
 $expected=if($Phase -eq 'Cleared'){$null}else{3}
 if($r.selected_take -ne $expected){throw "Unexpected selection during $Phase."}
 if($result.checks -notcontains $Phase){$result.checks+= $Phase}
}
Write-LFJson $file $result
Write-Output "PASS: $Phase - expected selection, three intact takes, unchanged provider attempts."
