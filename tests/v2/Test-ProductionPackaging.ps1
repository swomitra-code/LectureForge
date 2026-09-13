$ErrorActionPreference='Stop'
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Set-Location $repo
$files=@(Get-ChildItem studio,tests/v2 -Recurse -File | Where-Object Extension -in '.ps1','.psm1','.js','.cjs','.md','.json','.html','.css')
$files+=Get-Item Start-LectureForge.ps1,LectureForge.cmd,LECTUREFORGE_USER_GUIDE.md,LECTUREFORGE_DEVELOPER_GUIDE.md
$ps=0;$js=0
foreach($file in $files){
 $body=[IO.File]::ReadAllText($file.FullName)
 foreach($key in @($env:ELEVENLABS_API_KEY,$env:HEYGEN_API_KEY)){if($key -and $body.Contains($key)){throw "Credential literal in $($file.Name)"}}
 if($file.Extension -in '.ps1','.psm1'){
  $errors=$null;$tokens=$null;$null=[Management.Automation.Language.Parser]::ParseInput($body,[ref]$tokens,[ref]$errors)
  if($errors){throw "$($file.Name): $($errors.Message -join '; ')"};$ps++
 }
 if($file.Extension -in '.js','.cjs'){node --check $file.FullName;if($LASTEXITCODE){throw "JS syntax: $($file.Name)"};$js++}
}
$launcher=[IO.File]::ReadAllText((Join-Path $repo 'LectureForge.cmd'))
if(-not $launcher.Contains('"%~dp0Start-LectureForge.ps1"') -or $launcher -match '(?i)encodedcommand|windowstyle hidden'){throw 'Unsafe launcher'}
git diff --check
if($LASTEXITCODE){throw 'Whitespace errors'}
Write-Output "PASS: $ps PowerShell files, $js JavaScript files, $($files.Count) credential-scanned text files, launcher and Git whitespace."
