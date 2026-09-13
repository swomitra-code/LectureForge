param([string]$RuntimeRoot=(Join-Path $env:LOCALAPPDATA 'LectureForge'),[int]$Port=8770)
$ErrorActionPreference='Stop'
[Net.WebRequest]::DefaultWebProxy=$null
[Net.ServicePointManager]::Expect100Continue=$false
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Import-Module (Join-Path $repo 'studio/Production.psm1') -Force
$root=Get-LFRoot $RuntimeRoot
$test=Join-Path $root ('Tests/'+[guid]::NewGuid().ToString('N').Substring(0,8));[void][IO.Directory]::CreateDirectory($test)
$checks=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){$checks.Add(@{name=$Name;passed=$Passed});if(-not $Passed){throw "FAILED: $Name"};Write-Output "PASS: $Name"}
$source=Join-Path $repo 'source/RE-M2-Solar-Thermal-Overview.pptx';$before=(Get-FileHash $source).Hash
$copy=Join-Path $test 'input.pptx';Copy-Item -LiteralPath $source -Destination $copy
Check 'Acceptance input is an exact COPY of Solar Thermal' ((Get-FileHash $copy).Hash -eq $before)
$existing=Join-Path $root 'runtime.json'
if(Test-Path $existing){
 $prior=Get-Content -Raw $existing|ConvertFrom-Json
 if($prior.status -eq 'running' -and $prior.milestone -eq 'A'){
  [IO.File]::WriteAllText((Join-Path $root ($prior.instance+'.stop')),'stop')
  Start-Sleep -Seconds 3
 }
}
& (Join-Path $repo 'Start-LectureForge.ps1') -RuntimeRoot $root -Port $Port -NoBrowser
if($LASTEXITCODE -and $LASTEXITCODE -ne 0){throw 'Launcher failed'}
$url="http://127.0.0.1:$Port/"
function GetJson([string]$Path){Invoke-RestMethod -UseBasicParsing -Uri ($url+$Path) -TimeoutSec 10}
function PostJson([string]$Path,$Body){Invoke-RestMethod -UseBasicParsing -Method Post -Uri ($url+$Path) -Headers @{'X-LF-Token'=$script:token} -ContentType 'application/json; charset=utf-8' -Body ([Text.Encoding]::UTF8.GetBytes(($Body|ConvertTo-Json -Depth 20))) -TimeoutSec 20}
$boot=GetJson 'api/bootstrap';$script:token=$boot.token;$runtime=Get-Content -Raw (Join-Path $root 'runtime.json')|ConvertFrom-Json
Check 'Server and idle worker ready without production actions' ((GetJson 'api/health').ready)
& (Join-Path $repo 'Start-LectureForge.ps1') -RuntimeRoot $root -Port $Port -NoBrowser
Check 'Second launcher reuses existing instance' ((GetJson 'api/health').instance -eq $runtime.instance)
$name='Solar Thermal - Milestone A '+[char]0xE9
$p=Invoke-RestMethod -UseBasicParsing -Method Post -Uri ($url+'api/new?name='+[Uri]::EscapeDataString($name)+'&filename=Solar-Thermal-copy.pptx&preset=renewable-energy') -Headers @{'X-LF-Token'=$token} -ContentType 'application/octet-stream' -Body ([IO.File]::ReadAllBytes($copy)) -TimeoutSec 60
$id=$p.id
Check 'Project source, slide count, title and Unicode name inspected' ($p.source.sha256 -eq $before -and $p.slides.Count -eq 14 -and $p.slides[0].title.Length -gt 0 -and $p.name -eq $name)
Check 'Short root and project-relative source path' ($p.source.path -eq 'source.pptx' -and (Get-LFProjectPath $root $id).Length -lt 150)
$approved=Get-Content -Encoding UTF8 -Raw (Join-Path $repo 'work/solar-thermal-slides2-14-production-r001/scripts.json')|ConvertFrom-Json
$text='';for($n=2;$n -le 14;$n++){$text+="## Slide $n`n"+$approved."$n"+"`n`n"}
$p=PostJson "api/save?id=$id" @{version=$p.version;script_import=$text}
Check '13 scripts imported with exact body binding' (@($p.slides|Where-Object {$_.number -gt 1 -and $_.script -ceq $approved."$($_.number)"}).Count -eq 13)
$oldVersion=$p.version;$p.slides[0].enabled=$false;$p.slides[7].post_speech_silence_seconds=6
$p=PostJson "api/save?id=$id" @{version=$p.version;slides=$p.slides}
Check 'Enabled slides and pause metadata saved without media generation' (-not $p.slides[0].enabled -and $p.slides[7].post_speech_silence_seconds -eq 6 -and $p.provider_calls -eq 0)
$rejected=$false;try{PostJson "api/save?id=$id" @{version=$oldVersion;slides=$p.slides}|Out-Null}catch{$rejected=$true};Check 'Stale project revision rejected' $rejected
$rejected=$false;try{Invoke-RestMethod -UseBasicParsing -Method Post -Uri ($url+'api/stop') -ContentType 'application/json' -Body '{}'|Out-Null}catch{$rejected=$true};Check 'Mutation without session token rejected' $rejected
$rejected=$false;try{PostJson 'api/generate' @{}|Out-Null}catch{$rejected=$true};Check 'Provider generation endpoint absent' $rejected
$rejected=$false;try{ConvertFrom-LFScripts "## Slide 2`na`n## Slide 2`nb`n" 14|Out-Null}catch{$rejected=$true};Check 'Duplicate script headings rejected' $rejected
$saved=(Get-FileHash (Join-Path (Get-LFProjectPath $root $id) 'project.json')).Hash
$null=PostJson 'api/stop' @{}
for($i=0;$i -lt 40;$i++){Start-Sleep -Milliseconds 250;$r=Read-LFSharedJson (Join-Path $root 'runtime.json');if($r.status -eq 'stopped'){break}}
Check 'Clean stop persisted and worker stopped' ($r.status -eq 'stopped' -and (Get-Content -Raw (Join-Path $root ($runtime.instance+'.worker.json'))|ConvertFrom-Json).stage -eq 'stopped')
& (Join-Path $repo 'Start-LectureForge.ps1') -RuntimeRoot $root -Port $Port -NoBrowser
$script:token=(GetJson 'api/bootstrap').token;$restored=GetJson "api/project?id=$id"
Check 'Restart restores identical project state and source hash' ((Get-FileHash (Join-Path (Get-LFProjectPath $root $id) 'project.json')).Hash -eq $saved -and $restored.source.sha256 -eq $before)
Check 'Canonical Solar Thermal deck unchanged' ((Get-FileHash $source).Hash -eq $before)
Check 'No generated takes, avatars or provider jobs' (@(Get-ChildItem (Get-LFProjectPath $root $id) -File).Count -eq 2)
$result=@{passed=$true;checks=@($checks);project_id=$id;project_root=Get-LFProjectPath $root $id;runtime_root=$root;source_sha256=$before;provider_calls=0;url=$url;test_directory=$test}
Write-LFJson (Join-Path $test 'results.json') $result
Write-Output ('RESULTS='+ (Join-Path $test 'results.json'))
Write-Output ('PROJECT='+ (Get-LFProjectPath $root $id))
