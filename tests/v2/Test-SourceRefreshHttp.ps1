param([Parameter(Mandatory=$true)][string]$FixtureRoot,[int]$Port=8897)
$ErrorActionPreference='Stop'
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Import-Module "$repo/studio/Production.psm1"
Add-Type -AssemblyName System.IO.Compression.FileSystem
$root=Join-Path $FixtureRoot ('http-'+[guid]::NewGuid().ToString('N').Substring(0,6));[void][IO.Directory]::CreateDirectory($root)
$external=Join-Path $root 'external.pptx';[IO.File]::Copy((Join-Path $FixtureRoot 'source with spaces.pptx'),$external)
[void][IO.Directory]::CreateDirectory("$root/SourceChoices");$choice='a'*32
Write-LFJson "$root/SourceChoices/$choice.json" @{state='selected';path=$external}
Write-LFJson "$root/avatar-policy.json" @{provider_calls_enabled=$false}
$server=Start-Process (Join-Path $PSHOME 'powershell.exe') -WindowStyle Hidden -PassThru -RedirectStandardOutput "$root/server.log" -RedirectStandardError "$root/server.err" -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "'+$repo+'/studio/Server.ps1" -RuntimeRoot "'+$root+'" -Port '+$Port+' -Instance source-http -WorkerPid '+$PID)
[Net.WebRequest]::DefaultWebProxy=$null;$url="http://127.0.0.1:$Port/";$headers=@{}
function Check($Name,$Passed){if(-not $Passed){throw "FAIL: $Name"};"PASS: $Name"}
function Post($Route,$Value){Invoke-RestMethod -Method Post ($url+$Route) -Headers $headers -ContentType application/json -Body (ConvertTo-Json -InputObject $Value -Depth 20) -TimeoutSec 120}
try{
 $boot=$null;for($n=0;$n -lt 40;$n++){try{$boot=Invoke-RestMethod ($url+'api/bootstrap') -TimeoutSec 1;break}catch{Start-Sleep -Milliseconds 250}}
 if(-not $boot){throw 'Test server did not start.'};$headers=@{'X-LF-Token'=$boot.token}
 $p=Post 'api/new-source' @{name='HTTP native source import';choice=$choice;preset='renewable-energy'}
 Check 'Native import endpoint persists real external path' ($p.source.external_path -eq $external)
 $id=$p.id;$oldHash=$p.source.sha256
 $dir=Get-LFProjectPath $root $id
 function Inventory { @(Get-ChildItem -LiteralPath $dir -Recurse -Force -File | Sort-Object FullName | ForEach-Object {@{path=$_.FullName;sha256=(Get-FileHash -LiteralPath $_.FullName).Hash;written=$_.LastWriteTimeUtc.Ticks}}) | ConvertTo-Json -Compress }
 $unchanged=Inventory
 $c=Post "api/source-compare?id=$id" @{}
 Check 'Comparison endpoint reports unchanged compatible source' ($c.compatible -and -not $c.changed)
 $noop=Post "api/source-reload?id=$id" @{path=$external;expected=$c.expected;confirm=$true}
 Check 'Unchanged HTTP reload returns already up to date without revision change' (-not $noop.changed -and $noop.project.version -eq $p.version -and $noop.message -eq 'Source PowerPoint is already up to date.')
 Check 'Unchanged HTTP comparison and reload write no project files or backups' ((Inventory) -ceq $unchanged -and -not (Test-Path "$root/SourceBackups"))
 $z=[IO.Compression.ZipFile]::Open($external,'Update')
 try{$e=$z.GetEntry('ppt/slides/slide1.xml');$r=[IO.StreamReader]::new($e.Open());try{[xml]$x=$r.ReadToEnd()}finally{$r.Dispose()};$x.SelectSingleNode("//*[local-name()='t']").InnerText='HTTP refreshed content';$e.Delete();$w=[IO.StreamWriter]::new($z.CreateEntry('ppt/slides/slide1.xml').Open());try{$x.Save($w)}finally{$w.Dispose()}}finally{$z.Dispose()}
 $c=Post "api/source-compare?id=$id" @{}
 Check 'Comparison endpoint exposes both hashes and identities/dimensions' ($c.changed -and $c.compatible -and $c.old.sha256 -eq $oldHash -and $c.new.sha256 -ne $oldHash -and $c.ordered_ids_match -and $c.dimensions_match)
 $rejected=$false;try{$null=Post "api/source-reload?id=$id" @{path=$external;expected=$c.expected;confirm=$false}}catch{$rejected=$true}
 Check 'Reload endpoint refuses missing confirmation' $rejected
 $result=Post "api/source-reload?id=$id" @{path=$external;expected=$c.expected;confirm=$true}
 Check 'HTTP reload commits new source and renders previews' ($result.changed -and $result.project.source.sha256 -eq $c.new.sha256 -and (Test-Path "$root/Projects/$id/thumbs/s2.png"))
 $returned=Invoke-RestMethod ($url+"api/project?id=$id")
 Check 'Reopened HTTP project observes committed metadata' ($returned.source.sha256 -eq $c.new.sha256)
 $served=Invoke-WebRequest ($url+'source.js') -UseBasicParsing
 Check 'Source controls script served with application page' ($served.StatusCode -eq 200 -and $served.Content -match 'confirmSourceReload')
 Check 'Synthetic HTTP project generated no narration/avatar state' (-not (Test-Path "$root/Projects/$id/narration.json") -and -not (Test-Path "$root/Projects/$id/avatar-queue.json"))
 Write-LFJson "$root/results.json" @{passed=$true;provider_calls=0;project=$id;source=$external;backup=$result.backup}
 "PASS: HTTP source refresh. Evidence: $root"
}finally{
 if($headers.Count){try{$null=Post 'api/stop' @{}}catch{}}
 if(-not $server.WaitForExit(5000)){Stop-Process -Id $server.Id}
}
