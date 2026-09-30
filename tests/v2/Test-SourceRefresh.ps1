param([switch]$DesktopPowerPoint)
$ErrorActionPreference='Stop'
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Import-Module "$repo/studio/Assembly.psm1" -Force
Import-Module "$repo/studio/SourceRefresh.psm1" -Force
Import-Module "$repo/studio/Production.psm1"
Add-Type -AssemblyName System.IO.Compression.FileSystem
$root=Join-Path $repo ('work/sr-'+[guid]::NewGuid().ToString('N').Substring(0,8));[void][IO.Directory]::CreateDirectory($root)
function Check($Name,$Pass){if(-not $Pass){throw "FAIL: $Name"};"PASS: $Name"}
function Reject($Name,[scriptblock]$Code,[string]$Pattern){$caught='';try{& $Code|Out-Null}catch{$caught=$_.Exception.Message};Check $Name ($caught -match $Pattern)}
$source=Join-Path $root 'source with spaces.pptx'
if($DesktopPowerPoint){
 $app=$null;$deck=$null;$wasRunning=@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
 try{$app=New-Object -ComObject PowerPoint.Application;$deck=$app.Presentations.Add(0);$deck.PageSetup.SlideWidth=960;$deck.PageSetup.SlideHeight=540
  foreach($n in 1,2){$slide=$deck.Slides.Add($n,12);$shape=$slide.Shapes.AddTextbox(1,40,40,800,100);$shape.TextFrame.TextRange.Text="Synthetic slide $n"}
  $deck.SaveAs($source,24)
 }finally{if($deck){$deck.Close()};if($app -and -not $wasRunning){$app.Quit()};foreach($com in @($deck,$app)){if($com){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com)}}}
}else{
 $zip=[IO.Compression.ZipFile]::Open($source,'Create')
 function Part($Name,$Text){$w=[IO.StreamWriter]::new($zip.CreateEntry($Name).Open());try{$w.Write($Text)}finally{$w.Dispose()}}
 try{
  Part 'ppt/presentation.xml' '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>'
  Part 'ppt/_rels/presentation.xml.rels' '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/><Relationship Id="rId2" Target="slides/slide2.xml"/></Relationships>'
  foreach($n in 1,2){Part "ppt/slides/slide$n.xml" "<slide><t>Synthetic slide $n</t></slide>"}
 }finally{$zip.Dispose()}
}
function Edit-Part($Path,$Part,[scriptblock]$Edit){
 $z=[IO.Compression.ZipFile]::Open($Path,'Update')
 try{$e=$z.GetEntry($Part);$r=[IO.StreamReader]::new($e.Open());try{[xml]$x=$r.ReadToEnd()}finally{$r.Dispose()}; & $Edit $x;$e.Delete();$w=[IO.StreamWriter]::new($z.CreateEntry($Part).Open());try{$x.Save($w)}finally{$w.Dispose()}}finally{$z.Dispose()}
}
$p=New-LFProject $root 'Synthetic source refresh' ([IO.Path]::GetFileName($source)) ([IO.File]::ReadAllBytes($source)) 'renewable-energy' $source
$id=$p.id;$dir=Get-LFProjectPath $root $id
Check 'Future import remembers the absolute external path' ($p.source.external_path -eq $source)
# Production-state fixtures with actual hash bindings and immutable authorization receipts.
$revisions=@();foreach($n in 1,2){
 $p.slides[$n-1].script="Approved narration $n"
 $revision=('a'*11)+$n;$relative="n/$revision/bbbbbbbbbbbb/take.mp3";$audio=Join-Path $dir $relative
 [void][IO.Directory]::CreateDirectory((Split-Path $audio));[IO.File]::WriteAllText($audio,"Synthetic approved audio $n")
 $asset=@{path=$relative;sha256=(Get-FileHash $audio).Hash;decode_ok=$true;duration_seconds=1;post_speech_silence_seconds=0;pause_samples=0;pause_sample_rate=44100}
 $revisions+=@{id=$revision;slide_number=$n;preview=$false;script_sha256=Get-LFTextHash $p.slides[$n-1].script;binding=Get-LFBinding $p $p.slides[$n-1];selected_take=2;pause_seconds=0;takes=@(@{number=2;state='ready';asset=$asset;attempts=@(@{id='synthetic-approved';state='complete';submitted_utc='2026-01-01T00:00:00Z'})})}
}
Write-LFJson "$dir/project.json" $p;Write-LFJson "$dir/narration.json" @{schema='lectureforge-narration-1';revisions=$revisions;intent_revision=2}
[void][IO.Directory]::CreateDirectory("$dir/approvals")
$selection=Get-LFSelectionInputs $root $id $dir;$auth=$selection.binding_sha256
Write-LFImmutableJson (Get-LFApprovalPath $dir $auth authorization) @{id=$auth;binding_sha256=$auth;snapshot=$selection.snapshot}
Write-LFImmutableJson (Get-LFApprovalPath $dir $auth consumed) @{consumer='lectureforge-avatar-worker-1';authorization=$auth;authorization_sha256=(Get-FileHash (Get-LFApprovalPath $dir $auth authorization)).Hash}
$jobs=@();foreach($row in $selection.snapshot.slides){
 $jid=('b'*15)+$row.number;$jobs+=$jid;[void][IO.Directory]::CreateDirectory("$dir/a/$jid")
 $relative="a/$jid/w-12345678.mp4";[IO.File]::WriteAllText((Join-Path $dir $relative),"Synthetic existing avatar $($row.number)")
 $hash=(Get-FileHash (Join-Path $dir $relative)).Hash
 Write-LFJson "$dir/a/$jid/job.json" @{id=$jid;slide=$row.number;take=2;stage='Avatar Ready';error=$null;lease=$null;authorization=$auth;row=$row;preset=$selection.snapshot.avatar_preset;output_path=$relative;output_sha256=$hash;video_id="existing-heygen-$($row.number)";provider_terminal=$true;submission_intent=$true;provider_calls=1;recovery=@{preserved='existing history'};validation=@{passed=$true;output_sha256=$hash;narration_sha256=$row.narration.sha256;authoritative_audio_verified=$true;pause_verified=$true;prepared_pause_seconds=0}}
}
Write-LFJson "$dir/avatar-queue.json" @{schema='lectureforge-avatar-queue-1';paused=$false;jobs=$jobs;batches=@(@{authorization=$auth;jobs=$jobs})}
Write-LFJson "$dir/placement.json" @{version=3;overrides=@()}
Write-LFJson "$dir/selection-review.json" @{review_id=$auth}
$prior=@{id='12345678';stage='Complete';lease=$null;binding='old';output='existing-recording.pptx'}
Write-LFJson "$dir/assembly.json" @{schema='lectureforge-assembly-1';current=$prior;last_success=$prior;history=@()}
[void][IO.Directory]::CreateDirectory("$dir/thumbs");[IO.File]::WriteAllText("$dir/thumbs/s1.png",'old thumbnail');Write-LFJson "$dir/thumbs/status.json" @{state='ready';source_sha256=$p.source.sha256}
$readyBefore=Get-LFAssemblyInputs $root $id $dir
Check 'Synthetic production fixture is assembly-ready' ($readyBefore.ready -and $readyBefore.snapshot.slides.Count -eq 2)
$protected=@{};foreach($f in Get-ChildItem $dir -Recurse -File){if($f.Name -notin 'source.pptx','project.json','assembly.json' -and $f.FullName -notlike '*\thumbs\*'){$protected[$f.FullName]=(Get-FileHash $f.FullName).Hash}}
$module=Get-Module SourceRefresh
if(-not $DesktopPowerPoint){& $module {function script:Export-LFSourceThumbnails($Source,$Destination,$Count,$Width,$Height){[void][IO.Directory]::CreateDirectory($Destination);foreach($n in 1..$Count){[IO.File]::WriteAllText((Join-Path $Destination "s$n.png"),'new thumbnail')}}}}
function NoChangeInventory {
 @(Get-ChildItem -LiteralPath $dir -Recurse -Force -File | Sort-Object FullName | ForEach-Object {
  @{path=$_.FullName;sha256=(Get-FileHash -LiteralPath $_.FullName).Hash;written=$_.LastWriteTimeUtc.Ticks;created=$_.CreationTimeUtc.Ticks}
 }) | ConvertTo-Json -Depth 5 -Compress
}
$unchangedInventory=NoChangeInventory
$same=Get-LFSourceComparison $root $id
Check 'Unchanged source comparison reports all compatibility fields' (-not $same.changed -and $same.compatible -and $same.ordered_ids_match -and $same.dimensions_match)
$noChange=Invoke-LFSourceRefresh $root $id $source $same.expected $true
Check 'Unchanged reload performs no backup or revision change' (-not $noChange.changed -and $noChange.project.version -eq $p.version)
Check 'Unchanged reload creates no backup or staging directory' (-not (Test-Path (Join-Path $root 'SourceBackups')))
Check 'Unchanged comparison and reload preserve every project file byte and write time' ((NoChangeInventory) -ceq $unchangedInventory)
Check 'Unchanged reload reports already up to date' ($noChange.message -ceq 'Source PowerPoint is already up to date.')
Edit-Part $source 'ppt/slides/slide1.xml' {param($x) $x.SelectSingleNode("//*[local-name()='t']").InnerText='Edited source content'}
$c=Get-LFSourceComparison $root $id
Check 'Compatible edit changes hash, not slide identity/order/size' ($c.changed -and $c.compatible -and $c.old.slide_count -eq 2 -and $c.new.slide_count -eq 2)
Reject 'Confirmation is mandatory' {Invoke-LFSourceRefresh $root $id $source $c.expected $false} 'Confirm'
Reject 'Stale comparison cannot apply' {Invoke-LFSourceRefresh $root $id $source 'old-token' $true} 'changed since comparison'
foreach($kind in 'count','order','dimensions'){
 $bad=Join-Path $root ($kind+'.pptx');[IO.File]::Copy($source,$bad)
 Edit-Part $bad 'ppt/presentation.xml' {param($x)
  $list=$x.presentation.sldIdLst
  if($kind -eq 'count'){[void]$list.RemoveChild($list.LastChild)}elseif($kind -eq 'order'){[void]$list.AppendChild($list.FirstChild)}else{$x.presentation.sldSz.SetAttribute('cx','999999')}
 }
 $cmp=Get-LFSourceComparison $root $id $bad
 Check "$kind incompatibility is explained" (-not $cmp.compatible -and $cmp.problems.Count -gt 0)
 Reject "$kind incompatibility cannot commit" {Invoke-LFSourceRefresh $root $id $bad $cmp.expected $true} 'Reload stopped'
}
foreach($stage in 'Queued','HeyGen Processing','White Avatar Building'){
 $jobPath="$dir/a/$($jobs[0])/job.json";$bytes=[IO.File]::ReadAllBytes($jobPath);$j=Read-LFSharedJson $jobPath;$j.stage=$stage;Write-LFJson $jobPath $j
 try{Reject "$stage blocks reload" {Invoke-LFSourceRefresh $root $id $source $c.expected $true} 'Avatar work'}finally{[IO.File]::WriteAllBytes($jobPath,$bytes)}
}
$nPath="$dir/narration.json";$bytes=[IO.File]::ReadAllBytes($nPath);$n=Read-LFSharedJson $nPath;$n.revisions[0].takes[0].state='queued';Write-LFJson $nPath $n
try{Reject 'Pending narration blocks reload' {Invoke-LFSourceRefresh $root $id $source $c.expected $true} 'Narration work'}finally{[IO.File]::WriteAllBytes($nPath,$bytes)}
foreach($stateFile in 'assembly.json','thumbs/status.json'){
 $statePath=Join-Path $dir $stateFile;$bytes=[IO.File]::ReadAllBytes($statePath);$state=Read-LFSharedJson $statePath
 if($stateFile -eq 'assembly.json'){$state.current.stage='Queued'}else{$state.state='exporting'}
 Write-LFJson $statePath $state
 try{Reject ("Active $stateFile blocks reload") {Invoke-LFSourceRefresh $root $id $source $c.expected $true} 'active|exported'}finally{[IO.File]::WriteAllBytes($statePath,$bytes)}
}
$receipt=Get-LFApprovalPath $dir $auth consumed;$bytes=[IO.File]::ReadAllBytes($receipt);Remove-Item -LiteralPath $receipt
try{Reject 'Pending unconsumed authorization blocks reload' {Invoke-LFSourceRefresh $root $id $source $c.expected $true} 'authorized production batch'}finally{[IO.File]::WriteAllBytes($receipt,$bytes)}
$lockFile=Join-Path $root ('~$'+[IO.Path]::GetFileName($source));[IO.File]::WriteAllText($lockFile,'PowerPoint owner fixture')
try{Reject 'Open external deck must be saved and closed' {Get-LFSourceComparison $root $id} 'Save and close PowerPoint'}finally{Remove-Item -LiteralPath $lockFile}
# Export failure must occur before any source commit.
$renderer=& $module {(Get-Command Export-LFSourceThumbnails).ScriptBlock}
& $module {function script:Export-LFSourceThumbnails {throw 'synthetic thumbnail failure'}}
Reject 'Thumbnail failure leaves existing source untouched' {Invoke-LFSourceRefresh $root $id $source $c.expected $true} 'synthetic thumbnail failure'
Check 'Pre-commit failure leaves source SHA unchanged' ((Get-FileHash "$dir/source.pptx").Hash -eq $c.old.sha256)
& $module {param($body) Set-Item -LiteralPath Function:script:Export-LFSourceThumbnails -Value $body} $renderer
$inventoryFunction=& $module {(Get-Command Get-LFSourceInventory).ScriptBlock}
& $module {
 $script:realSourceInventory=(Get-Command Get-LFSourceInventory).ScriptBlock
 function script:Get-LFSourceInventory($Directory){$items=@(& $script:realSourceInventory $Directory);if($Directory -like '*SourceBackups*'){$items[0].sha256='synthetic-backup-mismatch'};$items}
}
Reject 'Unverified backup prevents any source change' {Invoke-LFSourceRefresh $root $id $source $c.expected $true} 'Backup verification failed'
& $module {param($body) Set-Item -LiteralPath Function:script:Get-LFSourceInventory -Value $body} $inventoryFunction
Check 'Backup failure retains original snapshot' ((Get-FileHash "$dir/source.pptx").Hash -eq $c.old.sha256)
# A separate process owns the same project mutex, as a worker or thumbnail export would.
$lockScript=Join-Path $root 'hold-project-lock.ps1'
$lockBody=@'
param($Repo,$Root,$Id,$Signal)
Import-Module "$Repo/studio/Production.psm1"
Invoke-LFProjectLock $Root $Id {param($dir) [IO.File]::WriteAllText($Signal,'locked');Start-Sleep -Seconds 3}
'@
[IO.File]::WriteAllText($lockScript,$lockBody)
$signal=Join-Path $root 'lock-ready'
$holder=Start-Process (Join-Path $PSHOME 'powershell.exe') -WindowStyle Hidden -PassThru -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "'+$lockScript+'" -Repo "'+$repo+'" -Root "'+$root+'" -Id '+$id+' -Signal "'+$signal+'"')
try{
 for($i=0;$i -lt 40 -and -not (Test-Path $signal);$i++){Start-Sleep -Milliseconds 100}
 Check 'Independent project-lock fixture acquired the mutex' (Test-Path $signal)
 Reject 'Concurrent project work blocks source replacement' {Invoke-LFSourceRefresh $root $id $source $c.expected $true} 'Project work is active'
}finally{if(-not $holder.WaitForExit(5000)){Stop-Process -Id $holder.Id}}
# Inject a failure after source copy: production code must restore byte-identical old state.
$projectHash=(Get-FileHash "$dir/project.json").Hash;$sourceHash=(Get-FileHash "$dir/source.pptx").Hash
& $module {function script:Set-LFReloadedProject($Directory,$Project){throw 'synthetic commit failure'}}
Reject 'Commit failure automatically rolls back' {Invoke-LFSourceRefresh $root $id $source $c.expected $true} 'rolled back.*synthetic commit failure'
Check 'Rollback restores snapshot, metadata and preview' ((Get-FileHash "$dir/project.json").Hash -eq $projectHash -and (Get-FileHash "$dir/source.pptx").Hash -eq $sourceHash -and [IO.File]::ReadAllText("$dir/thumbs/s1.png") -eq 'old thumbnail')
& $module {function script:Set-LFReloadedProject($Directory,$Project){Write-LFJson (Join-Path $Directory 'project.json') $Project}}
$result=Invoke-LFSourceRefresh $root $id $source $c.expected $true
$readyAfter=Get-LFAssemblyInputs $root $id $dir
Check 'Reload keeps ready avatar rows identical without new authorization' ($readyAfter.ready -and (ConvertTo-LFCanonicalJson $readyBefore.snapshot.slides) -ceq (ConvertTo-LFCanonicalJson $readyAfter.snapshot.slides))
Check 'New source changes assembly binding' ($readyAfter.binding -ne $readyBefore.binding)
Check 'Slide scripts, selections and project preset preserved' ((ConvertTo-LFCanonicalJson $p.slides) -ceq (ConvertTo-LFCanonicalJson $result.project.slides) -and (ConvertTo-LFCanonicalJson $p.preset) -ceq (ConvertTo-LFCanonicalJson $result.project.preset))
foreach($path in $protected.Keys){Check ('Preserved '+$path.Substring($dir.Length)) ((Get-FileHash -LiteralPath $path).Hash -eq $protected[$path])}
$a=Read-LFSharedJson "$dir/assembly.json"
Check 'Prior recording archived and no automatic assembly queued' (-not $a.current -and -not $a.last_success -and $a.history[0].id -eq '12345678')
Check 'Verified backup contains old snapshot and all assets' ((Test-Path $result.backup) -and (Get-FileHash "$($result.backup)/source.pptx").Hash -eq $sourceHash)
Check 'Previews regenerated against new source hash' ((Read-LFSharedJson "$dir/thumbs/status.json").source_sha256 -eq $c.new.sha256 -and (Test-Path "$dir/thumbs/s2.png"))
# Interrupted commit journal: next ordinary project read recovers before exposing mixed state.
$tx=Split-Path (Split-Path $result.backup) -Leaf
[IO.File]::WriteAllText("$dir/source.pptx",'partial write')
Write-LFJson "$dir/source-refresh-transaction.json" @{id=$tx}
$restored=Read-LFProject $root $id
Check 'Interrupted transaction rolls back on next project read' ($restored.source.sha256 -eq $sourceHash -and -not (Test-Path "$dir/source-refresh-transaction.json"))
# Legacy project can compare and attach an external path without discarding state.
$raw=Read-LFSharedJson "$dir/project.json";$raw.source.PSObject.Properties.Remove('external_path');Write-LFJson "$dir/project.json" $raw
Check 'Legacy reload requests a native selection' (Get-LFSourceComparison $root $id).needs_selection
$legacy=Get-LFSourceComparison $root $id $source
$result=Invoke-LFSourceRefresh $root $id $source $legacy.expected $true
Check 'Successful legacy confirmation saves the path' ($result.project.source.external_path -eq $source)
Reject 'Missing source gives a choose-different action' {Get-LFSourceComparison $root $id "$root/missing.pptx"} 'Choose Different Source PowerPoint'
# Shell launch tests capture real arguments and simulate a launch error; no Explorer window in unit suite.
& $module {function script:Invoke-LFExplorerOpen($Path,$SelectFile){$global:sourceExplorerPath=$Path;$global:sourceExplorerSelect=$SelectFile}}
$null=Open-LFProjectLocation $root $id source
Check 'Source Explorer action selects exact path with spaces' ($global:sourceExplorerPath -eq $source -and $global:sourceExplorerSelect)
$null=Open-LFProjectLocation $root $id project
Check 'Project Explorer action opens internal project directory' ($global:sourceExplorerPath -eq $dir -and -not $global:sourceExplorerSelect)
& $module {function script:Invoke-LFExplorerOpen {throw 'synthetic Explorer failure'}}
Reject 'Explorer failure includes resolved path and useful action' {Open-LFProjectLocation $root $id project} 'Could not open Windows Explorer.*copy this path'
Write-LFJson "$root/results.json" @{passed=$true;desktop_powerpoint=[bool]$DesktopPowerPoint;project=$id;root=$root;backup=$result.backup;protected_files=$protected.Count;provider_calls=0;ready_before=$readyBefore.ready;ready_after=$readyAfter.ready}
"PASS: source-refresh preservation verification. Evidence: $root"
