$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Narration.psm1')
Import-Module (Join-Path $PSScriptRoot 'Project.psm1')
Import-Module (Join-Path $PSScriptRoot 'Authorization.psm1')

function Get-LFSourceDescription([string]$Path){
 $asset=Get-V2Asset $Path;$info=Get-V2DeckInfo $Path
 $zip=[IO.Compression.ZipFile]::OpenRead($Path)
 try{
  $reader=[IO.StreamReader]::new($zip.GetEntry('ppt/presentation.xml').Open())
  try{[xml]$xml=$reader.ReadToEnd()}finally{$reader.Dispose()}
  $ids=@($xml.presentation.sldIdLst.sldId|ForEach-Object {$_.GetAttribute('id')})
  if($ids.Count -ne $info.slide_count -or @($ids|Sort-Object -Unique).Count -ne $ids.Count -or $ids.Count -lt 1){throw 'Missing or duplicate PowerPoint slide identities.'}
  if((Get-V2Asset $Path).sha256 -ne $asset.sha256){throw 'PowerPoint changed during comparison. Save and close it, then retry.'}
  [pscustomobject]@{sha256=$asset.sha256;slide_count=$info.slide_count;width_emu=$info.width_emu;height_emu=$info.height_emu;slide_ids=$ids}
 }finally{$zip.Dispose()}
}
function Resolve-LFExternalSource([string]$Path){
 if(-not $Path){throw 'Choose Different Source PowerPoint to locate the source file.'}
 if(-not [IO.Path]::IsPathRooted($Path) -or $Path -match '["\r\n]'){throw 'Choose an absolute PowerPoint file path.'}
 $full=[IO.Path]::GetFullPath($Path)
 if([IO.Path]::GetExtension($full) -ne '.pptx' -or -not (Test-Path -LiteralPath $full -PathType Leaf)){throw "Source PowerPoint not found: $full. Use Choose Different Source PowerPoint."}
 if(Test-Path -LiteralPath (Join-Path (Split-Path $full) ('~$'+[IO.Path]::GetFileName($full)))){throw "Save and close PowerPoint before reloading: $full"}
 $full
}
function Assert-LFSourceIdle([string]$Dir){
 $file=Join-Path $Dir 'assembly.json'
 if(Test-Path $file){$a=Read-LFSharedJson $file;if($a.current -and ($a.current.lease -or $a.current.stage -notin 'Complete','Exception')){throw 'Recording assembly is active. Wait for it to finish before reloading.'}}
 $file=Join-Path $Dir 'narration.json'
 if(Test-Path $file){$n=Read-LFSharedJson $file;foreach($r in @($n.revisions)){foreach($t in @($r.takes)){if($t.state -in 'queued','dispatched','generating','uncertain' -or $t.lease){throw 'Narration work is pending or unresolved. Finish it before reloading.'}}}}
 foreach($file in @(Get-ChildItem -Path (Join-Path $Dir 'a/*/job.json') -ErrorAction SilentlyContinue)){
  $j=Read-LFSharedJson $file.FullName
  if($j.lease -or $j.stage -notin 'Avatar Ready','Exception' -or ($j.stage -eq 'Exception' -and $j.submission_intent -and -not $j.provider_terminal -and -not $j.blocked_by_job)){throw 'Avatar work is pending or unresolved. Finish or recover it before reloading.'}
 }
 # An unconsumed paid authorization can otherwise be picked up as soon as the lock releases.
 foreach($f in @(Get-ChildItem -LiteralPath (Join-Path $Dir 'approvals') -Filter '*.authorization.json' -ErrorAction SilentlyContinue)){
  if(-not (Test-Path -LiteralPath ($f.FullName -replace '\.authorization\.json$','.consumed.json'))){
   $approval=Read-LFSharedJson $f.FullName
   $current=Get-LFSelectionInputs (Split-Path (Split-Path $Dir)) (Split-Path $Dir -Leaf) $Dir
   $binding=if($approval.snapshot.selection_binding){$approval.snapshot.selection_binding}else{$approval.binding_sha256}
   if($current.snapshot.can_authorize -and $current.binding_sha256 -eq $binding){throw 'An authorized production batch is pending. Finish it before reloading.'}
  }
 }
 $file=Join-Path $Dir 'thumbs/status.json'
 if(Test-Path $file){$s=Read-LFSharedJson $file;if($s.state -eq 'exporting'){throw 'Slide previews are being exported. Wait for completion before reloading.'}}
}
function Get-LFSourceComparison($Root,$Id,[string]$ExternalPath=''){
 Invoke-LFProjectLock $Root $Id {
  param($dir)
  $p=Read-LFProject $Root $Id
  if(-not $ExternalPath){$ExternalPath=$p.source.external_path}
  if(-not $ExternalPath){return @{needs_selection=$true}}
  $path=Resolve-LFExternalSource $ExternalPath
  if($path.StartsWith(([IO.Path]::GetFullPath($dir).TrimEnd('\')+'\'),[StringComparison]::OrdinalIgnoreCase)){throw 'Choose the external source, not a project snapshot or recording.'}
  $old=Get-LFSourceDescription (Join-Path $dir 'source.pptx');$new=Get-LFSourceDescription $path
  $order=($old.slide_ids -join ',') -ceq ($new.slide_ids -join ',')
  $dimensions=$old.width_emu -eq $new.width_emu -and $old.height_emu -eq $new.height_emu
  $problems=@();if($old.slide_count -ne $new.slide_count){$problems+='Slide count changed.'};if(-not $order){$problems+='Ordered slide identities changed.'};if(-not $dimensions){$problems+='Slide dimensions changed.'}
  $token=Get-LFTextHash ($p.version.ToString()+'|'+$old.sha256+'|'+$new.sha256+'|'+$path)
  @{needs_selection=$false;path=$path;old=$old;new=$new;ordered_ids_match=$order;dimensions_match=$dimensions;compatible=($problems.Count -eq 0);problems=$problems;changed=($old.sha256 -ne $new.sha256);expected=$token}
 } 0
}
function Get-LFSourceInventory([string]$Directory){
 $files=@(Get-ChildItem -LiteralPath $Directory -Force -Recurse)
 if(@($files|Where-Object {$_.Attributes -band [IO.FileAttributes]::ReparsePoint}).Count){throw 'Project backup cannot follow junctions or symbolic links.'}
 @($files|Where-Object {-not $_.PSIsContainer}|Sort-Object FullName|ForEach-Object {@{path=$_.FullName.Substring($Directory.Length).TrimStart('\');sha256=(Get-FileHash -LiteralPath $_.FullName).Hash}})
}
function Export-LFSourceThumbnails([string]$Source,[string]$Destination,[int]$Count,[long]$Width,[long]$Height){
 [void][IO.Directory]::CreateDirectory($Destination)
 $app=$null;$deck=$null;$slide=$null;$wasRunning=@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
 try{
  $app=New-Object -ComObject PowerPoint.Application;$deck=$app.Presentations.Open($Source,$true,$false,$false)
  if($deck.Slides.Count -ne $Count){throw 'PowerPoint slide count differs from the package.'}
  for($n=1;$n -le $Count;$n++){
   $slide=$deck.Slides.Item($n)
   try{$slide.Export((Join-Path $Destination "s$n.png"),'PNG',640,[int](640*$Height/$Width))}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($slide);$slide=$null}
  }
 }finally{
  if($deck){$deck.Close()};if($app -and -not $wasRunning -and $app.Presentations.Count -eq 0){$app.Quit()}
  foreach($com in @($slide,$deck,$app)){if($com){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com)}}
 }
 foreach($n in 1..$Count){if(-not (Test-Path -LiteralPath (Join-Path $Destination "s$n.png"))){throw 'Thumbnail export incomplete.'}}
}
function Set-LFReloadedProject([string]$Directory,$Project){Write-LFJson (Join-Path $Directory 'project.json') $Project}
function Invoke-LFSourceRefresh($Root,$Id,[string]$ExternalPath,[string]$Expected,[bool]$Confirm){
 if(-not $Confirm){throw 'Confirm the source comparison before reloading.'}
 Invoke-LFProjectLock $Root $Id {
  param($dir)
  Assert-LFSourceIdle $dir
  $comparison=Get-LFSourceComparison $Root $Id $ExternalPath
  if($comparison.needs_selection -or $comparison.expected -ne $Expected){throw 'Source or project changed since comparison. Compare again.'}
  if(-not $comparison.compatible){throw (($comparison.problems -join ' ')+' Reload stopped; production state was not remapped.')}
  $path=$comparison.path
  # Read raw metadata so serialization does not persist unrelated default normalization.
  $p=Read-LFSharedJson (Join-Path $dir 'project.json')
  if(-not $comparison.changed -and $p.source.external_path -eq $path){return @{changed=$false;project=$p;message='Source PowerPoint is already up to date.'}}
  $tx=[guid]::NewGuid().ToString('N');$txDir=Join-Path (Get-LFRoot $Root) "SourceBackups/$Id/$tx"
  $backup=Join-Path $txDir 'project';[void][IO.Directory]::CreateDirectory($txDir)
  $inventory=Get-LFSourceInventory $dir
  Copy-Item -LiteralPath $dir -Destination $backup -Recurse
  $verified=Get-LFSourceInventory $backup
  # Compare by relative paths and hash, independent of directory traversal order.
  $signature={param($items) (@($items|ForEach-Object {$_.path+'='+$_.sha256}|Sort-Object) -join "`n")}
  if((& $signature $inventory) -cne (& $signature $verified)){throw "Backup verification failed: $backup"}
  Write-LFJson (Join-Path $txDir 'inventory.json') $inventory
  $staged=Join-Path $txDir 'replacement.pptx'
  $input=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
  try{$output=[IO.File]::Open($staged,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None);try{$input.CopyTo($output)}finally{$output.Dispose()}}finally{$input.Dispose()}
  if((Get-V2Asset $staged).sha256 -ne $comparison.new.sha256){throw 'External PowerPoint changed during staging. Compare again.'}
  $stagedInfo=Get-LFSourceDescription $staged
  if($stagedInfo.slide_count -ne $comparison.old.slide_count -or ($stagedInfo.slide_ids -join ',') -cne ($comparison.old.slide_ids -join ',') -or $stagedInfo.width_emu -ne $comparison.old.width_emu -or $stagedInfo.height_emu -ne $comparison.old.height_emu){throw 'Staged source is incompatible. Reload stopped.'}
  $thumbs=Join-Path $txDir 'thumbnails'
  if($comparison.changed){
   Export-LFSourceThumbnails $staged $thumbs $comparison.new.slide_count $comparison.new.width_emu $comparison.new.height_emu
   Write-LFJson (Join-Path $thumbs 'status.json') @{state='ready';source_sha256=$comparison.new.sha256}
  }
  # No project-owned data has changed before this durable rollback journal.
  $journal=Join-Path $dir 'source-refresh-transaction.json'
  Write-LFJson $journal @{id=$tx}
  try{
   [IO.File]::Copy($staged,(Join-Path $dir 'source.pptx'),$true)
   $p.source.sha256=$comparison.new.sha256;$p.source.original_name=[IO.Path]::GetFileName($path)
   $p.source|Add-Member -Force NoteProperty external_path $path
   if($comparison.changed){$p.version++}
   $p|Add-Member -Force NoteProperty source_refreshed_utc ([DateTime]::UtcNow.ToString('o'))
   Set-LFReloadedProject $dir $p
   if($comparison.changed){
    $dest=[IO.Path]::GetFullPath((Join-Path $dir 'thumbs'))
    if($dest -ne ([IO.Path]::GetFullPath($dir).TrimEnd('\')+'\thumbs')){throw 'Unsafe thumbnail path.'}
    if(Test-Path -LiteralPath $dest){Remove-Item -LiteralPath $dest -Recurse -Force}
    Copy-Item -LiteralPath $thumbs -Destination $dest -Recurse
    $file=Join-Path $dir 'assembly.json'
    if(Test-Path $file){
     $a=Read-LFSharedJson $file
     if($a.current){$a.history+=,$a.current}
     if($a.last_success -and $a.last_success.id -notin @($a.history|ForEach-Object id)){$a.history+=,$a.last_success}
     # Old recordings remain on disk and in history, but cannot be presented as current.
     $a.current=$null;$a.last_success=$null;Write-LFJson $file $a
    }
   }
   if((Get-V2Asset (Join-Path $dir 'source.pptx')).sha256 -ne (Read-LFSharedJson (Join-Path $dir 'project.json')).source.sha256){throw 'Source commit verification failed.'}
   Remove-Item -LiteralPath $journal
  }catch{
   $failure=$_.Exception.Message
   try{Restore-LFSourceTransaction $dir}catch{throw "Reload failed: $failure. Rollback needs attention; verified backup: $backup. $($_.Exception.Message)"}
   throw "Reload failed and was rolled back: $failure"
  }
  @{changed=$comparison.changed;project=$p;backup=$backup;message='Source reloaded. Existing narration and avatar assets were preserved. Create a new recording when ready.'}
 } 0
}
function Invoke-LFExplorerOpen([string]$Path,[bool]$SelectFile){
 $folder=if($SelectFile){Split-Path $Path}else{$Path}
 $shell=New-Object -ComObject Shell.Application
 try{
  # Ask the desktop shell directly; explorer.exe exit codes do not reliably
  # describe whether an existing Explorer process accepted a folder request.
  $shell.Open($folder)
  for($i=0;$i -lt 40;$i++){
   foreach($window in @($shell.Windows())){
    try{
     if($window.Document.Folder.Self.Path -ne $folder){continue}
     if($SelectFile){
      $item=$window.Document.Folder.ParseName([IO.Path]::GetFileName($Path))
      if(-not $item){continue}
      # Select, deselect other items, scroll into view, and focus.
      $window.Document.SelectItem($item,29)
      if($Path -notin @($window.Document.SelectedItems()|ForEach-Object Path)){continue}
     }
     return
    }catch{continue}
   }
   Start-Sleep -Milliseconds 125
  }
  throw 'Windows did not expose the requested Explorer folder within five seconds.'
 }finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($shell)}
}
function Open-LFProjectLocation($Root,$Id,[ValidateSet('project','source')][string]$Kind){
 $dir=Get-LFProjectPath $Root $Id
 $path=if($Kind -eq 'source'){$p=Read-LFProject $Root $Id;$p.source.external_path}else{$dir}
 if(-not $path -or -not (Test-Path -LiteralPath $path)){throw "Location unavailable: $path. $(if($Kind -eq 'source'){'Use Choose Different Source PowerPoint.'}else{'The project folder is missing.'})"}
 $path=[IO.Path]::GetFullPath($path)
 if($path -match '["\r\n]'){throw "Unsupported Explorer path: $path"}
 try{
  Invoke-LFExplorerOpen $path ($Kind -eq 'source')
 }catch{throw "Could not open Windows Explorer for $path. $($_.Exception.Message) You can copy this path into Explorer."}
 @{opened=$true;path=$path;message="Opened Windows Explorer: $path"}
}
function Start-LFSourceChoice($Root){
 $id=[guid]::NewGuid().ToString('N');$choices=Join-Path (Get-LFRoot $Root) 'SourceChoices'
 [void][IO.Directory]::CreateDirectory($choices);$file=Join-Path $choices ($id+'.json')
 Write-LFJson $file @{state='pending'}
 try{$null=Start-Process (Join-Path $PSHOME 'powershell.exe') -WindowStyle Hidden -PassThru -ArgumentList ('-NoProfile -STA -ExecutionPolicy Bypass -File "'+(Join-Path $PSScriptRoot 'Select-SourcePowerPoint.ps1')+'" -ResultFile "'+$file+'"')}catch{Write-LFJson $file @{state='error';error=$_.Exception.Message};throw}
 @{choice=$id}
}
function Get-LFSourceChoice($Root,[string]$Choice){
 if($Choice -notmatch '^[a-f0-9]{32}$'){throw 'Invalid source choice.'}
 $file=Join-Path (Get-LFRoot $Root) ('SourceChoices/'+$Choice+'.json')
 if(-not (Test-Path $file) -or (Get-Item $file).LastWriteTimeUtc -lt [DateTime]::UtcNow.AddHours(-1)){throw 'Source selection expired. Choose the file again.'}
 Read-LFSharedJson $file
}
Export-ModuleMember -Function Get-LFSourceComparison,Invoke-LFSourceRefresh,Open-LFProjectLocation,Start-LFSourceChoice,Get-LFSourceChoice,Export-LFSourceThumbnails,Get-LFSourceDescription
