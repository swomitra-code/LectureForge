Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Project.psm1')
Import-Module (Join-Path $PSScriptRoot 'NarrationSettings.psm1')
function Get-LFRoot([string]$Root) {
 if(-not $Root){$Root=Join-Path $env:LOCALAPPDATA 'LectureForge'}
 $p=[IO.Path]::GetFullPath($Root).TrimEnd('\')
 if($p.Length -gt 100){throw 'Choose a shorter local runtime root (100 characters maximum).'}
 $p
}
function Write-LFJson([string]$Path,$Value){
 $tmp=$Path+'.tmp';$json=ConvertTo-Json -InputObject $Value -Depth 30
 [IO.File]::WriteAllText($tmp,$json,[Text.UTF8Encoding]::new($false))
 if([IO.File]::Exists($Path)){[IO.File]::Replace($tmp,$Path,[NullString]::Value)}else{[IO.File]::Move($tmp,$Path)}
}
function Read-LFSharedJson([string]$Path){
 for($attempt=0;$attempt -lt 20;$attempt++){
  try{
   $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
   try{$reader=[IO.StreamReader]::new($stream,[Text.Encoding]::UTF8);try{return ($reader.ReadToEnd()|ConvertFrom-Json)}finally{$reader.Dispose()}}finally{$stream.Dispose()}
  }catch [IO.IOException]{if($attempt -eq 19){throw};Start-Sleep -Milliseconds 25}
 }
}
function Get-LFProjectPath([string]$Root,[string]$Id){
 if($Id -notmatch '^[a-f0-9]{12}$'){throw 'Invalid project ID.'}
 Join-Path (Join-Path (Get-LFRoot $Root) 'Projects') $Id
}
function Restore-LFSourceTransaction([string]$Dir){
 $journal=Join-Path $Dir 'source-refresh-transaction.json'
 if(-not (Test-Path -LiteralPath $journal)){return}
 $tx=Read-LFSharedJson $journal
 if($tx.id -notmatch '^[a-f0-9]{32}$'){throw 'Invalid source transaction journal.'}
 $root=Split-Path (Split-Path $Dir)
 $backup=Join-Path $root ('SourceBackups/'+(Split-Path $Dir -Leaf)+'/'+$tx.id+'/project')
 $inventory=Get-Content -Raw -LiteralPath (Join-Path (Split-Path $backup) 'inventory.json')|ConvertFrom-Json
 # Restore only files owned by source refresh; never rewrite production assets.
 foreach($name in 'source.pptx','project.json','assembly.json','thumbs'){
  $dest=Join-Path $Dir $name;$src=Join-Path $backup $name
  foreach($item in @($inventory|Where-Object {$_.path -eq $name -or $_.path.StartsWith($name+'\')})){
   if((Get-FileHash -LiteralPath (Join-Path $backup $item.path)).Hash -ne $item.sha256){throw 'Source rollback backup failed verification.'}
  }
  if(Test-Path -LiteralPath $src){
   if($name -eq 'thumbs'){
    if([IO.Path]::GetFullPath($dest) -ne ([IO.Path]::GetFullPath($Dir).TrimEnd('\')+'\thumbs')){throw 'Unsafe thumbnail rollback path.'}
    if(Test-Path -LiteralPath $dest){Remove-Item -LiteralPath $dest -Recurse -Force}
    Copy-Item -LiteralPath $src -Destination $dest -Recurse
   }else{[IO.File]::Copy($src,$dest,$true)}
  }elseif(Test-Path -LiteralPath $dest){Remove-Item -LiteralPath $dest -Recurse -Force}
 }
 Remove-Item -LiteralPath $journal
}
function Invoke-LFProjectLock([string]$Root,[string]$Id,[scriptblock]$Action,[int]$Timeout=60000){
 $dir=Get-LFProjectPath $Root $Id
 $hash=[Security.Cryptography.SHA256]::Create()
 try{$key=([BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($dir)))).Replace('-','').Substring(0,24)}finally{$hash.Dispose()}
 # Same reentrant mutex as the existing narration/assembly/provider interfaces.
 $mutex=[Threading.Mutex]::new($false,('Local\LFN-'+$key));$held=$false
 try{
  try{$held=$mutex.WaitOne($Timeout)}catch [Threading.AbandonedMutexException]{$held=$true}
  if(-not $held){throw 'Project work is active. Wait for it to finish, then retry.'}
  Restore-LFSourceTransaction $dir
  & $Action $dir
 }finally{if($held){$mutex.ReleaseMutex()};$mutex.Dispose()}
}
function Read-LFProject([string]$Root,[string]$Id){
 Invoke-LFProjectLock $Root $Id {param($dir) Read-LFProjectUnlocked $Root $Id}
}
function Read-LFProjectUnlocked([string]$Root,[string]$Id){
 $dir=Get-LFProjectPath $Root $Id
 $p=Get-Content -Encoding UTF8 -LiteralPath (Join-Path $dir 'project.json') -Raw | ConvertFrom-Json
 if($p.schema -ne 'lectureforge-studio-1' -or $p.id -ne $Id){throw 'Unsupported project state.'}
 if((Get-V2Asset (Join-Path $dir 'source.pptx')).sha256 -ne $p.source.sha256){throw 'Project source hash mismatch. Original source remains untouched.'}
 if($p.preset.PSObject.Properties.Name -notcontains 'narration'){
  $p.preset|Add-Member -NotePropertyName narration -NotePropertyValue (Get-LFNarrationDefaults)
 }elseif($null -eq $p.preset.narration){$p.preset.narration=Get-LFNarrationDefaults}
 $resolved=Resolve-LFNarrationSettings $p.preset.narration
 foreach($key in $resolved.PSObject.Properties.Name){
  if($p.preset.narration.PSObject.Properties.Name -notcontains $key){$p.preset.narration|Add-Member -NotePropertyName $key -NotePropertyValue $resolved.$key}
 }
 $p
}
function Get-LFPreset([string]$Id='renewable-energy'){
 if($Id -ne 'renewable-energy'){throw 'Unknown production preset.'}
 Get-Content -Encoding UTF8 -Raw (Join-Path $PSScriptRoot 'presets/renewable-energy.json') | ConvertFrom-Json
}
function New-LFProject([string]$Root,[string]$Name,[string]$OriginalName,[byte[]]$Bytes,[string]$Preset='renewable-energy',[string]$ExternalPath=''){
 if([string]::IsNullOrWhiteSpace($Name) -or $Name.Length -gt 120){throw 'Enter a project name (1â€“120 characters).'}
 $settings=Get-LFPreset $Preset
 if($ExternalPath){
  $ExternalPath=[IO.Path]::GetFullPath($ExternalPath)
  $h=[Security.Cryptography.SHA256]::Create();try{$uploadHash=([BitConverter]::ToString($h.ComputeHash($Bytes))).Replace('-','')}finally{$h.Dispose()}
  if([IO.Path]::GetExtension($ExternalPath) -ne '.pptx' -or (Get-V2Asset $ExternalPath).sha256 -ne $uploadHash){throw 'External PowerPoint changed during import. Choose it again.'}
 }
 $id=[guid]::NewGuid().ToString('N').Substring(0,12);$dir=Get-LFProjectPath $Root $id
 [void][IO.Directory]::CreateDirectory($dir)
 $source=Join-Path $dir 'source.pptx'
 [IO.File]::WriteAllBytes($source,$Bytes)
 $asset=Get-V2Asset $source;$info=Get-V2DeckInfo $source
 if($info.slide_count -lt 1 -or $info.slide_count -gt 500){throw 'Expected a PowerPoint with 1â€“500 slides.'}
 $zip=[IO.Compression.ZipFile]::OpenRead($source)
 try{
  function Read-Part($part){$entry=$zip.GetEntry($part);if(-not $entry){throw "Missing PPTX part: $part"};$r=[IO.StreamReader]::new($entry.Open());try{[xml]$r.ReadToEnd()}finally{$r.Dispose()}}
  $pres=Read-Part 'ppt/presentation.xml';$rels=Read-Part 'ppt/_rels/presentation.xml.rels';$slides=@();$n=0
  foreach($node in $pres.presentation.sldIdLst.sldId){
   $n++;$rid=$node.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
   $rel=@($rels.Relationships.Relationship|Where-Object Id -eq $rid)[0]
   if($rel.GetAttribute('TargetMode') -eq 'External'){throw 'External slide reference is unsupported.'}
   $uri=[Uri]::new([Uri]'http://package.invalid/ppt/presentation.xml',[string]$rel.Target)
   $sx=Read-Part $uri.AbsolutePath.TrimStart('/')
   $text=@($sx.SelectNodes("//*[local-name()='t']")|ForEach-Object InnerText)
   $title=if($text.Count){$text[0]}else{"Slide $n"}
   $slides+=@{number=$n;title=$title;enabled=$true;script='';post_speech_silence_seconds=0;placement_override=$null;status='awaiting script'}
  }
 }finally{$zip.Dispose()}
 $p=@{schema='lectureforge-studio-1';id=$id;name=$Name;version=1;created_utc=[DateTime]::UtcNow.ToString('o');source=@{path='source.pptx';original_name=[IO.Path]::GetFileName($OriginalName);sha256=$asset.sha256;slide_count=$info.slide_count;width_emu=$info.width_emu;height_emu=$info.height_emu};preset=$settings;slides=$slides;recording=$null;provider_calls=0;stage='project setup';assembly_policy='explicit final stage after all enabled avatars are ready'}
 if($ExternalPath){$p.source.external_path=$ExternalPath}
 Write-LFJson (Join-Path $dir 'project.json') $p
 Read-LFProject $Root $id
}
function ConvertFrom-LFScripts([string]$Text,[int]$SlideCount){
 $matches=[regex]::Matches($Text,'(?im)^[ \t]*(?:#{1,6}[ \t]*)?Slide[ \t]+(\d+)[ \t]*:?[^\r\n]*\r?\n')
 if(-not $matches.Count){throw 'Use headings such as ## Slide 1 followed by narration text.'}
 $result=@{}
 for($i=0;$i -lt $matches.Count;$i++){
  $n=[int]$matches[$i].Groups[1].Value
  if($n -lt 1 -or $n -gt $SlideCount -or $result.ContainsKey([string]$n)){throw "Duplicate or out-of-range Slide $n."}
  $start=$matches[$i].Index+$matches[$i].Length;$end=if($i+1 -lt $matches.Count){$matches[$i+1].Index}else{$Text.Length}
  $result[[string]$n]=$Text.Substring($start,$end-$start).TrimEnd([char[]]"`r`n")
 }
 $result
}
function Update-LFProject([string]$Root,[string]$Id,$Update){
 $p=Read-LFProject $Root $Id
 if([int]$Update.version -ne $p.version){throw 'Project changed in another window. Reload before saving.'}
 if($Update.PSObject.Properties.Name -contains 'narration_settings'){
  if($null -eq $Update.narration_settings){throw 'Narration settings must be an object.'}
  $p.preset.narration=Resolve-LFNarrationSettings $Update.narration_settings
 }
 if($Update.PSObject.Properties.Name -contains 'script_import'){
  $scripts=ConvertFrom-LFScripts $Update.script_import $p.source.slide_count
  foreach($s in $p.slides){if($scripts.ContainsKey([string]$s.number)){$s.script=$scripts[[string]$s.number]}}
 }else{
  if(@($Update.slides).Count -ne $p.slides.Count){throw 'Slide count mismatch.'}
  for($i=0;$i -lt $p.slides.Count;$i++){
   $s=$Update.slides[$i];if([int]$s.number -ne $p.slides[$i].number){throw 'Slide order mismatch.'}
   $pause=[double]$s.post_speech_silence_seconds
   if([double]::IsNaN($pause) -or [double]::IsInfinity($pause) -or $pause -lt 0 -or $pause -gt 120){throw 'Pause must be between 0 and 120 seconds.'}
   $p.slides[$i].script=[string]$s.script;$p.slides[$i].enabled=[bool]$s.enabled;$p.slides[$i].post_speech_silence_seconds=$pause
  }
 }
 foreach($s in $p.slides){$s.status=if(-not $s.enabled){'disabled'}elseif([string]::IsNullOrWhiteSpace($s.script)){'awaiting script'}else{'script ready'}}
 $p.version++;Write-LFJson (Join-Path (Get-LFProjectPath $Root $Id) 'project.json') $p;$p
}
function Get-LFProjects([string]$Root){
 $dir=Join-Path (Get-LFRoot $Root) 'Projects'
 if(Test-Path $dir){foreach($f in Get-ChildItem -LiteralPath $dir -Directory){if($f.Name -match '^[a-f0-9]{12}$' -and (Test-Path (Join-Path $f.FullName 'project.json'))){$p=Get-Content -Encoding UTF8 -Raw (Join-Path $f.FullName 'project.json')|ConvertFrom-Json;[pscustomobject]@{id=$p.id;name=$p.name;slides=$p.source.slide_count}}}}
}
Export-ModuleMember -Function Read-LFSharedJson,Get-LFRoot,Write-LFJson,Get-LFProjectPath,Read-LFProject,Get-LFPreset,New-LFProject,Update-LFProject,Get-LFProjects,ConvertFrom-LFScripts,Invoke-LFProjectLock,Restore-LFSourceTransaction
