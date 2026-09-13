param([ValidateSet('prepare','import','capture','verify','recover-local','recover-assembly')][string]$Mode='prepare',[string]$RuntimeRoot=(Join-Path $env:LOCALAPPDATA 'LectureForgeF'),[string]$Id)
# Test setup/evidence only. Never part of the instructor workflow; no provider transport.
$ErrorActionPreference='Stop'
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Import-Module (Join-Path $repo 'studio/Assembly.psm1') -Force
$prior=Join-Path $env:LOCALAPPDATA 'LectureForgeEF/Projects/9a6884f7a78e'
$root=Get-LFRoot $RuntimeRoot
if($Mode -eq 'prepare'){
 if(Test-Path $root){throw 'Use a fresh short fixture root; existing evidence is preserved.'}
 [void][IO.Directory]::CreateDirectory($root)
 Write-LFJson (Join-Path $root 'avatar-policy.json') @{provider_calls_enabled=$false;purpose='F offline acceptance'}
 Copy-Item -LiteralPath (Join-Path $prior 'source.pptx') -Destination (Join-Path $root 'input.pptx')
 $p=Read-LFSharedJson (Join-Path $prior 'project.json')
 $scripts=($p.slides|ForEach-Object {"## Slide $($_.number)`r`n$($_.script)`r`n"}) -join "`r`n"
 [IO.File]::WriteAllText((Join-Path $root 'scripts.md'),$scripts,[Text.UTF8Encoding]::new($false))
 Write-Output 'Prepared source/scripts COPY; provider transport disabled.'
 exit
}
$dir=Get-LFProjectPath $root $Id
if($Mode -eq 'import'){
 if(Test-Path (Join-Path $dir 'narration.json')){throw 'Fixture assets already imported; never overwrite.'}
 $p=Read-LFProject $root $Id
 $n=Read-LFNarrationFile $prior
 $current=@(foreach($s in $p.slides){$r=Get-LFCurrentRevision $n $p $s;if(-not $r){throw "Script/pause/preset fixture mismatch: Slide $($s.number)"};$r})
 $expected=@{}
 foreach($r in $current){
  $expected["$($r.slide_number)"]=$r.selected_take
  foreach($t in $r.takes){if(-not(Test-LFTake $prior $t)){throw 'Unverified fixture take'}}
  $r.slide_id="$Id/$($r.slide_number)";$r.script_revision=$p.version;$r.selected_take=$null
 }
 Copy-Item -LiteralPath (Join-Path $prior 'n') -Destination (Join-Path $dir 'n') -Recurse
 Copy-Item -LiteralPath (Join-Path $prior 'avatars') -Destination (Join-Path $dir 'avatars') -Recurse
 Copy-Item -LiteralPath (Join-Path $prior 'avatars.json') -Destination (Join-Path $dir 'avatars.json')
 Write-LFJson (Join-Path $dir 'narration.json') @{schema='lectureforge-narration-1';revisions=$current;intent_revision=0}
 $assets=@(foreach($r in $current){foreach($t in $r.takes){@{path=$t.asset.path;sha256=$t.asset.sha256}}})
 $catalog=Read-LFSharedJson (Join-Path $dir 'avatars.json')
 foreach($a in $catalog.assets){$assets+=@{path=$a.path;sha256=$a.sha256}}
 Write-LFJson (Join-Path $root 'fixture.json') @{id=$Id;expected=$expected;assets=$assets;source_sha256=$p.source.sha256;prior=$prior}
 Write-Output 'Imported 42 existing takes and approved/historical avatars; no selections or authorizations created.'
 exit
}
$fixture=Read-LFSharedJson (Join-Path $root 'fixture.json')
if($Mode -in 'recover-local','recover-assembly'){
 # Deterministic release recovery checks alter only isolated fixture metadata.
 # The worker must be stopped; no media or provider records are rewritten.
 if($fixture.id -ne $Id -or (Read-LFSharedJson (Join-Path $root 'avatar-policy.json')).provider_calls_enabled -ne $false){throw 'Offline fixture required'}
 if((Read-LFSharedJson (Join-Path $root 'runtime.json')).status -ne 'stopped'){throw 'Stop fixture before fault injection'}
 if($Mode -eq 'recover-local'){
  $status=Get-LFAvatarStatus $root $Id
  if($status.ready -ne 14){throw 'Expected all existing avatars ready'}
  $j=Read-LFAvatarJob $dir ($status.jobs|Where-Object slide -eq 5).id
  if($j.lease -or -not $j.row.reusable_avatar){throw 'Unleased reusable fixture required'}
  $j.stage='Exception';$j.error='Narration state busy. Try again.';$j.exception_kind='local';$j.output_path=$null;$j.output_sha256=$null
  Save-LFAvatarJob $dir $j
 }else{
  $ready=Get-LFAssemblyReadiness $root $Id
  if(-not $ready.ready -or (Read-LFAssembly $dir).current){throw 'Fresh ready assembly fixture required'}
  $null=Request-LFAssembly $root $Id $ready.binding (Join-Path $root 'Output')
  $a=Read-LFAssembly $dir;$a.current.stage='Inserting avatars';$a.current.slide=4;$a.current.lease=@{pid=2147483647;started='0'}
  Write-LFJson (Join-Path $dir 'assembly.json') $a
  $null=Start-LFAssemblyTask $root
  if((Read-LFAssembly $dir).current.stage -ne 'Exception'){throw 'Interrupted assembly not detected'}
 }
 Write-Output 'Prepared isolated local recovery exception; assets unchanged and no provider calls.'
 exit
}
$state=Get-LFAssemblyStatus $root $Id
$ready=Get-LFAssemblyReadiness $root $Id
if(-not $ready.ready -or $state.current.stage -ne 'Complete'){throw 'Recording not ready'}
$v=Read-LFSharedJson (Join-Path $dir "r/$($state.current.id)/v/validation.json")
if(-not $v.passed -or $v.slide_count -ne 14 -or $v.avatar_count -ne 14 -or $v.save_calls -ne 1){throw 'Full deck validation failed'}
foreach($a in $fixture.assets){if((Get-FileHash (Join-Path $dir $a.path)).Hash -ne $a.sha256 -or (Get-FileHash (Join-Path $prior $a.path)).Hash -ne $a.sha256){throw 'Existing paid media changed'}}
if((Get-FileHash (Join-Path $repo 'source/RE-M2-Solar-Thermal-Overview.pptx')).Hash -ne $fixture.source_sha256){throw 'Canonical source changed'}
if((Get-LFNarration $root $Id).provider_calls -ne 0 -or (Get-LFAvatarStatus $root $Id).provider_calls -ne 0){throw 'Provider boundary violated'}
$hashes=@{}
foreach($name in @('project.json','narration.json','assembly.json','avatars.json')){$hashes[$name]=(Get-FileHash (Join-Path $dir $name)).Hash}
foreach($f in Get-ChildItem (Join-Path $dir 'approvals') -File){$hashes['approvals/'+$f.Name]=(Get-FileHash $f.FullName).Hash}
foreach($jid in (Read-LFAvatarQueue $dir).jobs){$hashes['job/'+$jid]=(Get-FileHash (Get-LFAvatarJobPath $dir $jid)).Hash}
$hashes['recording']=(Get-FileHash $state.current.output).Hash
if($Mode -eq 'capture'){Write-LFJson (Join-Path $root 'before-restart.json') $hashes}else{
 $before=Read-LFSharedJson (Join-Path $root 'before-restart.json')
 foreach($name in $hashes.Keys){if($before.$name -ne $hashes[$name]){throw "Restart changed $name"}}
 Write-LFJson (Join-Path $root 'verified.json') @{output=$state.current.output;source_sha256=$fixture.source_sha256;slides=14;avatars=14;elevenlabs_calls=0;heygen_calls=0;restart_identical=$true;validation=$v}
}
Write-Output 'PASS: native deck/media/authoritative bindings, source integrity, immutable fixture assets and zero calls.'
