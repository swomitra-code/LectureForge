param([string]$RuntimeRoot=(Join-Path $env:LOCALAPPDATA 'LectureForgeE'))
$ErrorActionPreference='Stop'
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Import-Module (Join-Path $repo 'studio/Assembly.psm1') -Force
Import-Module (Join-Path $repo 'studio/Recording.psm1') -Force
$root=Get-LFRoot $RuntimeRoot
$test=Read-LFSharedJson (Join-Path $root 'results.json')
$dir=$test.project_root;$id=$test.project_id
$checks=@()
function Check($Name,[bool]$OK){
 if(-not $OK){throw "FAIL: $Name"}
 $script:checks+=@{name=$Name;passed=$true}
 Write-LFJson (Join-Path $root 'supplemental-results.json') @{checks=$checks;elevenlabs_calls=0;heygen_calls=0}
 "PASS: $Name"
}
$ready=Get-LFAssemblyReadiness $root $id
$state=Read-LFAssembly $dir
$beforeHash=(Get-FileHash (Join-Path $dir 'assembly.json')).Hash
$same=Request-LFAssembly $root $id $ready.binding $state.current.folder
Check 'Repeated Create after completion returns the same validated recording' ($same.current.id -eq $state.current.id -and (Get-FileHash (Join-Path $dir 'assembly.json')).Hash -eq $beforeHash)
$bad=$false
try { $null=Request-LFAssembly $root $id ('0'*64) $state.current.folder } catch { $bad=$true }
Check 'Stale browser assembly binding is rejected' $bad
$p=Read-LFProject $root $id
$bad=$false
try { Assert-LFPlacement @{left=0;top=0;width=1;height=1} $p } catch { $bad=$true }
Check 'Full-slide avatar placement is rejected' $bad
$bad=$false
try { Set-LFAssemblyPlacement $root $id 6 $null -1 } catch { $bad=$true }
Check 'Stale-tab placement update is rejected' $bad
$small=Join-Path $root ('Subset-'+[guid]::NewGuid().ToString('N').Substring(0,8))
[void][IO.Directory]::CreateDirectory($small)
$avatar=@($ready.snapshot.slides | Where-Object slide -eq 6)
$v=New-V2RecordingBatch -Source (Join-Path $dir 'source.pptx') -SourceSha256 $p.source.sha256 -Output (Join-Path $small 'deck.pptx') -EvidenceDirectory (Join-Path $small 'v') -Avatars $avatar -Progress {param($stage,$slide) Write-Host "$stage slide=$slide"}
Check 'Subset deck preserves 13 disabled/unproduced slides' ($v.passed -and $v.avatar_count -eq 1 -and @($v.slides|Where-Object {-not $_.produced -and $_.native_preserved -and $_.native_render_identical}).Count -eq 13)
Check 'Subset insertion retains the requested single-slide placement' ($v.slides[5].produced -and $v.slides[5].placement.left -eq .84)
Check 'Subset test leaves complete Recording output and state untouched' ((Get-FileHash (Join-Path $dir 'assembly.json')).Hash -eq $beforeHash -and (Get-FileHash $state.current.output).Hash -eq $state.current.output_sha256)
Check 'No provider calls from subset assembly or readiness reads' ((Get-LFNarration $root $id).provider_calls -eq 0 -and (Get-LFAvatarStatus $root $id).provider_calls -eq 0)
