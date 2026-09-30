param([string]$RuntimeRoot,[string]$Id)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Production.psm1')
Import-Module (Join-Path $PSScriptRoot 'SourceRefresh.psm1') -Force
Invoke-LFProjectLock $RuntimeRoot $Id {
 param($dir)
 $p=Read-LFProject $RuntimeRoot $Id;$thumbs=Join-Path $dir 'thumbs'
 [void][IO.Directory]::CreateDirectory($thumbs)
 try{
  Write-LFJson (Join-Path $thumbs 'status.json') @{state='exporting';pid=$PID}
  Export-LFSourceThumbnails (Join-Path $dir 'source.pptx') $thumbs $p.source.slide_count $p.source.width_emu $p.source.height_emu
  if((Get-FileHash -LiteralPath (Join-Path $dir 'source.pptx')).Hash -ne $p.source.sha256){throw 'Source changed during thumbnail export.'}
  Write-LFJson (Join-Path $thumbs 'status.json') @{state='ready';source_sha256=$p.source.sha256}
 }catch{
  Write-LFJson (Join-Path $thumbs 'status.json') @{state='exception';error=$_.Exception.Message}
  throw
 }
}
