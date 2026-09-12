$ErrorActionPreference = "Continue"
Write-Host "LectureForge prerequisite check" -ForegroundColor Cyan

function Check-Command($name) {
  $cmd = Get-Command $name -ErrorAction SilentlyContinue
  if ($cmd) { Write-Host ("[OK] {0}: {1}" -f $name, $cmd.Source) -ForegroundColor Green }
  else { Write-Host ("[MISSING] {0}" -f $name) -ForegroundColor Yellow }
}

Check-Command "codex"
Check-Command "git"
Check-Command "python"
Check-Command "ffmpeg"

try {
  $ppt = New-Object -ComObject PowerPoint.Application
  Write-Host ("[OK] PowerPoint COM, version {0}" -f $ppt.Version) -ForegroundColor Green
  $ppt.Quit()
} catch {
  Write-Host ("[MISSING/ERROR] PowerPoint COM: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Next: put the source PPTX in .\source and the approved Slide 10 WebM in .\media, then run codex from this folder." -ForegroundColor Cyan
