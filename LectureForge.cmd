@echo off
setlocal
rem Run in the logged-in Windows session used by desktop PowerPoint.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Start-LectureForge.ps1" %*
if errorlevel 1 (
  echo LectureForge could not start. Review the message above.
  pause
  exit /b 1
)
endlocal
