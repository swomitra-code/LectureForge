param([string]$ProjectDirectory)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Production.psm1') -Force
Add-Type -AssemblyName System.Windows.Forms
$file=Join-Path $ProjectDirectory 'folder-choice.json'
$mutex=[Threading.Mutex]::new($false,('Local\LF-Folder-'+[IO.Path]::GetFileName($ProjectDirectory)));$held=$false
try{
 $held=$mutex.WaitOne(0);if(-not $held){exit}
 Write-LFJson $file @{state='choosing'}
 $dialog=New-Object Windows.Forms.FolderBrowserDialog;$dialog.Description='Choose the Recording PowerPoint output folder';$dialog.ShowNewFolderButton=$true
 try{$result=$dialog.ShowDialog();Write-LFJson $file @{state=if($result -eq 'OK'){'selected'}else{'cancelled'};path=$dialog.SelectedPath;timestamp=[DateTime]::UtcNow.ToString('o')}}finally{$dialog.Dispose()}
}finally{if($held){$mutex.ReleaseMutex()};$mutex.Dispose()}
