param([Parameter(Mandatory=$true)][string]$ResultFile)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Production.psm1')
Add-Type -AssemblyName System.Windows.Forms
$dialog=$null;$owner=$null
try{
 $dialog=[Windows.Forms.OpenFileDialog]::new();$dialog.Filter='PowerPoint presentation (*.pptx)|*.pptx';$dialog.Title='Choose source PowerPoint';$dialog.Multiselect=$false;$dialog.CheckFileExists=$true
 $owner=[Windows.Forms.Form]::new();$owner.TopMost=$true;$owner.ShowInTaskbar=$false;$owner.Opacity=0
 $owner.Show();$owner.Activate()
 if($dialog.ShowDialog($owner) -eq [Windows.Forms.DialogResult]::OK){Write-LFJson $ResultFile @{state='selected';path=[IO.Path]::GetFullPath($dialog.FileName)}}else{Write-LFJson $ResultFile @{state='cancelled'}}
}catch{Write-LFJson $ResultFile @{state='error';error=$_.Exception.Message}}
finally{if($dialog){$dialog.Dispose()};if($owner){$owner.Dispose()}}
