param([string]$RuntimeRoot,[string]$Id,[string]$Revision,[int]$Take,[string]$Attempt)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Narration.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'NarrationProvider.psm1') -Force
$context=Invoke-LFNarrationLock $RuntimeRoot $Id {param($dir)
 $state=Read-LFNarrationFile $dir;$r=@($state.revisions|Where-Object id -eq $Revision)[0];$t=$r.takes[$Take-1];$a=$t.attempts[-1]
 if($a.id -ne $Attempt -or $t.state -ne 'dispatched'){throw 'Attempt is no longer dispatched.'}
 $a.owner_pid=$PID;$a.owner_start=(Get-Process -Id $PID).StartTime.ToUniversalTime().ToString('o')
 $a.state='calling';$t.state='generating';$a.submitted_utc=[DateTime]::UtcNow.ToString('o')
 Write-LFJson (Join-Path $dir 'narration.json') $state
 @{dir=$dir;revision=$r;attempt=$a}
}
$result=$null;$failureState='uncertain';$diagnostic=$null;$received=$false
try{
 $raw=Resolve-LFNarrationAsset $context.dir $context.attempt.raw_path
 [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($raw))
 Invoke-LFElevenLabs $context.revision.settings $context.revision.script $raw
 $received=$true
 $result=Complete-LFNarrationAudio $context.dir $context.revision $context.attempt
}catch{
 $diagnostic=Get-NarrationFailureDiagnostic $_ @($env:ELEVENLABS_API_KEY,$env:HEYGEN_API_KEY,$context.revision.script)
 if($received){$failureState='local exception'}else{
  $status=0;try{$status=[int]$_.Exception.Response.StatusCode}catch{}
  if(($status -ge 400 -and $status -lt 500 -and $status -ne 408) -or -not $env:ELEVENLABS_API_KEY){$failureState='failed'}
 }
}
Invoke-LFNarrationLock $RuntimeRoot $Id {param($dir)
 $state=Read-LFNarrationFile $dir;$r=@($state.revisions|Where-Object id -eq $Revision)[0];$t=$r.takes[$Take-1];$a=$t.attempts[-1]
 if($a.id -ne $Attempt){throw 'Attempt identity changed; preserving files without changing state.'}
 $a.finished_utc=[DateTime]::UtcNow.ToString('o')
 if($result){$t.asset=$result;$t.state='ready';$a.state='complete'}else{$t.state=$failureState;$a.state=$failureState;$a.error=$diagnostic}
 Write-LFJson (Join-Path $dir 'narration.json') $state
}
