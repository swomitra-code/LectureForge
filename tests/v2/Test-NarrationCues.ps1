param([string]$RuntimeRoot=(Join-Path $PSScriptRoot '../../work/narration-cue-tests'))
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '../../studio/Narration.psm1') -Force
$provider=Import-Module (Join-Path $PSScriptRoot '../../studio/NarrationProvider.psm1') -Force -PassThru
$root=Join-Path $RuntimeRoot ([guid]::NewGuid().ToString('N').Substring(0,8))
[void][IO.Directory]::CreateDirectory($root)
function Check($Name,$OK){if(-not $OK){throw "FAIL: $Name"};"PASS: $Name"}
# Intercept the real provider transport, including its serialized UTF-8 body.
# No network requests or paid generations are made by this test.
& $provider {
 function script:Invoke-WebRequest {
  param([switch]$UseBasicParsing,$TimeoutSec,$Method,$Uri,$Headers,$ContentType,$Body,$OutFile)
  $script:captured=[Text.Encoding]::UTF8.GetString($Body)|ConvertFrom-Json
  $script:requests++
 }
 $script:requests=0
}
$settings=[pscustomobject]@{voice_id='fixture';model_id='fixture-model';stability=0.4;similarity_boost=0.8;style=0;use_speaker_boost=$true;speed=1}
$cases=@(
 @{name='All supported directives';input='[POINT: collector][TRACE: pipe][CIRCLE: tank][SWEEP: left to right][REVEAL][PAUSE 2 seconds][PAUSE 1.5s, REVEAL]Heat flows.';expected='       Heat flows.'},
 @{name='Ordinary brackets and cue-name prefixes';input='[Figure 2] [kWh] [POINTS: three] [TRACEABILITY: source] [REVEAL more] [POINT of interest]';expected='[Figure 2] [kWh] [POINTS: three] [TRACEABILITY: source] [REVEAL more] [POINT of interest]'},
 @{name='No cues stays byte-for-byte unchanged';input="  Heat flows.`r`nKeep  spacing and punctuation!  ";expected="  Heat flows.`r`nKeep  spacing and punctuation!  "},
 @{name='Consecutive cues do not join words';input='Heat[POINT: collector][REVEAL]flows[PAUSE 2]here.';expected='Heat  flows here.'},
 @{name='Mixed brackets and case';input="[point: collector]Read [kWh].`n[ pause 1, REVEAL ]Next [source: 2024].";expected=" Read [kWh].`n Next [source: 2024]."},
 @{name='Unsupported and incomplete directives preserved';input='[WAIT 2] [PAUSED] [POINT: unfinished';expected='[WAIT 2] [PAUSED] [POINT: unfinished'}
)
$priorKey=$env:ELEVENLABS_API_KEY
try{
 $env:ELEVENLABS_API_KEY='offline-fixture'
 $n=0
 foreach($case in $cases){
  $n++;$output=Join-Path $root "$n.mp3";$original=$case.input
  Invoke-LFElevenLabs $settings $case.input $output
  $payload=& $provider {$script:captured}
  Check ($case.name+' in ElevenLabs payload') ($payload.text -ceq $case.expected)
  Check ($case.name+' exact spoken-text log') ([IO.File]::ReadAllText($output+'.spoken.txt') -ceq $payload.text)
  Check ($case.name+' source intact') ($case.input -ceq $original)
  Check ($case.name+' provider settings intact') ($payload.model_id -eq $settings.model_id -and $payload.voice_settings.speed -eq 1 -and $payload.voice_settings.stability -eq 0.4)
 }
 # Exercise the real importer, persistence, and narration revision snapshot.
 Add-Type -AssemblyName System.IO.Compression.FileSystem
 $source=Join-Path $root 'input.pptx';$zip=[IO.Compression.ZipFile]::Open($source,'Create')
 function Part($Name,$Text){$w=[IO.StreamWriter]::new($zip.CreateEntry($Name).Open());try{$w.Write($Text)}finally{$w.Dispose()}}
 try{
  Part 'ppt/presentation.xml' '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>'
  Part 'ppt/_rels/presentation.xml.rels' '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/></Relationships>'
  Part 'ppt/slides/slide1.xml' '<slide><t>Collector</t></slide>'
 }finally{$zip.Dispose()}
 $p=New-LFProject $root 'Cue regression' 'input.pptx' ([IO.File]::ReadAllBytes($source))
 $imported="Look at the [POINT: collector]collector.`n[PAUSE 2, REVEAL]Read [Figure 2]."
 $beforeCalls=& $provider {$script:requests}
 $p=Update-LFProject $root $p.id ([pscustomobject]@{version=$p.version;script_import="Slide 1:`n$imported"})
 $dir=Get-LFProjectPath $root $p.id
 Check 'Import preserves exact original script including cues' ($p.slides[0].script -ceq $imported)
 Check 'Import is provider-free and creates no narration history' ((& $provider {$script:requests}) -eq $beforeCalls -and -not (Test-Path (Join-Path $dir 'narration.json')))
 Add-LFNarrationBatch $root $p.id $p.version
 $r=(Read-LFNarrationFile $dir).revisions[0]
 $projectHash=(Get-FileHash (Join-Path $dir 'project.json')).Hash
 $historyHash=(Get-FileHash (Join-Path $dir 'narration.json')).Hash
 Invoke-LFElevenLabs $r.settings $r.script (Join-Path $root 'imported.mp3')
 Check 'Imported narration payload is sanitized' ((& $provider {$script:captured.text}) -ceq "Look at the  collector.`n Read [Figure 2].")
 Check 'Saved project and revision remain byte-for-byte intact after send' ((Get-FileHash (Join-Path $dir 'project.json')).Hash -eq $projectHash -and (Get-FileHash (Join-Path $dir 'narration.json')).Hash -eq $historyHash)
 Check 'Reopened script, revision, and script hash retain cues' ((Read-LFProject $root $p.id).slides[0].script -ceq $imported -and $r.script -ceq $imported -and $r.script_sha256 -eq (Get-LFTextHash $imported))
 Check 'Three takes and selection behavior unchanged' ($r.takes.Count -eq 3 -and $null -eq $r.selected_take -and @($r.takes|Where-Object state -eq 'queued').Count -eq 3)
 'PASS: narration cue regression; real ElevenLabs calls 0; HeyGen calls 0.'
}finally{
 $env:ELEVENLABS_API_KEY=$priorKey
 Remove-Module $provider
}
