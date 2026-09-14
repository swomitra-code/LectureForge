param([string]$RuntimeRoot=(Join-Path $PSScriptRoot '../../work/script-import-tests'))
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '../../studio/Narration.psm1') -Force
$root=Join-Path $RuntimeRoot ([guid]::NewGuid().ToString('N').Substring(0,8))
[void][IO.Directory]::CreateDirectory($root)
function Check($Name,$Value){if(-not $Value){throw "FAIL: $Name"};"PASS: $Name"}
# Minimal package fixture for the production importer; no Office or external assets.
Add-Type -AssemblyName System.IO.Compression.FileSystem
$source=Join-Path $root 'input.pptx';$zip=[IO.Compression.ZipFile]::Open($source,'Create')
function Part($Name,$Text){$w=[IO.StreamWriter]::new($zip.CreateEntry($Name).Open());try{$w.Write($Text)}finally{$w.Dispose()}}
try{
 $ids='';$rels=''
 foreach($n in 1..20){$ids+='<p:sldId id="'+(255+$n)+'" r:id="rId'+$n+'"/>';$rels+='<Relationship Id="rId'+$n+'" Target="slides/slide'+$n+'.xml"/>';Part "ppt/slides/slide$n.xml" "<slide><t>Title $n</t></slide>"}
 Part 'ppt/presentation.xml' ('<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst>'+$ids+'</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>')
 Part 'ppt/_rels/presentation.xml.rels' ('<Relationships>'+$rels+'</Relationships>')
}finally{$zip.Dispose()}
$p=New-LFProject $root 'Import regression' 'input.pptx' ([IO.File]::ReadAllBytes($source));$id=$p.id;$dir=Get-LFProjectPath $root $id
function Wire {Get-LFNarration $root $id|ConvertTo-Json -Depth 30|ConvertFrom-Json}
function ImportText($Text){$script:p=Update-LFProject $root $id ([pscustomobject]@{version=$script:p.version;script_import=$Text})}
$minimal="Slide 1:`nHere is the question for this lesson.`n`nSlide 2:`nNow let's make the puzzle harder."
Check 'New lecture initializes every narration slide with an array' (@((Wire).slides|Where-Object {$_.takes -isnot [array] -or $_.takes.Count}).Count -eq 0)
ImportText $minimal
Check 'New lecture imports exact minimal scripts' ($p.slides[0].script -ceq 'Here is the question for this lesson.' -and $p.slides[1].script -ceq "Now let's make the puzzle harder.")
Check 'Import with no takes returns arrays and creates no narration file' ((Wire).slides[0].takes -is [array] -and -not (Test-Path (Join-Path $dir 'narration.json')))
$p=Read-LFProject $root $id;ImportText $minimal
Check 'Reopened lecture import preserves slide content' ($p.slides.Count -eq 20 -and $p.slides[19].title -eq 'Title 20' -and $p.slides[19].script -eq '')
$take=[pscustomobject]@{number=2;state='ready';asset=[pscustomobject]@{path='n/aaaaaaaaaaaa/bbbbbbbbbbbb/take.mp3';sha256='preserved';duration_seconds=1.5;post_speech_silence_seconds=0};attempts=@([pscustomobject]@{id='original';state='complete';submitted_utc=$null});custom='preserve me'}
$take1=$take|ConvertTo-Json -Depth 20|ConvertFrom-Json;$take1.number=1
$take3=$take|ConvertTo-Json -Depth 20|ConvertFrom-Json;$take3.number=3
$cases=@(
 @{name='missing';value=$null;count=0},
 @{name='null';value=$null;count=0},@{name='empty object';value=[pscustomobject]@{};count=0},
 @{name='single take';value=$take;count=1},@{name='keyed legacy takes';value=[pscustomobject]@{'2'=$take};count=1},
 @{name='existing array';value=@($take);count=1},@{name='malformed scalar';value='legacy';count=0},
 @{name='three existing takes';value=@($take,$take1,$take3);count=3},
 @{name='mixed array';value=@($null,'legacy',$take);count=1}
)
foreach($case in $cases){
 $r=[pscustomobject]@{id='aaaaaaaaaaaa';slide_number=1;binding=Get-LFBinding $p $p.slides[0];selected_take=2;takes=$case.value}
 if($case.name -eq 'missing'){$r.PSObject.Properties.Remove('takes')}
 $file=Join-Path $dir 'narration.json';Write-LFJson $file @{schema='lectureforge-narration-1';revisions=@($r)};$before=(Get-FileHash $file).Hash
 ImportText $minimal;$wire=Wire;$actual=$wire.slides[0].takes
 Check ($case.name+' normalizes on JSON boundary') ($actual -is [array] -and $actual.Count -eq $case.count)
 if($case.count){Check ($case.name+' retains exact take and selection') (($actual[0]|ConvertTo-Json -Depth 20 -Compress) -ceq ($take|ConvertTo-Json -Depth 20 -Compress) -and $wire.slides[0].selected_take -eq 2)}
 if($case.count -eq 3){Check 'All three existing takes preserved in their original order' (($actual|ConvertTo-Json -Depth 20 -Compress) -ceq ($case.value|ConvertTo-Json -Depth 20 -Compress))}
 Check ($case.name+' import leaves saved narration untouched') ((Get-FileHash $file).Hash -eq $before)
 if($case.name -eq 'malformed scalar'){Check 'Unrecognized legacy data retained for future writes' ((Read-LFNarrationFile $dir).revisions[0].takes_legacy -contains 'legacy')}
}
$historyHash=(Get-FileHash $file).Hash
$text=(1..20|ForEach-Object {"Slide ${_}:`nNarration for slide $_."}) -join "`n`n"
ImportText $text;$snapshot=$p.slides|ConvertTo-Json -Depth 20 -Compress
Check 'Twenty scripts bind to the correct slide numbers' (@($p.slides|Where-Object {$_.script -ceq "Narration for slide $($_.number)."}).Count -eq 20)
ImportText $text
Check 'Repeated import preserves all slide data and take history' (($p.slides|ConvertTo-Json -Depth 20 -Compress) -ceq $snapshot -and (Get-FileHash $file).Hash -eq $historyHash)
ImportText $minimal
Check 'Restoring script restores prior take binding' ((Wire).slides[0].takes[0].custom -eq 'preserve me')
Check 'Import never queues provider work' ((Wire).provider_calls -eq 0 -and (Wire).heygen_calls -eq 0 -and @(Get-ChildItem $dir -File).Count -eq 3)
# Sparse legacy collections must select by take number, never array position.
$state=Read-LFNarrationFile $dir;$t=$state.revisions[0].takes[0]
$asset=Resolve-LFNarrationAsset $dir $t.asset.path;[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($asset));[IO.File]::WriteAllText($asset,'local hash fixture')
$t.asset.sha256=(Get-FileHash $asset).Hash;Write-LFJson $file $state
Set-LFNarrationSelection $root $id 1 'aaaaaaaaaaaa' 0
Set-LFNarrationSelection $root $id 1 'aaaaaaaaaaaa' 2
Check 'Legacy singleton take 2 remains selectable by identity' ((Wire).slides[0].selected_take -eq 2)
$saved=Read-LFNarrationFile $dir
Check 'Later narration writes preserve valid takes and unrecognized legacy data' ($saved.revisions[0].takes[0].custom -eq 'preserve me' -and $saved.revisions[0].takes_legacy -contains 'legacy')
# Pass the fixture to the HTTP/UI regression, using the real production server.
Write-LFJson (Join-Path $root 'fixture.json') @{root=$root;id=$id;source=$source}
"FIXTURE=$(Join-Path $root 'fixture.json')"
& node (Join-Path $PSScriptRoot 'Test-ScriptImport.cjs') (Join-Path $root 'fixture.json')
if($LASTEXITCODE){throw 'Script import HTTP/UI regression failed.'}
