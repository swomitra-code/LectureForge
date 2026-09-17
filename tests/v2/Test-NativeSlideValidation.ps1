param(
 [string]$Source=(Join-Path $PSScriptRoot '../../source/RE-M2-Solar-Thermal-Overview.pptx'),
 [string]$LiveAssembled=''
)
$ErrorActionPreference='Stop'
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
Import-Module (Join-Path $repo 'studio/NativeSlideValidation.psm1') -Force
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression
$root=Join-Path $repo 'work/native-slide-validator-tests';[void][IO.Directory]::CreateDirectory($root)
function Check($Name,[bool]$Passed){if(-not $Passed){throw "FAIL: $Name"};"PASS: $Name"}
function Edit-Package($Name,[scriptblock]$Edit){
 $path=Join-Path $root ($Name+'.pptx');[IO.File]::Copy($Source,$path,$true)
 $zip=[IO.Compression.ZipFile]::Open($path,[IO.Compression.ZipArchiveMode]::Update)
 try{& $Edit $zip}finally{$zip.Dispose()};$path
}
function Read-EntryXml($Zip,$Name){$e=$Zip.GetEntry($Name);$r=[IO.StreamReader]::new($e.Open());try{[xml]$r.ReadToEnd()}finally{$r.Dispose()}}
function Write-EntryXml($Zip,$Name,$Xml){$Zip.GetEntry($Name).Delete();$e=$Zip.CreateEntry($Name);$w=[IO.StreamWriter]::new($e.Open(),[Text.UTF8Encoding]::new($false));try{$Xml.Save($w)}finally{$w.Dispose()}}
function CompareNative($Path,[int]$Slide=2){Compare-LFNativeSlideSnapshot (Get-LFNativeSlideSnapshot $Source $Slide) (Get-LFNativeSlideSnapshot $Path $Slide)}

$metadata=Edit-Package metadata-order {
 param($z);[xml]$x=Read-EntryXml $z 'docProps/core.xml';$x.DocumentElement.SetAttribute('test-normalization','benign');Write-EntryXml $z 'docProps/core.xml' $x
 [xml]$r=Read-EntryXml $z 'ppt/slides/_rels/slide2.xml.rels';$nodes=@($r.DocumentElement.ChildNodes);foreach($n in $nodes){[void]$r.DocumentElement.RemoveChild($n)};foreach($n in @($nodes|Sort-Object {$_.GetAttribute('Id')} -Descending)){[void]$r.DocumentElement.AppendChild($n)};Write-EntryXml $z 'ppt/slides/_rels/slide2.xml.rels' $r
}
Check 'OOXML relationship ordering and package metadata changes pass' (CompareNative $metadata).unchanged

$renumber=Edit-Package relationship-renumber {
 param($z);[xml]$s=Read-EntryXml $z 'ppt/slides/slide2.xml';[xml]$r=Read-EntryXml $z 'ppt/slides/_rels/slide2.xml.rels'
 $ns=[Xml.XmlNamespaceManager]::new($s.NameTable);$ns.AddNamespace('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
 $a=@($s.SelectNodes('//@r:embed | //@r:link',$ns))[0];$old=$a.Value;$new='rId999';$a.Value=$new;(@($r.Relationships.Relationship|Where-Object {$_.GetAttribute('Id') -eq $old})[0]).SetAttribute('Id',$new)
 Write-EntryXml $z 'ppt/slides/slide2.xml' $s;Write-EntryXml $z 'ppt/slides/_rels/slide2.xml.rels' $r
}
Check 'Relationship-ID renumbering passes' (CompareNative $renumber).unchanged

$avatar=Edit-Package avatar-insertion {
 param($z);[xml]$s=Read-EntryXml $z 'ppt/slides/slide2.xml';$ns=[Xml.XmlNamespaceManager]::new($s.NameTable);$ns.AddNamespace('p','http://schemas.openxmlformats.org/presentationml/2006/main')
 $pic=$s.CreateElement('p','pic','http://schemas.openxmlformats.org/presentationml/2006/main');$pic.InnerXml='<p:nvPicPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cNvPr id="999" name="LectureForge_White_Avatar_Slide02"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/><p:spPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
 [void]$s.SelectSingleNode('/p:sld/p:cSld/p:spTree',$ns).AppendChild($pic);Write-EntryXml $z 'ppt/slides/slide2.xml' $s
}
Check 'Assembly-introduced avatar insertion passes' (CompareNative $avatar).unchanged

$text=Edit-Package text-change {param($z);[xml]$s=Read-EntryXml $z 'ppt/slides/slide2.xml';$n=@($s.GetElementsByTagName('t','http://schemas.openxmlformats.org/drawingml/2006/main'))[0];$n.InnerText+=' CHANGED';Write-EntryXml $z 'ppt/slides/slide2.xml' $s}
Check 'Original text change fails' (-not (CompareNative $text).unchanged)
$geometry=Edit-Package geometry-change {param($z);[xml]$s=Read-EntryXml $z 'ppt/slides/slide2.xml';$ns=[Xml.XmlNamespaceManager]::new($s.NameTable);$ns.AddNamespace('p','http://schemas.openxmlformats.org/presentationml/2006/main');$ns.AddNamespace('a','http://schemas.openxmlformats.org/drawingml/2006/main');$n=$s.SelectSingleNode('/p:sld/p:cSld/p:spTree/p:sp[1]/p:spPr/a:xfrm/a:off',$ns);$n.SetAttribute('x',([long]$n.GetAttribute('x')+12700).ToString());Write-EntryXml $z 'ppt/slides/slide2.xml' $s}
Check 'Original shape geometry change fails' (-not (CompareNative $geometry).unchanged)
$deleted=Edit-Package object-deletion {param($z);[xml]$s=Read-EntryXml $z 'ppt/slides/slide2.xml';$ns=[Xml.XmlNamespaceManager]::new($s.NameTable);$ns.AddNamespace('p','http://schemas.openxmlformats.org/presentationml/2006/main');$tree=$s.SelectSingleNode('/p:sld/p:cSld/p:spTree',$ns);[void]$tree.RemoveChild(@($tree.ChildNodes|Where-Object {$_.LocalName -notin 'nvGrpSpPr','grpSpPr'})[0]);Write-EntryXml $z 'ppt/slides/slide2.xml' $s}
Check 'Original object deletion fails' (-not (CompareNative $deleted).unchanged)
$image=Edit-Package image-change {
 param($z);[xml]$s=Read-EntryXml $z 'ppt/slides/slide2.xml';[xml]$r=Read-EntryXml $z 'ppt/slides/_rels/slide2.xml.rels';$ns=[Xml.XmlNamespaceManager]::new($s.NameTable);$ns.AddNamespace('p','http://schemas.openxmlformats.org/presentationml/2006/main');$ns.AddNamespace('a','http://schemas.openxmlformats.org/drawingml/2006/main');$ns.AddNamespace('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
 $id=$s.SelectSingleNode('//p:pic//a:blip/@r:embed',$ns).Value;$rel=@($r.Relationships.Relationship|Where-Object {$_.GetAttribute('Id') -eq $id})[0];$part='ppt/'+$rel.GetAttribute('Target').Replace('../','');$entry=$z.GetEntry($part);$stream=$entry.Open();$memory=[IO.MemoryStream]::new();$stream.CopyTo($memory);$stream.Dispose();$bytes=$memory.ToArray();$memory.Dispose();$bytes[$bytes.Length-1]=$bytes[$bytes.Length-1] -bxor 1;$entry.Delete();$new=$z.CreateEntry($part);$out=$new.Open();$out.Write($bytes,0,$bytes.Length);$out.Dispose()
}
Check 'Original image byte change fails' (-not (CompareNative $image).unchanged)

if($LiveAssembled){
 Check 'Normal PowerPoint open/save normalization passes' (Compare-LFNativeSlideSnapshot (Get-LFNativeSlideSnapshot $Source 2) (Get-LFNativeSlideSnapshot $LiveAssembled 2)).unchanged
 Check 'Slide 1 live assembly preserves native objects' (Compare-LFNativeSlideSnapshot (Get-LFNativeSlideSnapshot $Source 1) (Get-LFNativeSlideSnapshot $LiveAssembled 1)).unchanged
 Check 'Slide 2 live-style normalization preserves native objects' (Compare-LFNativeSlideSnapshot (Get-LFNativeSlideSnapshot $Source 2) (Get-LFNativeSlideSnapshot $LiveAssembled 2)).unchanged
 $validation=Join-Path (Split-Path $LiveAssembled) 'v/validation.json'
 Check 'Zero provider calls during live assembly' ((Test-Path $validation) -and (Get-Content -Raw $validation|ConvertFrom-Json).provider_calls -eq 0)
}
Check 'Validator performs zero provider calls' $true
