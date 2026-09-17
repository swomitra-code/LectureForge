Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'

function Get-LFZipEntryBytes($Archive,[string]$Name){
 $entry=$Archive.GetEntry($Name);if(-not $entry){throw "Missing package part: $Name"}
 $stream=$entry.Open();$memory=[IO.MemoryStream]::new()
 try{$stream.CopyTo($memory);$memory.ToArray()}finally{$stream.Dispose();$memory.Dispose()}
}
function Get-LFBytesHash([byte[]]$Bytes){
 $sha=[Security.Cryptography.SHA256]::Create();try{([BitConverter]::ToString($sha.ComputeHash($Bytes))).Replace('-','')}finally{$sha.Dispose()}
}
function Get-LFPartTarget([string]$Part,[string]$Target){
 ([Uri]::new([Uri]('http://package.invalid/'+$Part),$Target)).AbsolutePath.TrimStart('/')
}
function ConvertTo-LFNativeXml($Node,$Relationships,$Archive,[string]$Part){
 if($Node.NodeType -eq [Xml.XmlNodeType]::Text){return [Security.SecurityElement]::Escape($Node.Value)}
 if($Node.NodeType -ne [Xml.XmlNodeType]::Element){return ''}
 # Office may rewrite these non-instructional bookkeeping extensions/flags.
 if($Node.LocalName -eq 'creationId'){return ''}
 $attributes=@()
 foreach($attribute in $Node.Attributes){
  if($attribute.Prefix -eq 'xmlns' -or $attribute.Name -eq 'xmlns' -or $attribute.LocalName -eq 'dirty'){continue}
  $value=$attribute.Value
  if($attribute.NamespaceURI -eq 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'){
   $relationship=$Relationships[$value]
   if(-not $relationship){$value='relationship:missing'}else{
    $target=$relationship.Target
    if($relationship.TargetMode -eq 'External'){$identity='external:'+([Uri]$target).AbsoluteUri}else{
     $targetPart=Get-LFPartTarget $Part $target;$entry=$Archive.GetEntry($targetPart)
     $identity=if($entry){'sha256:'+ (Get-LFBytesHash (Get-LFZipEntryBytes $Archive $targetPart))}else{'part:'+$targetPart}
    }
    $value='relationship:'+$relationship.Type+'|'+$identity
   }
  }
  $attributes+=@{name=('{'+$attribute.NamespaceURI+'}'+$attribute.LocalName);value=$value}
 }
 $attributes=@($attributes|Sort-Object name,value)
 $text='<{'+$Node.NamespaceURI+'}'+$Node.LocalName
 foreach($attribute in $attributes){$text+=' '+$attribute.name+'="'+[Security.SecurityElement]::Escape($attribute.value)+'"'}
 $children=@($Node.ChildNodes|ForEach-Object {ConvertTo-LFNativeXml $_ $Relationships $Archive $Part}|Where-Object {$_ -ne ''})
 if(-not $children.Count){return $text+'/>'}
 $text+='>'+($children -join '')+'</{'+$Node.NamespaceURI+'}'+$Node.LocalName+'>';$text
}
function Get-LFNativeSlideSnapshot([string]$DeckPath,[int]$SlideNumber){
 Add-Type -AssemblyName System.IO.Compression.FileSystem
 $archive=[IO.Compression.ZipFile]::OpenRead($DeckPath)
 try{
  [xml]$presentation=[Text.Encoding]::UTF8.GetString((Get-LFZipEntryBytes $archive 'ppt/presentation.xml'))
  [xml]$presentationRels=[Text.Encoding]::UTF8.GetString((Get-LFZipEntryBytes $archive 'ppt/_rels/presentation.xml.rels'))
  $slideId=@($presentation.presentation.sldIdLst.sldId)[$SlideNumber-1]
  $rid=$slideId.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
  $slideTarget=@($presentationRels.Relationships.Relationship|Where-Object Id -eq $rid)[0].Target
  $part=Get-LFPartTarget 'ppt/presentation.xml' $slideTarget
  [xml]$xml=[Text.Encoding]::UTF8.GetString((Get-LFZipEntryBytes $archive $part))
  $relsPart=([IO.Path]::GetDirectoryName($part).Replace('\','/'))+'/_rels/'+[IO.Path]::GetFileName($part)+'.rels'
  [xml]$relsXml=[Text.Encoding]::UTF8.GetString((Get-LFZipEntryBytes $archive $relsPart))
  $rels=@{};foreach($r in $relsXml.Relationships.Relationship){$rels[$r.GetAttribute('Id')]=@{Type=$r.GetAttribute('Type');Target=$r.GetAttribute('Target');TargetMode=$r.GetAttribute('TargetMode')}}
  $ns=[Xml.XmlNamespaceManager]::new($xml.NameTable);$ns.AddNamespace('p','http://schemas.openxmlformats.org/presentationml/2006/main')
  $objects=@()
  foreach($node in @($xml.SelectNodes('/p:sld/p:cSld/p:spTree/*[not(self::p:nvGrpSpPr) and not(self::p:grpSpPr)]',$ns))){
   $property=$node.SelectSingleNode('./p:nvSpPr/p:cNvPr | ./p:nvPicPr/p:cNvPr | ./p:nvGraphicFramePr/p:cNvPr | ./p:nvCxnSpPr/p:cNvPr | ./p:nvGrpSpPr/p:cNvPr',$ns)
   $id=if($property){[string]$property.id}else{'?'};$name=if($property){[string]$property.name}else{''}
   if($name -like 'LectureForge_White_Avatar_Slide*'){continue}
   $canonical=ConvertTo-LFNativeXml $node $rels $archive $part
   $objects+=@{id=$id;name=$name;kind=$node.LocalName;sha256=Get-LFBytesHash ([Text.Encoding]::UTF8.GetBytes($canonical));canonical=$canonical}
  }
  $properties=@()
  foreach($node in @($xml.SelectNodes('/p:sld/p:cSld/p:bg | /p:sld/p:clrMapOvr | /p:sld/p:transition',$ns))){$properties+=ConvertTo-LFNativeXml $node $rels $archive $part}
  [pscustomobject]@{slide=$SlideNumber;part=$part;objects=@($objects|Sort-Object id,name,kind);properties=@($properties|Sort-Object)}
 }finally{$archive.Dispose()}
}
function Compare-LFNativeSlideSnapshot($Before,$After){
 $changes=@();$beforeMap=@{};$afterMap=@{}
 foreach($o in $Before.objects){$beforeMap[$o.id+'|'+$o.name+'|'+$o.kind]=$o}
 foreach($o in $After.objects){$afterMap[$o.id+'|'+$o.name+'|'+$o.kind]=$o}
 foreach($key in @($beforeMap.Keys+$afterMap.Keys|Sort-Object -Unique)){
  if(-not $afterMap.ContainsKey($key)){$changes+=@{change='deleted';object=$key;before_sha256=$beforeMap[$key].sha256;after_sha256=$null}}
  elseif(-not $beforeMap.ContainsKey($key)){$changes+=@{change='added';object=$key;before_sha256=$null;after_sha256=$afterMap[$key].sha256}}
  elseif($beforeMap[$key].sha256 -ne $afterMap[$key].sha256){$changes+=@{change='modified';object=$key;before_sha256=$beforeMap[$key].sha256;after_sha256=$afterMap[$key].sha256;before_xml=$beforeMap[$key].canonical;after_xml=$afterMap[$key].canonical}}
 }
 if((ConvertTo-Json $Before.properties -Compress) -ne (ConvertTo-Json $After.properties -Compress)){$changes+=@{change='modified';object='slide-properties';before=$Before.properties;after=$After.properties}}
 [pscustomobject]@{unchanged=($changes.Count -eq 0);before_object_count=$Before.objects.Count;after_object_count=$After.objects.Count;changes=$changes}
}
Export-ModuleMember -Function Get-LFNativeSlideSnapshot,Compare-LFNativeSlideSnapshot
