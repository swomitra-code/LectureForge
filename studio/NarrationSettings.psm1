function Get-LFNarrationDefaults {
 Get-Content -Raw -Encoding UTF8 (Join-Path $PSScriptRoot 'presets/renewable-energy.json') | ConvertFrom-Json | ForEach-Object narration
}
function Resolve-LFNarrationSettings($Settings) {
 # One schema and validator for project persistence, previews, and provider calls.
 $result=Get-LFNarrationDefaults
 if($null -eq $Settings){return $result}
 if($Settings -isnot [pscustomobject] -and $Settings -isnot [Collections.IDictionary]){throw 'Narration settings must be an object.'}
 foreach($key in @($result.PSObject.Properties.Name)){
  $present=if($Settings -is [Collections.IDictionary]){$Settings.Contains($key)}else{$Settings.PSObject.Properties.Name -contains $key}
  if(-not $present){continue}
  $value=$Settings.$key
  if($key -in 'voice_id','model_id'){
   if($value -isnot [string] -or $value -notmatch '^[A-Za-z0-9_-]{1,128}$'){throw "Invalid narration $key."}
  }elseif($key -eq 'use_speaker_boost'){
   if($value -isnot [bool]){throw 'Speaker boost must be true or false.'}
  }else{
   if($null -eq $value -or $value -is [bool] -or $value -is [string] -or $value -is [array]){throw "Narration $key must be numeric."}
   $value=[double]$value
   $min=0;$max=1
   if($key -eq 'speed'){$min=0.7;$max=1.2}
   if($key -eq 'takes_per_slide'){$min=1;$max=5}
   if([double]::IsNaN($value) -or [double]::IsInfinity($value) -or $value -lt $min -or $value -gt $max){throw "Narration $key must be between $min and $max."}
   if($key -eq 'takes_per_slide'){if($value -ne [math]::Floor($value)){throw 'Takes per slide must be an integer.'};$value=[int]$value}
  }
  # Preserve production JSON number representation (e.g. 0.80 and 1.00):
  # existing revision hashes include the serialized preset.
  if($result.$key -ne $value){$result.$key=$value}
 }
 return $result
}
Export-ModuleMember -Function Get-LFNarrationDefaults,Resolve-LFNarrationSettings
