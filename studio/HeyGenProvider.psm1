$ErrorActionPreference='Stop'
# Same v3 request shapes used by the completed Solar Thermal production run.
function Invoke-LFHeyGen([string]$Method,[string]$Path,$Body=$null){
 if(-not $env:HEYGEN_API_KEY){throw 'HEYGEN_API_KEY unavailable.'}
 $args=@{Method=$Method;Uri=('https://api.heygen.com'+$Path);Headers=@{'x-api-key'=$env:HEYGEN_API_KEY};TimeoutSec=90;UseBasicParsing=$true}
 if($null -ne $Body){$args.ContentType='application/json';$args.Body=($Body|ConvertTo-Json -Depth 15 -Compress)}
 try{(Invoke-RestMethod @args).data}catch{$code=0;try{$code=[int]$_.Exception.Response.StatusCode}catch{};throw "HeyGen request failed (HTTP $code). No automatic submission retry."}
}
function Send-LFHeyGenAudio([string]$Path){
 if(-not $env:HEYGEN_API_KEY){throw 'HEYGEN_API_KEY unavailable.'}
 # Keep the credential in memory, never in subprocess arguments or logs.
 Add-Type -AssemblyName System.Net.Http
 $client=[Net.Http.HttpClient]::new();$client.Timeout=[TimeSpan]::FromSeconds(120)
 $client.DefaultRequestHeaders.Add('x-api-key',$env:HEYGEN_API_KEY)
 $form=[Net.Http.MultipartFormDataContent]::new();$stream=[IO.File]::OpenRead($Path)
 try{
  $content=[Net.Http.StreamContent]::new($stream);$mime=if([IO.Path]::GetExtension($Path) -eq '.wav'){'audio/x-wav'}else{'audio/mpeg'};$content.Headers.ContentType=[Net.Http.Headers.MediaTypeHeaderValue]::new($mime);$form.Add($content,'file',[IO.Path]::GetFileName($Path))
  $response=$client.PostAsync('https://api.heygen.com/v3/assets',$form).GetAwaiter().GetResult()
  try{if(-not $response.IsSuccessStatusCode){throw 'Upload failed.'};$r=$response.Content.ReadAsStringAsync().GetAwaiter().GetResult()|ConvertFrom-Json;$id=$r.data.asset_id;if(-not $id){$id=$r.data.id};if(-not $id){throw 'Upload returned no asset ID.'};[string]$id}finally{$response.Dispose()}
 }catch{throw 'HeyGen audio upload failed; no generation was submitted.'}finally{$form.Dispose();$stream.Dispose();$client.Dispose()}
}
Export-ModuleMember -Function Invoke-LFHeyGen,Send-LFHeyGenAudio
