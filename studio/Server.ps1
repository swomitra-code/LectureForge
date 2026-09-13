param([Parameter(Mandatory=$true)][string]$RuntimeRoot,[int]$Port=8770,[Parameter(Mandatory=$true)][string]$Instance,[int]$WorkerPid)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Production.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Narration.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Authorization.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'AvatarProduction.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Assembly.psm1') -Force
Add-Type -AssemblyName System.Web
$root=Get-LFRoot $RuntimeRoot;$token=[guid]::NewGuid().ToString('N')+[guid]::NewGuid().ToString('N')
$listener=[Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback,$Port)
$stopping=$false
function Read-WorkerHeartbeat([string]$Path){
 # The worker atomically replaces this file. Permit replacement while reading;
 # briefly retry Windows sharing conflicts, but report persistent I/O failures.
 for($attempt=0;$attempt -lt 5;$attempt++){
  $reader=$null
  try{
   $share=[IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete
   $reader=[IO.StreamReader]::new([IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,$share),[Text.Encoding]::UTF8)
   return ($reader.ReadToEnd()|ConvertFrom-Json)
  }catch [IO.IOException]{if($attempt -eq 4){throw};Start-Sleep -Milliseconds 20}finally{if($reader){$reader.Dispose()}}
 }
}
function Reply($Stream,[int]$Code,$Value,[string]$Type='application/json; charset=utf-8'){
 $bytes=if($Value -is [byte[]]){$Value}elseif($Type.StartsWith('application/json')){[Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $Value -Depth 30 -Compress))}else{[Text.Encoding]::UTF8.GetBytes([string]$Value)}
 $header=[Text.Encoding]::ASCII.GetBytes("HTTP/1.1 $Code OK`r`nContent-Type: $Type`r`nContent-Length: $($bytes.Length)`r`nConnection: close`r`nCache-Control: no-store`r`nX-Content-Type-Options: nosniff`r`nContent-Security-Policy: default-src 'self'; frame-ancestors 'none'; base-uri 'none'`r`n`r`n")
 $Stream.Write($header,0,$header.Length);$Stream.Write($bytes,0,$bytes.Length)
}
function Send-Audio($Stream,[string]$Path,[string]$Range){
 $f=[IO.File]::OpenRead($Path)
 try{
  $start=0L;$end=$f.Length-1;$code=200;$extra=''
  if($Range){
   if($Range -notmatch '^bytes=(\d*)-(\d*)$' -or (-not $Matches[1] -and -not $Matches[2])){throw 'Unsupported audio range.'}
   if($Matches[1]){$start=[long]$Matches[1];if($Matches[2]){$end=[math]::Min($end,[long]$Matches[2])}}else{$start=[math]::Max(0,$f.Length-[long]$Matches[2])}
   if($start -gt $end -or $start -ge $f.Length){
    $h=[Text.Encoding]::ASCII.GetBytes("HTTP/1.1 416 Range Not Satisfiable`r`nContent-Range: bytes */$($f.Length)`r`nContent-Length: 0`r`nConnection: close`r`n`r`n");$Stream.Write($h,0,$h.Length);return
   }
   $code=206;$extra="Content-Range: bytes $start-$end/$($f.Length)`r`n"
  }
  $type=if($Path.EndsWith('.wav')){'audio/wav'}else{'audio/mpeg'};$length=$end-$start+1
  $h=[Text.Encoding]::ASCII.GetBytes("HTTP/1.1 $code OK`r`nContent-Type: $type`r`nAccept-Ranges: bytes`r`n$($extra)Content-Length: $length`r`nCache-Control: private, max-age=3600`r`nX-Content-Type-Options: nosniff`r`nConnection: close`r`n`r`n")
  $Stream.Write($h,0,$h.Length);$null=$f.Seek($start,[IO.SeekOrigin]::Begin);$buffer=New-Object byte[] 65536
  while($length -gt 0){$n=$f.Read($buffer,0,[int][math]::Min($buffer.Length,$length));if($n -le 0){throw 'Audio read incomplete.'};$Stream.Write($buffer,0,$n);$length-=$n}
 }finally{$f.Dispose()}
}
try{
 $listener.Start()
 [Console]::Out.WriteLine('Loopback listener started.')
 Write-LFJson (Join-Path $root 'runtime.json') @{instance=$Instance;server_pid=$PID;worker_pid=$WorkerPid;port=$Port;url="http://127.0.0.1:$Port/";status='running';milestone='E'}
 while(-not $stopping -and -not (Test-Path (Join-Path $root "$Instance.stop"))){
  if(-not (Get-Process -Id $WorkerPid -ErrorAction SilentlyContinue)){
   # Recover only this instance's worker. Durable leases protect existing children.
   $log=Join-Path $root ($Instance+'.recovery-'+[guid]::NewGuid().ToString('N').Substring(0,8))
   $args='-NoProfile -ExecutionPolicy Bypass -File "'+(Join-Path $PSScriptRoot 'Worker.ps1')+'" -RuntimeRoot "'+$root+'" -Instance '+$Instance
   $recovered=Start-Process (Join-Path $PSHOME 'powershell.exe') -PassThru -RedirectStandardOutput ($log+'.log') -RedirectStandardError ($log+'.err') -ArgumentList $args
   $WorkerPid=$recovered.Id
   Write-LFJson (Join-Path $root 'runtime.json') @{instance=$Instance;server_pid=$PID;worker_pid=$WorkerPid;port=$Port;url="http://127.0.0.1:$Port/";status='running';milestone='E'}
  }
  if(-not $listener.Pending()){Start-Sleep -Milliseconds 100;continue}
  $client=$listener.AcceptTcpClient();$stream=$client.GetStream();$stream.ReadTimeout=2000;$stream.WriteTimeout=10000
  [Console]::Out.WriteLine('Local request received.')
  try{
   $header=[Collections.Generic.List[byte]]::new()
   while($true){$b=$stream.ReadByte();if($b -lt 0){throw 'Incomplete HTTP request.'};$header.Add([byte]$b);$len=$header.Count;if($len -gt 16384){throw 'HTTP headers too large.'};if($len -ge 4 -and $header[$len-4] -eq 13 -and $header[$len-3] -eq 10 -and $header[$len-2] -eq 13 -and $header[$len-1] -eq 10){break}}
   $lines=[Text.Encoding]::ASCII.GetString($header.ToArray()) -split "`r`n";$request=$lines[0].Split(' ');$method=$request[0]
   $headers=@{};foreach($line in $lines|Select-Object -Skip 1){if($line.Contains(':')){$k,$v=$line.Split(':',2);$headers[$k.ToLowerInvariant()]=$v.Trim()}}
   if($headers['host'] -notin @("127.0.0.1:$Port","localhost:$Port")){throw 'Invalid local host.'}
   if($request[1] -notmatch '^/'){throw 'Invalid request target.'}
   $uri=[Uri]("http://127.0.0.1:$Port"+$request[1]);$path=$uri.AbsolutePath;$query=[Web.HttpUtility]::ParseQueryString($uri.Query)
   if($method -notin @('GET','POST')){throw 'Unsupported method.'}
   if($method -eq 'POST'){
    if($headers['x-lf-token'] -ne $token){throw 'Invalid session token. Reload the page.'}
    if($headers.ContainsKey('origin') -and $headers['origin'] -notin @("http://127.0.0.1:$Port","http://localhost:$Port")){throw 'Invalid origin.'}
   }
   if($headers.ContainsKey('transfer-encoding')){throw 'Chunked requests are unsupported.'}
   $size=0L;if($headers.ContainsKey('content-length')){$size=[long]$headers['content-length']}
   if($size -lt 0 -or $size -gt 104857600){throw 'Request exceeds 100 MB limit.'}
   if($headers['expect'] -eq '100-continue'){$continue=[Text.Encoding]::ASCII.GetBytes("HTTP/1.1 100 Continue`r`n`r`n");$stream.Write($continue,0,$continue.Length)}
   $stream.ReadTimeout=15000
   $body=New-Object byte[] ([int]$size);$read=0
   while($read -lt $size){$got=$stream.Read($body,$read,[int]$size-$read);if($got -le 0){throw 'Incomplete upload.'};$read+=$got}
   if($method -eq 'GET' -and $path -eq '/api/health'){
    $beatFile=Join-Path $root "$Instance.worker.json"
    $beat=if(Test-Path $beatFile){Read-WorkerHeartbeat $beatFile}else{[pscustomobject]@{stage='starting';utc=[DateTime]::UtcNow.ToString('o')}}
    Reply $stream 200 @{instance=$Instance;ready=($beat.stage -in @('idle','working') -and ([DateTime]::UtcNow-[DateTimeOffset]::Parse($beat.utc).UtcDateTime).TotalSeconds -lt 5);worker=$beat.stage;milestone='E'}
   }elseif($method -eq 'GET' -and $path -eq '/api/bootstrap'){
    Reply $stream 200 @{token=$token;root=(Join-Path $root 'Projects');presets=@(Get-LFPreset);projects=@(Get-LFProjects $root);milestone='E'}
   }elseif($method -eq 'GET' -and $path -eq '/api/project'){
    Reply $stream 200 (Read-LFProject $root $query['id'])
   }elseif($method -eq 'GET' -and $path -eq '/api/assembly-ready'){
    Reply $stream 200 (Get-LFAssemblyReadiness $root $query['id'])
   }elseif($method -eq 'GET' -and $path -eq '/api/assembly-status'){
    Reply $stream 200 (Get-LFAssemblyStatus $root $query['id'])
   }elseif($method -eq 'POST' -and $path -eq '/api/assembly-create'){
    $data=[Text.Encoding]::UTF8.GetString($body)|ConvertFrom-Json
    Reply $stream 200 (Request-LFAssembly $root $query['id'] $data.expected $data.folder)
   }elseif($method -eq 'POST' -and $path -eq '/api/assembly-placement'){
    $data=[Text.Encoding]::UTF8.GetString($body)|ConvertFrom-Json
    Set-LFAssemblyPlacement $root $query['id'] $data.slide $data.placement $data.version
    Reply $stream 200 @{saved=$true}
   }elseif($method -eq 'POST' -and $path -eq '/api/recording-open'){
    $data=[Text.Encoding]::UTF8.GetString($body)|ConvertFrom-Json
    Reply $stream 200 (Open-LFRecording $root $query['id'] $data.kind)
   }elseif($method -eq 'POST' -and $path -eq '/api/output-choose'){
    $dir=Get-LFProjectPath $root $query['id']
    Write-LFJson (Join-Path $dir 'folder-choice.json') @{state='pending'}
    $args='-NoProfile -STA -ExecutionPolicy Bypass -File "'+(Join-Path $PSScriptRoot 'Select-OutputFolder.ps1')+'" -ProjectDirectory "'+$dir+'"'
    $null=Start-Process (Join-Path $PSHOME 'powershell.exe') -PassThru -ArgumentList $args
    Reply $stream 200 @{choosing=$true}
   }elseif($method -eq 'GET' -and $path -eq '/api/output-choice'){
    $choice=Join-Path (Get-LFProjectPath $root $query['id']) 'folder-choice.json'
    Reply $stream 200 $(if(Test-Path $choice){Read-LFSharedJson $choice}else{@{state='pending'}})
   }elseif($method -eq 'GET' -and $path -eq '/api/avatar-status'){
    Reply $stream 200 (Get-LFAvatarStatus $root $query['id'])
   }elseif($method -eq 'POST' -and $path -eq '/api/avatar-pause'){
    $data=[Text.Encoding]::UTF8.GetString($body)|ConvertFrom-Json
    Set-LFAvatarPause $root $query['id'] ([bool]$data.paused)
    Reply $stream 200 @{saved=$true}
   }elseif($method -eq 'POST' -and $path -eq '/api/avatar-retry'){
    $data=[Text.Encoding]::UTF8.GetString($body)|ConvertFrom-Json
    Repair-LFAvatarJob $root $query['id'] $data.job $data.action $data.video_id
    Reply $stream 200 @{saved=$true}
   }elseif($method -eq 'GET' -and $path -eq '/api/selections'){
    Reply $stream 200 (Get-LFSelectionReview $root $query['id'])
   }elseif($method -eq 'POST' -and $path -eq '/api/selection-review'){
    $data=[Text.Encoding]::UTF8.GetString($body)|ConvertFrom-Json
    $null=Save-LFSelectionReview $root $query['id'] $data.expected $data.phase
    Reply $stream 200 (Get-LFSelectionReview $root $query['id'])
   }elseif($method -eq 'POST' -and $path -eq '/api/selection-change'){
    $data=[Text.Encoding]::UTF8.GetString($body)|ConvertFrom-Json
    Save-LFReviewNavigation $root $query['id'] $data.slide
    Reply $stream 200 @{saved=$true}
   }elseif($method -eq 'POST' -and $path -eq '/api/avatar-authorize'){
    $data=[Text.Encoding]::UTF8.GetString($body)|ConvertFrom-Json
    if($data.confirm -ne $true){throw 'Explicit final paid-generation authorization is required.'}
    Reply $stream 200 (Confirm-LFAvatarAuthorization $root $query['id'] $data.review_id $data.expected)
   }elseif($method -eq 'GET' -and $path -eq '/api/narration'){
    Reply $stream 200 (Get-LFNarration $root $query['id'])
   }elseif($method -eq 'GET' -and $path -eq '/api/narration-plan'){
    Reply $stream 200 (Get-LFNarrationPlan $root $query['id'])
   }elseif($method -eq 'POST' -and $path -eq '/api/narration-generate'){
    $data=[Text.Encoding]::UTF8.GetString($body)|ConvertFrom-Json
    if($data.confirm -ne $true){throw 'Explicit narration generation confirmation required.'}
    Add-LFNarrationBatch $root $query['id'] $data.version
    Reply $stream 200 (Get-LFNarration $root $query['id'])
   }elseif($method -eq 'POST' -and $path -eq '/api/narration-select'){
    $data=[Text.Encoding]::UTF8.GetString($body)|ConvertFrom-Json
    Set-LFNarrationSelection $root $query['id'] $data.slide $data.revision $data.take
    Reply $stream 200 (Get-LFNarration $root $query['id'])
   }elseif($method -eq 'POST' -and $path -eq '/api/narration-retry'){
    $data=[Text.Encoding]::UTF8.GetString($body)|ConvertFrom-Json
    if($data.confirm -ne $true){throw 'Explicit retry confirmation required.'}
    Retry-LFNarration $root $query['id'] $data.revision $data.take
    Reply $stream 200 (Get-LFNarration $root $query['id'])
   }elseif($method -eq 'POST' -and $path -eq '/api/thumbnails'){
    $null=Read-LFProject $root $query['id'];$dir=Get-LFProjectPath $root $query['id']
    [void][IO.Directory]::CreateDirectory((Join-Path $dir 'thumbs'))
    Write-LFJson (Join-Path $dir 'thumbs/request.json') @{requested=$true}
    Reply $stream 200 @{queued=$true}
   }elseif($method -eq 'GET' -and $path -eq '/api/thumbnail'){
    $dir=Get-LFProjectPath $root $query['id'];$n=[int]$query['slide']
    if($n -lt 1 -or $n -gt 500){throw 'Invalid slide.'}
    $file=Join-Path $dir "thumbs/s$n.png"
    if(Test-Path $file){Reply $stream 200 ([IO.File]::ReadAllBytes($file)) 'image/png'}else{Reply $stream 404 @{error='Thumbnail pending or unavailable.'}}
   }elseif($method -eq 'GET' -and $path -eq '/api/audio'){
    $dir=Get-LFProjectPath $root $query['id'];$state=Read-LFNarrationFile $dir
    $r=@($state.revisions|Where-Object id -eq $query['revision'])[0];$n=[int]$query['take']
    if(-not $r -or $n -notin 1,2,3){throw 'Unknown narration take.'}
    $t=$r.takes[$n-1];if(-not (Test-LFTake $dir $t)){throw 'Narration media hash mismatch or take unavailable.'}
    $file=Resolve-LFNarrationAsset $dir $t.asset.path
    Send-Audio $stream $file $headers['range']
   }elseif($method -eq 'POST' -and $path -eq '/api/new'){
    if($headers['content-type'] -ne 'application/octet-stream'){throw 'Expected a PPTX upload.'}
    Reply $stream 200 (New-LFProject $root $query['name'] $query['filename'] $body $query['preset'])
   }elseif($method -eq 'POST' -and $path -eq '/api/save'){
    $update=[Text.Encoding]::UTF8.GetString($body)|ConvertFrom-Json
    Reply $stream 200 (Invoke-LFNarrationLock $root $query['id'] {param($dir) Update-LFProject $root $query['id'] $update})
   }elseif($method -eq 'POST' -and $path -eq '/api/folder'){
    $null=Read-LFProject $root $query['id'];Start-Process explorer.exe -ArgumentList ('"'+(Get-LFProjectPath $root $query['id'])+'"')
    Reply $stream 200 @{opened=$true}
   }elseif($method -eq 'POST' -and $path -eq '/api/stop'){
    $stopping=$true;Reply $stream 200 @{stopped=$true}
   }elseif($method -eq 'GET' -and $path -in @('/','/studio.js','/studio.css','/selections.js','/avatars.js','/assembly.js')){
    $file=if($path -eq '/'){'index.html'}else{$path.TrimStart('/')};$type=if($file.EndsWith('.js')){'text/javascript'}elseif($file.EndsWith('.css')){'text/css'}else{'text/html; charset=utf-8'}
    Reply $stream 200 ([IO.File]::ReadAllBytes((Join-Path $PSScriptRoot "web/$file"))) $type
   }else{Reply $stream 404 @{error='This action is not available in Milestone E.'}}
  }catch{
   $message=$_.Exception.Message
   foreach($key in @($env:ELEVENLABS_API_KEY,$env:HEYGEN_API_KEY)){if($key){$message=$message.Replace($key,'[REDACTED]')}}
   [Console]::Error.WriteLine($message)
   try{Reply $stream 400 @{error=$message}}catch{[Console]::Error.WriteLine('Failed to send HTTP response: '+$_.Exception.Message)}
  }finally{$stream.Dispose();$client.Close()}
 }
}finally{
 $listener.Stop();[IO.File]::WriteAllText((Join-Path $root "$Instance.stop"),'stop')
 for($i=0;$i -lt 20;$i++){if(-not (Get-Process -Id $WorkerPid -ErrorAction SilentlyContinue)){break};Start-Sleep -Milliseconds 250}
 $current=Read-LFSharedJson (Join-Path $root 'runtime.json')
 if($current.instance -eq $Instance){Write-LFJson (Join-Path $root 'runtime.json') @{instance=$Instance;server_pid=$PID;worker_pid=$WorkerPid;port=$Port;url="http://127.0.0.1:$Port/";status='stopped';milestone='E'}}
}
