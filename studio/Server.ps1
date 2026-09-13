param([Parameter(Mandatory=$true)][string]$RuntimeRoot,[int]$Port=8770,[Parameter(Mandatory=$true)][string]$Instance,[int]$WorkerPid)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Production.psm1') -Force
Add-Type -AssemblyName System.Web
$root=Get-LFRoot $RuntimeRoot;$token=[guid]::NewGuid().ToString('N')+[guid]::NewGuid().ToString('N')
$listener=[Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback,$Port)
$stopping=$false
function Reply($Stream,[int]$Code,$Value,[string]$Type='application/json; charset=utf-8'){
 $bytes=if($Value -is [byte[]]){$Value}elseif($Type.StartsWith('application/json')){[Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $Value -Depth 30 -Compress))}else{[Text.Encoding]::UTF8.GetBytes([string]$Value)}
 $header=[Text.Encoding]::ASCII.GetBytes("HTTP/1.1 $Code OK`r`nContent-Type: $Type`r`nContent-Length: $($bytes.Length)`r`nConnection: close`r`nCache-Control: no-store`r`nX-Content-Type-Options: nosniff`r`nContent-Security-Policy: default-src 'self'; frame-ancestors 'none'; base-uri 'none'`r`n`r`n")
 $Stream.Write($header,0,$header.Length);$Stream.Write($bytes,0,$bytes.Length)
}
try{
 $listener.Start()
 [Console]::Out.WriteLine('Loopback listener started.')
 Write-LFJson (Join-Path $root 'runtime.json') @{instance=$Instance;server_pid=$PID;worker_pid=$WorkerPid;port=$Port;url="http://127.0.0.1:$Port/";status='running';milestone='A'}
 while(-not $stopping -and -not (Test-Path (Join-Path $root "$Instance.stop"))){
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
    $beat=Get-Content -Encoding UTF8 -Raw (Join-Path $root "$Instance.worker.json")|ConvertFrom-Json
    Reply $stream 200 @{instance=$Instance;ready=($beat.stage -eq 'idle' -and ([DateTime]::UtcNow-[DateTimeOffset]::Parse($beat.utc).UtcDateTime).TotalSeconds -lt 5);worker=$beat.stage;provider_calls=0;milestone='A'}
   }elseif($method -eq 'GET' -and $path -eq '/api/bootstrap'){
    Reply $stream 200 @{token=$token;root=(Join-Path $root 'Projects');presets=@(Get-LFPreset);projects=@(Get-LFProjects $root);milestone='A'}
   }elseif($method -eq 'GET' -and $path -eq '/api/project'){
    Reply $stream 200 (Read-LFProject $root $query['id'])
   }elseif($method -eq 'POST' -and $path -eq '/api/new'){
    if($headers['content-type'] -ne 'application/octet-stream'){throw 'Expected a PPTX upload.'}
    Reply $stream 200 (New-LFProject $root $query['name'] $query['filename'] $body $query['preset'])
   }elseif($method -eq 'POST' -and $path -eq '/api/save'){
    $update=[Text.Encoding]::UTF8.GetString($body)|ConvertFrom-Json
    Reply $stream 200 (Update-LFProject $root $query['id'] $update)
   }elseif($method -eq 'POST' -and $path -eq '/api/folder'){
    $null=Read-LFProject $root $query['id'];Start-Process explorer.exe -ArgumentList ('"'+(Get-LFProjectPath $root $query['id'])+'"')
    Reply $stream 200 @{opened=$true}
   }elseif($method -eq 'POST' -and $path -eq '/api/stop'){
    $stopping=$true;Reply $stream 200 @{stopped=$true}
   }elseif($method -eq 'GET' -and $path -in @('/','/studio.js','/studio.css')){
    $file=if($path -eq '/'){'index.html'}else{$path.TrimStart('/')};$type=if($file.EndsWith('.js')){'text/javascript'}elseif($file.EndsWith('.css')){'text/css'}else{'text/html; charset=utf-8'}
    Reply $stream 200 ([IO.File]::ReadAllBytes((Join-Path $PSScriptRoot "web/$file"))) $type
   }else{Reply $stream 404 @{error='This action is not available in Milestone A.'}}
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
 $current=Get-Content -Encoding UTF8 -Raw (Join-Path $root 'runtime.json')|ConvertFrom-Json
 if($current.instance -eq $Instance){Write-LFJson (Join-Path $root 'runtime.json') @{instance=$Instance;server_pid=$PID;worker_pid=$WorkerPid;port=$Port;url="http://127.0.0.1:$Port/";status='stopped';milestone='A'}}
}
