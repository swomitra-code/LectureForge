Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-V2Asset([string]$Path) {
    $file = Get-Item -LiteralPath $Path -ErrorAction Stop
    if ($file.PSIsContainer -or $file.Length -le 0) { throw "Invalid local asset: $Path" }
    [ordered]@{ path = $file.FullName; bytes = $file.Length; sha256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash }
}

function Get-V2DeckInfo([string]$Path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $archive.GetEntry('ppt/presentation.xml')
        if ($null -eq $entry) { throw 'Not a PowerPoint presentation.' }
        $reader = [IO.StreamReader]::new($entry.Open())
        try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        [ordered]@{
            slide_count = @($xml.presentation.sldIdLst.sldId).Count
            width_emu = [long]$xml.presentation.sldSz.cx
            height_emu = [long]$xml.presentation.sldSz.cy
            is_16_9 = ([long]$xml.presentation.sldSz.cx * 9 -eq [long]$xml.presentation.sldSz.cy * 16)
        }
    } finally { $archive.Dispose() }
}

function Write-V2Json([string]$Path, $Value) {
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 16) + [Environment]::NewLine, [Text.UTF8Encoding]::new($false))
}

function New-V2ProofProject {
    param([string]$Directory, [string]$Source, [string]$Narration, [string]$Avatar)
    $sourceAsset = Get-V2Asset $Source
    $narrationAsset = Get-V2Asset $Narration
    $avatarAsset = Get-V2Asset $Avatar
    $info = Get-V2DeckInfo $sourceAsset.path
    if ($info.slide_count -ne 14 -or -not $info.is_16_9) { throw 'Phase 2A requires the 14-slide 16:9 Solar Thermal deck.' }
    $root = [IO.Path]::GetFullPath($Directory)
    if (Test-Path -LiteralPath $root) { throw 'Project directory already exists; use the existing project for recomposition or choose a new directory.' }
    foreach ($folder in 'source','assets','renders','videos','validation') {
        New-Item -ItemType Directory -Path (Join-Path $root $folder) -Force | Out-Null
    }
    $bindings = @(
        @{asset=$sourceAsset; relative='source/lecture.pptx'},
        @{asset=$narrationAsset; relative=('assets/narration'+[IO.Path]::GetExtension($Narration))},
        @{asset=$avatarAsset; relative=('assets/avatar'+[IO.Path]::GetExtension($Avatar))}
    )
    foreach ($binding in $bindings) {
        $destination = Join-Path $root $binding.relative
        Copy-Item -LiteralPath $binding.asset.path -Destination $destination
        if ((Get-FileHash -LiteralPath $destination).Hash -ne $binding.asset.sha256) { throw 'Imported asset hash mismatch.' }
        $binding.asset.original_path = $binding.asset.path
        $binding.asset.path = $binding.relative
    }
    $record = [ordered]@{
        schema_version = '2.0-phase2a'
        project_name = 'Solar Thermal Module 2A - local technical proof'
        technical_proof_only = $true
        provider_calls_allowed = $false
        source = $sourceAsset
        source_info = $info
        narration = $narrationAsset
        narration_approval = 'Existing Energy Management selection; not approved Solar Thermal narration'
        avatar = $avatarAsset
        lip_sync_accepted = $false
        slide_number = 1
        project_placement = @{region='lower-right'; approved=$false; x=$null; y=$null; width=$null}
        # Provisional candidate placement only, never promoted to a project default here.
        candidate_placement = @{x=0.84375; y=0.78; width=0.1458333333333333; approved=$false}
        avatar_crop = @{width=640; height=400; x=870; y=80}
    }
    Write-V2Json (Join-Path $root 'project.json') $record
    Join-Path $root 'project.json'
}

function Read-V2ProofProject([string]$Path) {
    $manifest = (Resolve-Path -LiteralPath $Path).Path
    $record = Get-Content -LiteralPath $manifest -Raw | ConvertFrom-Json
    if ($record.schema_version -ne '2.0-phase2a' -or $record.provider_calls_allowed -ne $false -or $record.slide_number -ne 1) {
        throw 'Only offline Phase 2A Slide 1 projects are supported.'
    }
    $root = Split-Path -Parent $manifest
    foreach ($role in 'source','narration','avatar') {
        $asset = $record.$role
        $full = [IO.Path]::GetFullPath((Join-Path $root $asset.path))
        if (-not $full.StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Asset path escapes project.' }
        if ((Get-V2Asset $full).sha256 -ne $asset.sha256) { throw "Changed $role asset; refusing to proceed." }
    }
    [pscustomobject]@{root=$root; record=$record; manifest=$manifest}
}

Export-ModuleMember -Function Get-V2Asset,Get-V2DeckInfo,Write-V2Json,New-V2ProofProject,Read-V2ProofProject
