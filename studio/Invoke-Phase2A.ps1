param(
    [string]$ProjectDirectory = (Join-Path $PSScriptRoot '../work/v2-phase2a/solar-thermal'),
    [string]$Source = (Join-Path $PSScriptRoot '../source/RE-M2-Solar-Thermal-Overview.pptx'),
    [string]$Narration = (Join-Path $PSScriptRoot '../media/energy-management/slide-1-take-1.mp3'),
    [string]$Avatar = (Join-Path $PSScriptRoot '../media/energy-management/slide-1-SKM-BLUE-7-production-transparent.webm')
)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'Project.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Render.psm1') -Force
$manifest = New-V2ProofProject $ProjectDirectory $Source $Narration $Avatar
Export-V2Slide $manifest | Out-Null
Build-V2Slide $manifest | ConvertTo-Json -Depth 8
