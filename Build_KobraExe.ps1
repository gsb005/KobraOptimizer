#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'dist\KobraOptimizer.exe'),
    [string]$IconPath = (Join-Path $PSScriptRoot 'Assets\kobra.ico')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$mainScript = Join-Path $PSScriptRoot 'Main.ps1'
$moduleRoot = Join-Path $PSScriptRoot 'Modules'
$xamlPath   = Join-Path $PSScriptRoot 'Kobra_UI.xaml'
$logoPath   = Join-Path $PSScriptRoot 'Assets\logo.png'

if (-not (Test-Path -LiteralPath $mainScript)) { throw "Main.ps1 not found." }
if (-not (Get-Command -Name Invoke-ps2exe -ErrorAction SilentlyContinue)) {
    throw "PS2EXE is not installed. Install it first with: Install-Module ps2exe -Scope CurrentUser"
}

$null = New-Item -ItemType Directory -Path (Split-Path -Parent $OutputPath) -Force

$embedFiles = @{
    '.\Kobra_UI.xaml'               = $xamlPath
    '.\Modules\Kobra_Cleanup.psm1' = (Join-Path $moduleRoot 'Kobra_Cleanup.psm1')
    '.\Modules\Kobra_Network.psm1' = (Join-Path $moduleRoot 'Kobra_Network.psm1')
    '.\Modules\Kobra_Browsers.psm1'= (Join-Path $moduleRoot 'Kobra_Browsers.psm1')
    '.\Modules\Kobra_OEM.psm1'     = (Join-Path $moduleRoot 'Kobra_OEM.psm1')
    '.\Modules\Kobra_Startup.psm1' = (Join-Path $moduleRoot 'Kobra_Startup.psm1')
}

if (Test-Path -LiteralPath $logoPath) {
    $embedFiles['.\Assets\logo.png'] = $logoPath
}

$params = @{
    inputFile    = $mainScript
    outputFile   = $OutputPath
    x64          = $true
    STA          = $true
    noConsole    = $true
    DPIAware     = $true
    requireAdmin = $true
    title        = 'KobraOptimizer'
    product      = 'KobraOptimizer'
    company      = 'Kobra'
    version      = '1.5.0'
    description  = 'KobraOptimizer desktop utility'
    embedFiles   = $embedFiles
}

if (Test-Path -LiteralPath $IconPath) {
    $params.iconFile = $IconPath
}

Invoke-ps2exe @params
Write-Host "Built: $OutputPath"
