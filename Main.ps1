#requires -Version 5.1
# ==============================================================================
# KobraOptimizer v1.8.0 - Main Launcher
# Next-gen shell pass: section-based product UI built on the existing PowerShell backend
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-KobraAppRoot {
    if ($PSCommandPath) {
        return (Split-Path -Parent $PSCommandPath)
    }

    if ($MyInvocation.MyCommand.CommandType -eq 'ExternalScript' -and $MyInvocation.MyCommand.Definition) {
        return (Split-Path -Parent $MyInvocation.MyCommand.Definition)
    }

    $commandLine = [Environment]::GetCommandLineArgs()[0]
    if ($commandLine) {
        $candidate = Split-Path -Parent $commandLine
        if (-not [string]::IsNullOrWhiteSpace($candidate)) {
            return $candidate
        }
    }

    return (Get-Location).Path
}

$script:AppVersion   = '1.8.9'
$script:DonationUrl  = 'https://ko-fi.com/kobraoptimizer'
$script:ProjectRoot  = Get-KobraAppRoot
$script:ModuleRoot   = Join-Path $script:ProjectRoot 'Modules'
$script:XamlPath     = Join-Path $script:ProjectRoot 'Kobra_UI.xaml'
$script:LogoPath     = Join-Path $script:ProjectRoot 'Assets\logo.png'
$script:LogRoot      = Join-Path $script:ProjectRoot 'Logs'
$script:BackupRoot   = Join-Path $script:ProjectRoot 'Backups'
$script:TempRoot     = Join-Path 'C:\Temp' 'KobraOptimizer'
$script:ManifestRoot = Join-Path $script:TempRoot 'Manifests'
$null = New-Item -Path $script:LogRoot -ItemType Directory -Force -ErrorAction SilentlyContinue
$null = New-Item -Path $script:BackupRoot -ItemType Directory -Force -ErrorAction SilentlyContinue
$null = New-Item -Path $script:TempRoot -ItemType Directory -Force -ErrorAction SilentlyContinue
$null = New-Item -Path $script:ManifestRoot -ItemType Directory -Force -ErrorAction SilentlyContinue
$script:LogFile = Join-Path $script:LogRoot ("Kobra_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
$script:DebugRoot = 'C:\Temp'
$null = New-Item -Path $script:DebugRoot -ItemType Directory -Force -ErrorAction SilentlyContinue
$script:DebugLogFile = Join-Path $script:DebugRoot 'Kobra_Debug.log'
$script:LastRegistryBackup = $null

function Test-KobraAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Start-KobraElevated {
    $psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $psExe)) {
        $psExe = 'powershell.exe'
    }

    $argList = @(
        '-NoProfile',
        '-ExecutionPolicy','Bypass',
        '-File', ('"{0}"' -f (Join-Path $script:ProjectRoot 'Main.ps1'))
    )

    Start-Process -FilePath $psExe -Verb RunAs -ArgumentList $argList | Out-Null
}

function Get-KobraSafeCount {
    param($Value)

    if ($null -eq $Value) {
        return 0
    }

    return @($Value).Count
}

function Get-KobraSafeInt64 {
    param($Value)

    if ($null -eq $Value) {
        return [int64]0
    }

    $sum = ($Value | Measure-Object -Sum).Sum
    if ($null -eq $sum) {
        return [int64]0
    }

    return [int64]$sum
}

function Test-KobraPathExists {
    param([Parameter(Mandatory)][string]$Path)
    return [bool](Test-Path -LiteralPath $Path)
}

function Write-KobraDebug {
    param(
        [Parameter(Mandatory)][string]$Message,
        [switch]$NoTimestamp
    )

    try {
        $line = if ($NoTimestamp) { $Message } else { "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Message }
        Add-Content -Path $script:DebugLogFile -Value $line -Encoding UTF8
    }
    catch {}
}

function Get-KobraDebugSelectionState {
    $pairs = [ordered]@{}

    foreach ($name in @(
'ChkUserTemp','ChkSystemTemp','ChkWinUpdate','ChkThumbCache','ChkShaderCache','ChkRecycleBin',
'ChkChrome','ChkEdge','ChkFirefox','ChkOpera','ChkBrave','ChkVivaldi','ChkBrowserCookies','ChkBrowserHistory','ChkBrowserBackupBundle',
        'ChkRegistryClean','ChkRegistryBackup','ChkRestorePoint','ChkRegistry','ChkNetwork',
        'ChkDnsFlush','ChkDnsProfile','ChkHPDebloat','ChkSystemRestorePoint'
    )) {
        try {
            $control = Get-Variable -Scope Script -Name $name -ValueOnly -ErrorAction SilentlyContinue
            if ($null -ne $control) {
                $pairs[$name] = [bool]$control.IsChecked
            }
        }
        catch {
            $pairs[$name] = '<error>'
        }
    }

    try {
        $pairs['DnsProvider'] = Get-KobraSelectedDnsProfile
    }
    catch {
        $pairs['DnsProvider'] = '<error>'
    }

    return ($pairs.GetEnumerator() | ForEach-Object { '{0}={1}' -f $_.Key, $_.Value }) -join '; '
}


Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName PresentationCore

Write-KobraDebug -Message ('=== Kobra Debug Session Start ===')
Write-KobraDebug -Message ('ProjectRoot=' + $script:ProjectRoot)
Write-KobraDebug -Message ('ModuleRoot=' + $script:ModuleRoot)
Write-KobraDebug -Message ('XamlPath=' + $script:XamlPath)
Write-KobraDebug -Message ('LogFile=' + $script:LogFile)
Write-KobraDebug -Message ('DebugLog=' + $script:DebugLogFile)

if (-not (Test-KobraAdministrator)) {
    try {
        Start-KobraElevated
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "KobraOptimizer needs Administrator rights. Right-click Launch_Kobra.cmd or Main.ps1 and choose 'Run as administrator'.",
            "KobraOptimizer - Admin Required",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
    }
    exit
}

try {
    Get-ChildItem -LiteralPath $script:ProjectRoot -Recurse -File -ErrorAction SilentlyContinue |
        Unblock-File -ErrorAction SilentlyContinue
}
catch {
}

$modules = @(
    'Kobra_Cleanup.psm1',
    'Kobra_Network.psm1',
    'Kobra_Browsers.psm1',
    'Kobra_OEM.psm1',
    'Kobra_Startup.psm1'
)

foreach ($module in $modules) {
    $modulePath = Join-Path $script:ModuleRoot $module
    Write-KobraDebug -Message ('Importing module: ' + $modulePath)
    if (-not (Test-KobraPathExists -Path $modulePath)) {
        throw "Required module missing: $modulePath"
    }

    Import-Module $modulePath -Force
}

try {
    [xml]$appearance = Get-Content -Path $script:XamlPath -Raw
    $reader = New-Object System.Xml.XmlNodeReader $appearance
    $script:Window = [Windows.Markup.XamlReader]::Load($reader)
    Write-KobraDebug -Message 'XAML loaded successfully.'
}
catch {
    Write-KobraDebug -Message ('Failed to load XAML: ' + $_.Exception.Message)
    Write-Host "Failed to load XAML: $($_.Exception.Message)" -ForegroundColor Red
    Read-Host 'Press Enter to exit'
    exit
}

$controlNames = @(
    'BtnAnalyze','BtnRunSelected','BtnQuickShed','BtnCreateBackup','BtnOpenLogs','BtnOpenManifests',
    'StatusTextBox','ProgBar','BigProgBar','TxtBigStatus','TxtBigDetail','BtnCancelScan','KobraLogo','ResultsList','TxtResultsHeadline','TxtResultsSubHeadline','TxtResultsMode',
    'ChkRestorePoint','ChkRegistry','ChkNetwork','ChkDnsFlush','ChkDnsProfile','ChkHPDebloat',
'ChkUserTemp','ChkSystemTemp','ChkWinUpdate','ChkThumbCache','ChkShaderCache','ChkRecycleBin','TxtRegistryBackupStatus',
'ChkChrome','ChkEdge','ChkFirefox','ChkOpera','ChkBrave','ChkVivaldi','ChkBrowserCookies','ChkBrowserHistory','ChkBrowserBackupBundle','ChkRegistryClean','ChkRegistryBackup','ChkSystemRestorePoint','CmbDnsProvider',
    'BtnSystemScan','BtnSystemClean','BtnBrowserScan','BtnBrowserClean','BtnRegistryScan','BtnRegistryBackup','BtnRegistryClean','BtnCancelScan','BtnResultsRescan','BtnResultsBackCustom','BtnResultsOpenDebugLog',
    'StartupList','BtnStartupRefresh','BtnStartupDisable','BtnStartupEnable','ChkStartupShowMicrosoft',
    'BtnWindowsUpdate','BtnWindowsUpdateSettings','BtnWindowsStorage','BtnWindowsApps',
    'BtnWindowsStartupSettings','BtnWindowsGameMode','BtnWindowsGraphics','BtnWindowsPower',
    'BtnExit','BtnDonate','BtnDonateSidebar','BtnDisclaimer','BtnAboutMe','BtnToggleLog',
    'BtnNavDashboard','BtnNavAnalyze','BtnNavCustomClean','BtnNavResults','BtnNavTools','BtnNavStartup','BtnNavUtilities','BtnNavAbout',
    'ViewDashboard','ViewAnalyze','ViewOperationProgress','ViewCustomClean','ViewResults','ViewTools','ViewStartup','ViewUtilities','ViewAbout',
    'BtnDashboardQuickScan','BtnDashboardCustomClean','BtnDashboardPerformance','BtnDashboardStartup',
    'BtnQuickScanCustomize','BtnQuickScanBackup','BtnCustomAnalyze','BtnPerformanceApply','BtnPerformanceBackup',
    'TxtDashLastScan','TxtDashReclaimable','TxtDashRecords','TxtDashCategories','TxtDashSelections','TxtDashSafety','TxtRecentActivity',
    'TxtAnalyzeStage','TxtAnalyzeSubStage','TxtSelectedCategoryCount','TxtSelectedRecordCount','TxtSelectedBytes','TxtSelectedWarnings'
)

foreach ($name in $controlNames) {
    Set-Variable -Scope Script -Name $name -Value $script:Window.FindName($name)
}

$script:LogRow = $script:Window.FindName('LogRow')
$script:RecentActivityMax = 6
$script:LastAnalyzeManifest = $null
$script:LastAnalyzeTime = $null
$script:LastAnalyzeSelectionSignature = ''
$script:LastAnalyzeScope = 'All'
$script:CurrentView = 'Dashboard'
$script:CurrentResultsScope = 'All'
$script:LastCleanupSummary = $null
$script:ResultsMode = 'ScanResults'
$script:LastActionStatus = 'No scan yet'
$script:OperationOriginView = 'Analyze'
$script:SectionScanReady = @{ System = $false; Browser = $false; Registry = $false }
Write-KobraDebug -Message ('Initial selections: ' + (Get-KobraDebugSelectionState))

function Invoke-KobraUiRefresh {
    $null = $script:Window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Render)
}

function Write-KobraUiLog {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [switch]$NoTimestamp,
        [switch]$BlankLine
    )

    $line = if ($NoTimestamp) {
        $Message
    }
    else {
        "[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message
    }

    if ($BlankLine) {
        $line = "`r`n$line"
    }

    $script:StatusTextBox.AppendText($line + "`r`n")
    $script:StatusTextBox.ScrollToEnd()
    Add-Content -Path $script:LogFile -Value $line
    Write-KobraDebug -Message $line -NoTimestamp
    Update-KobraRecentActivity
    Invoke-KobraUiRefresh
}

function Set-KobraProgress {
    param(
        [int]$Value = 0,
        [switch]$Indeterminate
    )

    foreach ($bar in @($script:ProgBar, $script:BigProgBar)) {
        if ($null -eq $bar) { continue }
        if ($Indeterminate) {
            $bar.IsIndeterminate = $true
        }
        else {
            $bar.IsIndeterminate = $false
            $bar.Value = [Math]::Max(0, [Math]::Min(100, $Value))
        }
    }

    if ($null -ne $script:BigProgBar) {
        if ($Indeterminate -or ($Value -gt 0 -and $Value -lt 100)) {
            $script:BigProgBar.Height = 34
            try {
                $script:BigProgBar.Effect = New-Object System.Windows.Media.Effects.DropShadowEffect -Property @{ Color = [System.Windows.Media.ColorConverter]::ConvertFromString('#FF5A75'); BlurRadius = 26; ShadowDepth = 0; Opacity = 0.90 }
            }
            catch {}
        }
        else {
            $script:BigProgBar.Height = 18
            try {
                $script:BigProgBar.Effect = New-Object System.Windows.Media.Effects.DropShadowEffect -Property @{ Color = [System.Windows.Media.ColorConverter]::ConvertFromString('#FF5A75'); BlurRadius = 16; ShadowDepth = 0; Opacity = 0.45 }
            }
            catch {}
        }
    }

    Invoke-KobraUiRefresh
}

function Show-KobraOperationView {
    param(
        [Parameter(Mandatory)][string]$Status,
        [string]$Detail = 'Preparing modules...',
        [string]$OriginView = 'Analyze',
        [int]$Value = 0,
        [switch]$Indeterminate
    )

    $script:OperationOriginView = $OriginView
    if ($null -ne $script:TxtBigStatus) { $script:TxtBigStatus.Text = $Status }
    if ($null -ne $script:TxtBigDetail) { $script:TxtBigDetail.Text = $Detail }
    Switch-KobraView -ViewName 'OperationProgress'
    Set-KobraProgress -Value $Value -Indeterminate:$Indeterminate
}

function Update-KobraOperationView {
    param(
        [string]$Status,
        [string]$Detail,
        [int]$Value = 0,
        [switch]$Indeterminate
    )

    if ($null -ne $script:TxtBigStatus -and -not [string]::IsNullOrWhiteSpace($Status)) { $script:TxtBigStatus.Text = $Status }
    if ($null -ne $script:TxtBigDetail -and -not [string]::IsNullOrWhiteSpace($Detail)) { $script:TxtBigDetail.Text = $Detail }
    Set-KobraProgress -Value $Value -Indeterminate:$Indeterminate
}

function Complete-KobraOperationView {
    param(
        [string]$Status = 'Operation complete',
        [string]$Detail = 'Kobra finished the current task.',
        [string]$NextView = 'Results'
    )

    if ($null -ne $script:TxtBigStatus) { $script:TxtBigStatus.Text = $Status }
    if ($null -ne $script:TxtBigDetail) { $script:TxtBigDetail.Text = $Detail }
    Set-KobraProgress 100
    if (-not [string]::IsNullOrWhiteSpace($NextView)) {
        Switch-KobraView -ViewName $NextView
    }
}

function Set-KobraButtonsEnabled {
    param([bool]$Enabled)

    $buttonNames = @(
        'BtnAnalyze','BtnRunSelected','BtnQuickShed','BtnCreateBackup','BtnOpenLogs','BtnOpenManifests','BtnResultsRescan','BtnResultsBackCustom','BtnResultsOpenDebugLog',
        'BtnStartupRefresh','BtnStartupDisable','BtnStartupEnable',
        'BtnExit','BtnDonate','BtnDonateSidebar','BtnDisclaimer','BtnAboutMe','BtnToggleLog',
        'BtnNavDashboard','BtnNavAnalyze','BtnNavCustomClean','BtnNavResults','BtnNavTools','BtnNavStartup','BtnNavUtilities','BtnNavAbout',
        'BtnDashboardQuickScan','BtnDashboardCustomClean','BtnDashboardPerformance','BtnDashboardStartup',
        'BtnQuickScanCustomize','BtnQuickScanBackup','BtnCustomAnalyze','BtnPerformanceApply','BtnPerformanceBackup',
        'BtnSystemScan','BtnSystemClean','BtnBrowserScan','BtnBrowserClean','BtnRegistryScan','BtnRegistryBackup','BtnRegistryClean','BtnCancelScan'
    )

    foreach ($name in $buttonNames) {
        $control = Get-Variable -Scope Script -Name $name -ValueOnly -ErrorAction SilentlyContinue
        if ($null -ne $control) { $control.IsEnabled = $Enabled }
    }

    if ($null -ne $script:ChkStartupShowMicrosoft) {
        $script:ChkStartupShowMicrosoft.IsEnabled = $Enabled
    }

    if ($Enabled) {
        Set-KobraSectionActionState
    }
    else {
        if ($null -ne $script:BtnSystemClean) { $script:BtnSystemClean.IsEnabled = $false }
        if ($null -ne $script:BtnBrowserClean) { $script:BtnBrowserClean.IsEnabled = $false }
        if ($null -ne $script:BtnRegistryClean) { $script:BtnRegistryClean.IsEnabled = $false }
    }

    Invoke-KobraUiRefresh
}

function Set-KobraPreferredFont {
    $fontName = 'Segoe UI Variable Text'
    try {
        $font = New-Object System.Windows.Media.FontFamily($fontName)
        $script:Window.FontFamily = $font
        $script:StatusTextBox.FontFamily = $font
    }
    catch {
        $fontName = 'Segoe UI Variable Text'
    }
    return $fontName
}

function Set-KobraWindowBounds {
    try {
        $workArea = [System.Windows.SystemParameters]::WorkArea
        $maxHeight = [Math]::Floor($workArea.Height - 20)
        $maxWidth  = [Math]::Floor($workArea.Width - 20)

        if ($maxHeight -lt 680) { $maxHeight = 680 }
        if ($maxWidth -lt 1100) { $maxWidth = 1100 }

        if ($script:Window.MaxHeight -gt $maxHeight) {
            $script:Window.MaxHeight = $maxHeight
        }
        if ($script:Window.Height -gt $script:Window.MaxHeight) {
            $script:Window.Height = $script:Window.MaxHeight
        }
        if ($script:Window.MinHeight -gt $script:Window.MaxHeight) {
            $script:Window.MinHeight = $script:Window.MaxHeight
        }

        if ($script:Window.MaxWidth -gt $maxWidth) {
            $script:Window.MaxWidth = $maxWidth
        }
        if ($script:Window.Width -gt $script:Window.MaxWidth) {
            $script:Window.Width = $script:Window.MaxWidth
        }
        if ($script:Window.MinWidth -gt $script:Window.MaxWidth) {
            $script:Window.MinWidth = $script:Window.MaxWidth
        }
    }
    catch {
    }
}

function Update-KobraDnsControls {
    $isEnabled = [bool]$script:ChkDnsProfile.IsChecked
    $script:CmbDnsProvider.IsEnabled = $isEnabled
    $script:CmbDnsProvider.Opacity = if ($isEnabled) { 1.0 } else { 0.72 }
    Invoke-KobraUiRefresh
}

function Get-KobraSafeMeasureSum {
    param(
        [Parameter(Mandatory)]
        [object[]]$Items,
        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    $sum = [int64]0
    foreach ($item in $Items) {
        if ($null -eq $item) { continue }
        if ($item.PSObject.Properties.Name -contains $PropertyName) {
            $value = $item.$PropertyName
            if ($null -ne $value) {
                try { $sum += [int64]$value } catch {}
            }
        }
    }
    return $sum
}

function New-KobraDeleteManifest {
    param(
        [string[]]$CleanupTargets,
        [string[]]$BrowserTargets,
        [string[]]$BrowserComponents = @('Cache'),
        [bool]$IncludeRegistryTraces = $false
    )

    $cleanupCandidates = @()
    $browserCandidates = @()
    $registryCandidates = @()

    if ((Get-KobraSafeCount $CleanupTargets) -gt 0) {
        $cleanupCandidates = @(Get-KobraCleanupCandidates -Targets $CleanupTargets)
    }
    if ((Get-KobraSafeCount $BrowserTargets) -gt 0) {
        $browserCandidates = @(Get-KobraBrowserCandidates -Browsers $BrowserTargets -Components $BrowserComponents)
    }
    if ($IncludeRegistryTraces) {
        $registryCandidates = @(Get-KobraRegistryCandidates)
    }

    $allCandidates = @($cleanupCandidates) + @($browserCandidates) + @($registryCandidates)
    $totalBytes = Get-KobraSafeMeasureSum -Items $allCandidates -PropertyName 'SizeBytes'
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $timeDisplay = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $manifestPath = Join-Path $script:ManifestRoot ("DeleteManifest_{0}.txt" -f $timestamp)
    $latestPath = Join-Path $script:ManifestRoot 'DeleteManifest_latest.txt'

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('KobraOptimizer Delete Manifest')
    $lines.Add(('Generated: {0}' -f $timeDisplay))
    $lines.Add(('Project Root: {0}' -f $script:ProjectRoot))
    $lines.Add(('Estimated Candidate Files: {0}' -f (Get-KobraSafeCount $allCandidates)))
    $lines.Add(('Estimated Candidate Size: {0:N2} MB' -f ($totalBytes / 1MB)))
    $lines.Add('')

    if ((Get-KobraSafeCount $CleanupTargets) -gt 0) {
        $lines.Add('Cleanup Targets:')
        foreach ($target in $CleanupTargets) { $lines.Add(('  - {0}' -f $target)) }
        $lines.Add('')
    }

    if ((Get-KobraSafeCount $BrowserTargets) -gt 0) {
        $lines.Add('Browser Targets:')
        foreach ($target in $BrowserTargets) { $lines.Add(('  - {0}' -f $target)) }
        if ((Get-KobraSafeCount $BrowserComponents) -gt 0) {
            $lines.Add(('  Components: {0}' -f ($BrowserComponents -join ', ')))
        }
        $lines.Add('')
    }

    if ($IncludeRegistryTraces) {
        $lines.Add('Registry Targets:')
        $lines.Add('  - Safe user MRU / typed history traces')
        $lines.Add('')
    }

    if ((Get-KobraSafeCount $cleanupCandidates) -gt 0) {
        $lines.Add('Cleanup Candidate Files:')
        foreach ($item in $cleanupCandidates) {
            $lines.Add(('  [{0}] {1} ({2:N2} KB)' -f $item.Category, $item.Path, ($item.SizeBytes / 1KB)))
        }
        $lines.Add('')
    }

    if ((Get-KobraSafeCount $browserCandidates) -gt 0) {
        $lines.Add('Browser Candidate Files:')
        foreach ($item in $browserCandidates) {
            $lines.Add(('  [{0}] {1} ({2:N2} KB)' -f $item.Category, $item.Path, ($item.SizeBytes / 1KB)))
        }
        $lines.Add('')
    }

    if ((Get-KobraSafeCount $registryCandidates) -gt 0) {
        $lines.Add('Registry Candidate Items:')
        foreach ($item in $registryCandidates) {
            $lines.Add(('  [{0}] {1} ({2:N2} KB)' -f $item.Category, $item.Path, ($item.SizeBytes / 1KB)))
        }
        $lines.Add('')
    }

    if ((Get-KobraSafeCount $allCandidates) -eq 0) {
        $lines.Add('No file deletions are currently queued.')
    }

    Set-Content -Path $manifestPath -Value $lines -Encoding UTF8
    Copy-Item -LiteralPath $manifestPath -Destination $latestPath -Force

    $summary = @()
    foreach ($group in @($allCandidates | Group-Object Category)) {
        $groupItems = @($group.Group)
        $firstItem = $groupItems | Select-Object -First 1
        $note = ''
        if ($null -ne $firstItem -and ($firstItem.PSObject.Properties.Name -contains 'Note')) {
            $note = [string]$firstItem.Note
        }

        $summary += [pscustomobject]@{
            Category  = $(if ([string]::IsNullOrWhiteSpace(($group.Name -as [string]))) { '<unknown>' } else { $group.Name })
            Items     = @($groupItems).Count
            SizeBytes = Get-KobraSafeMeasureSum -Items $groupItems -PropertyName 'SizeBytes'
            Note      = $note
        }
    }

    [pscustomobject]@{
        ManifestPath    = $manifestPath
        LatestPath      = $latestPath
        CandidateCount  = (Get-KobraSafeCount $allCandidates)
        TotalBytes      = $totalBytes
        TotalMB         = [Math]::Round(($totalBytes / 1MB), 2)
        CategorySummary = $summary
    }
}

function Show-KobraExecutionConfirmation {
    param(
        [string[]]$CleanupTargets,
        [string[]]$BrowserTargets,
        [string[]]$BrowserComponents = @('Cache'),
        [bool]$DoTls,
        [bool]$DoNetwork,
        [bool]$DoDnsFlush,
        [bool]$DoDnsProfile,
        [bool]$DoHpDebloat,
        [bool]$DoRestore,
        [bool]$DoRegistryTraces,
        [string]$DnsProfile,
        [psobject]$ManifestInfo
    )

    $actions = New-Object System.Collections.Generic.List[string]
    if ((Get-KobraSafeCount $CleanupTargets) -gt 0) {
        $actions.Add(('Cleanup: {0}' -f (($CleanupTargets -join ', '))))
    }
    if ((Get-KobraSafeCount $BrowserTargets) -gt 0) {
        $componentLabel = if ((Get-KobraSafeCount $BrowserComponents) -gt 0) { $BrowserComponents -join ' + ' } else { 'Cache' }
        $actions.Add(('Browser cleanup ({0}): {1}' -f $componentLabel, ($BrowserTargets -join ', ')))
    }
    if ($DoTls) { $actions.Add('TLS hardening') }
    if ($DoNetwork) { $actions.Add('Network TCP strike') }
    if ($DoDnsFlush) { $actions.Add('Flush DNS cache') }
    if ($DoDnsProfile) { $actions.Add(('Apply DNS profile: {0}' -f $DnsProfile)) }
    if ($DoHpDebloat) { $actions.Add('HP telemetry trim') }
    if ($DoRegistryTraces) { $actions.Add('Registry traces cleanup (safe MRU / typed history only)') }
    if ((Get-KobraSafeCount $BrowserTargets) -gt 0 -and ($null -ne $script:ChkBrowserBackupBundle) -and [bool]$script:ChkBrowserBackupBundle.IsChecked) { $actions.Add('Create system backup bundle before browser cleanup') }
    if (Test-KobraRegistryBackupSelected) { $actions.Add('Registry backup bundle before cleanup') }
    if ($DoRestore -and ($DoTls -or $DoNetwork -or $DoHpDebloat)) { $actions.Add('Create restore point first') }

    $messageLines = New-Object System.Collections.Generic.List[string]
    $messageLines.Add('You are about to execute the following Kobra actions:')
    $messageLines.Add('')
    foreach ($action in $actions) { $messageLines.Add(('- {0}' -f $action)) }
    $messageLines.Add('')
    if ($null -ne $ManifestInfo) {
        $messageLines.Add(('Estimated deletion candidates: {0}' -f $ManifestInfo.CandidateCount))
        $messageLines.Add(('Estimated reclaimable space: {0:N2} MB' -f $ManifestInfo.TotalMB))
        $messageLines.Add(('Delete manifest: {0}' -f $ManifestInfo.ManifestPath))
        $messageLines.Add('')
    }
    $messageLines.Add('Proceed?')

    $result = [System.Windows.MessageBox]::Show(
        ($messageLines -join "`r`n"),
        'KobraOptimizer - Confirm Execution',
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning,
        [System.Windows.MessageBoxResult]::No
    )

    return ($result -eq [System.Windows.MessageBoxResult]::Yes)
}

function Show-KobraYesNo {
    param(
        [Parameter(Mandatory)][string]$Message,
        [string]$Title = 'KobraOptimizer'
    )

    $result = [System.Windows.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question,
        [System.Windows.MessageBoxResult]::No
    )

    return ($result -eq [System.Windows.MessageBoxResult]::Yes)
}

function Get-KobraCleanupTargetsFromUi {
    $targets = @()
    if ($script:ChkUserTemp.IsChecked)    { $targets += 'UserTemp' }
    if ($script:ChkSystemTemp.IsChecked)  { $targets += 'SystemTemp' }
    if ($script:ChkWinUpdate.IsChecked)   { $targets += 'WindowsUpdate' }
    if ($script:ChkThumbCache.IsChecked)  { $targets += 'ThumbnailCache' }
    if ($script:ChkShaderCache.IsChecked) { $targets += 'ShaderCache' }
    if ($script:ChkRecycleBin.IsChecked)  { $targets += 'RecycleBin' }
    return $targets
}

function Get-KobraBrowserTargetsFromUi {
    $browsers = @()
    if ($script:ChkChrome.IsChecked)  { $browsers += 'Chrome' }
    if ($script:ChkEdge.IsChecked)    { $browsers += 'Edge' }
    if ($script:ChkFirefox.IsChecked) { $browsers += 'Firefox' }
    if ($null -ne $script:ChkOpera -and $script:ChkOpera.IsChecked) { $browsers += 'Opera' }
    if ($null -ne $script:ChkBrave -and $script:ChkBrave.IsChecked) { $browsers += 'Brave' }
    if ($null -ne $script:ChkVivaldi -and $script:ChkVivaldi.IsChecked) { $browsers += 'Vivaldi' }
    return $browsers
}

function Get-KobraBrowserComponentsFromUi {
    $components = @('Cache')
    if ($script:ChkBrowserCookies.IsChecked) {
        $components += 'Cookies'
    }
    if ($null -ne $script:ChkBrowserHistory -and $script:ChkBrowserHistory.IsChecked) {
        $components += 'History'
    }
    return $components
}

function Test-KobraRegistryCleanupSelected {
    return ($null -ne $script:ChkRegistryClean -and [bool]$script:ChkRegistryClean.IsChecked)
}

function Test-KobraRegistryBackupSelected {
    return (Test-KobraRegistryCleanupSelected) -and ($null -ne $script:ChkRegistryBackup -and [bool]$script:ChkRegistryBackup.IsChecked)
}

function New-KobraExecutionCategoryResult {
    param(
        [Parameter(Mandatory)][string]$Section,
        [Parameter(Mandatory)][string]$Category,
        [int]$FoundCount = 0,
        [int]$AttemptedCount = 0,
        [int]$RemovedCount = 0,
        [int]$SkippedCount = 0,
        [int]$FailedCount = 0,
        [int]$LockedCount = 0,
        [int64]$FoundBytes = 0,
        [int64]$RemovedBytes = 0,
        [string]$Reason = ''
    )

    [pscustomobject]@{
        Section        = $Section
        Category       = $Category
        FoundCount     = $FoundCount
        AttemptedCount = $AttemptedCount
        RemovedCount   = $RemovedCount
        SkippedCount   = $SkippedCount
        FailedCount    = $FailedCount
        LockedCount    = $LockedCount
        FoundBytes     = $FoundBytes
        RemovedBytes   = $RemovedBytes
        Reason         = $Reason
    }
}

function Get-KobraSelectionSignature {
    param(
        [string[]]$CleanupTargets = @(),
        [string[]]$BrowserTargets = @(),
        [string[]]$BrowserComponents = @(),
        [bool]$IncludeRegistryTraces = $false,
        [string]$Scope = 'All'
    )

    $parts = @(
        'scope=' + $Scope,
        'cleanup=' + (@($CleanupTargets | Sort-Object) -join ','),
        'browser=' + (@($BrowserTargets | Sort-Object) -join ','),
        'components=' + (@($BrowserComponents | Sort-Object) -join ','),
        'registry=' + $IncludeRegistryTraces
    )

    return ($parts -join ';')
}

function Get-KobraCurrentAnalyzeSelectionSignature {
    param([string]$Scope = 'All')

    switch ($Scope) {
        'System' {
            return Get-KobraSelectionSignature -CleanupTargets @(Get-KobraCleanupTargetsFromUi) -Scope 'System'
        }
        'Browser' {
            return Get-KobraSelectionSignature -BrowserTargets @(Get-KobraBrowserTargetsFromUi) -BrowserComponents @(Get-KobraBrowserComponentsFromUi) -Scope 'Browser'
        }
        'Registry' {
            return Get-KobraSelectionSignature -IncludeRegistryTraces (Test-KobraRegistryCleanupSelected) -Scope 'Registry'
        }
        default {
            return Get-KobraSelectionSignature -CleanupTargets @(Get-KobraCleanupTargetsFromUi) -BrowserTargets @(Get-KobraBrowserTargetsFromUi) -BrowserComponents @(Get-KobraBrowserComponentsFromUi) -IncludeRegistryTraces (Test-KobraRegistryCleanupSelected) -Scope 'All'
        }
    }
}

function Test-KobraAnalyzeManifestMatchesCurrentSelection {
    param([string]$Scope = 'All')

    if ($null -eq $script:LastAnalyzeManifest) {
        return $false
    }

    $expected = Get-KobraCurrentAnalyzeSelectionSignature -Scope $Scope
    return ($script:LastAnalyzeSelectionSignature -eq $expected -and $script:LastAnalyzeScope -eq $Scope)
}

function Get-KobraRegistryBackupStatus {
    if ($null -eq $script:LastRegistryBackup) {
        return 'No backup yet'
    }

    if (($script:LastRegistryBackup.ExportCount -gt 0) -and ($script:LastRegistryBackup.SkippedCount -gt 0)) {
        return 'Backup partially completed'
    }

    if ($script:LastRegistryBackup.ExportCount -gt 0) {
        return 'Backup completed'
    }

    return 'No backup yet'
}

function Update-KobraLastActionStatus {
    param([Parameter(Mandatory)][string]$Status)

    $script:LastActionStatus = $Status
    if ($null -ne $script:TxtResultsMode) {
        $script:TxtResultsMode.Text = $Status
    }
}

function Update-KobraRegistryBackupStatusUi {
    if ($null -eq $script:TxtRegistryBackupStatus) {
        return
    }

    $status = Get-KobraRegistryBackupStatus
    $script:TxtRegistryBackupStatus.Text = ('Registry backup status: {0}' -f $status)
}

function Convert-KobraRegistryTracePathToNative {
    param([Parameter(Mandatory)][string]$Path)

    if ($Path.StartsWith('HKCU:\', [System.StringComparison]::OrdinalIgnoreCase)) {
        return 'HKCU\' + $Path.Substring(6)
    }

    if ($Path.StartsWith('HKLM:\', [System.StringComparison]::OrdinalIgnoreCase)) {
        return 'HKLM\' + $Path.Substring(6)
    }

    if ($Path.StartsWith('HKCR:\', [System.StringComparison]::OrdinalIgnoreCase)) {
        return 'HKCR\' + $Path.Substring(6)
    }

    if ($Path.StartsWith('HKU:\', [System.StringComparison]::OrdinalIgnoreCase)) {
        return 'HKU\' + $Path.Substring(5)
    }

    return $Path
}

function New-KobraRegistryTraceBackup {
    param([scriptblock]$Log)

    $definitions = @(Get-KobraRegistryTraceDefinitions)
    if ($definitions.Count -eq 0) {
        if ($Log) { & $Log 'Registry backup skipped: no matching registry trace keys currently exist.' }
        return $null
    }

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backupPath = Join-Path $script:BackupRoot ('RegistryTraceBackup_' + $stamp)
    $null = New-Item -Path $backupPath -ItemType Directory -Force -ErrorAction Stop

    $manifestPath = Join-Path $backupPath 'RegistryBackupManifest.txt'
    $manifestLines = New-Object System.Collections.Generic.List[string]
    $manifestLines.Add('KobraOptimizer registry trace backup')
    $manifestLines.Add(('Created: {0}' -f (Get-Date)))
    $manifestLines.Add('')

    $index = 1
    $foundCount = 0
    $exportCount = 0
    $skippedCount = 0
    $successfulPaths = @()
    foreach ($definition in $definitions) {
        if (-not (Test-Path -LiteralPath $definition.Path)) {
            $skippedCount++
            $manifestLines.Add(('{0} -> SKIPPED (missing key)' -f $definition.Path))
            if ($Log) { & $Log ("Registry backup skipped: {0} - missing key" -f $definition.Path) }
            continue
        }

        $foundCount++
        if ($Log) { & $Log ("Registry key found: {0}" -f $definition.Path) }
        $regNativePath = Convert-KobraRegistryTracePathToNative -Path $definition.Path
        $safeSeed = $regNativePath.Replace('HKCU\','HKCU_').Replace('HKLM\','HKLM_')
        $safeName = '{0:D2}_{1}.reg' -f $index, ($safeSeed -replace '[\\/:*?"<>| ]','_')
        $targetFile = Join-Path $backupPath $safeName
        $manifestLines.Add(('{0} -> {1}' -f $regNativePath, $safeName))

        $output = & reg.exe export $regNativePath $targetFile /y 2>&1
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $targetFile)) {
            $skippedCount++
            $detail = ($output | Out-String).Trim()
            if ([string]::IsNullOrWhiteSpace($detail)) { $detail = 'reg.exe did not create an export file.' }
            $manifestLines.Add(('  SKIPPED: {0}' -f $detail))
            if ($Log) { & $Log ("Registry backup failed: {0} - {1}" -f $regNativePath, $detail) }
            $index++
            continue
        }

        $exportCount++
        $successfulPaths += $definition.Path
        if ($Log) { & $Log ("Registry backup exported: {0}" -f $targetFile) }
        $index++
    }

    if ($exportCount -eq 0) {
        if ($Log) { & $Log 'Registry backup failed: zero registry keys exported successfully.' }
        throw 'Registry backup failed: zero registry keys exported successfully.'
    }

    Set-Content -Path $manifestPath -Value $manifestLines -Encoding UTF8
    if ($Log) { & $Log ("Registry backup bundle ready: {0} ({1} exported, {2} skipped)." -f $backupPath, $exportCount, $skippedCount) }

    return [pscustomobject]@{
        BackupPath = $backupPath
        ManifestPath = $manifestPath
        FoundCount = $foundCount
        ExportCount = $exportCount
        SkippedCount = $skippedCount
        SuccessfulPaths = @($successfulPaths)
        CreatedAt = Get-Date
    }
}


function Get-KobraRegistryTraceDefinitions {
    return @(
        [pscustomobject]@{ Path='HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU'; Mode='Values'; Ignore=@('MRUList'); Note='Safe user Run dialog history only'; Category='Registry - Run history'; Preview='Run history' },
        [pscustomobject]@{ Path='HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\TypedPaths'; Mode='Values'; Ignore=@(); Note='Safe user typed file path history only'; Category='Registry - Typed Paths'; Preview='Typed Paths' },
        [pscustomobject]@{ Path='HKCU:\Software\Microsoft\Internet Explorer\TypedURLs'; Mode='Values'; Ignore=@(); Note='Safe user typed URL history only'; Category='Registry - Typed URLs'; Preview='Typed URLs' },
        [pscustomobject]@{ Path='HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\WordWheelQuery'; Mode='Mixed'; Ignore=@(); Note='Safe Windows Explorer search history only'; Category='Registry - Explorer Search'; Preview='Explorer Search' },
        [pscustomobject]@{ Path='HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs'; Mode='Mixed'; Ignore=@('MRUListEx'); Note='Safe recent documents history only'; Category='Registry - Recent Docs'; Preview='Recent Docs' }
    )
}

function Get-KobraRegistryValueCandidates {
    param([Parameter(Mandatory)][psobject]$Definition)

    if (-not (Test-Path -LiteralPath $Definition.Path)) { return @() }
    $item = Get-ItemProperty -LiteralPath $Definition.Path -ErrorAction SilentlyContinue
    if ($null -eq $item) { return @() }

    $ignored = @('PSPath','PSParentPath','PSChildName','PSDrive','PSProvider','(default)') + @($Definition.Ignore)
    $candidates = @()
    foreach ($property in $item.PSObject.Properties) {
        if ($ignored -contains $property.Name) { continue }
        if ($property.Name -like 'PS*') { continue }
        $valueText = if ($null -eq $property.Value) { '' } else { [string]$property.Value }
        $sizeBytes = [Math]::Max(192, ($property.Name.Length + $valueText.Length) * 2)
        $candidates += [pscustomobject]@{
            Category  = $Definition.Category
            Items     = 1
            SizeBytes = [int64]$sizeBytes
            Path      = ('{0} [{1}]' -f $Definition.Path, $property.Name)
            Note      = $Definition.Note
            Preview   = $Definition.Preview
        }
    }
    return $candidates
}

function Get-KobraRegistryCandidates {
    $all = @()
    foreach ($definition in @(Get-KobraRegistryTraceDefinitions)) {
        if (Test-Path -LiteralPath $definition.Path) {
            Write-KobraDebug -Message ('Registry scan found key: ' + $definition.Path)
        }
        else {
            Write-KobraDebug -Message ('Registry scan skipped missing key: ' + $definition.Path)
        }
        $all += @(Get-KobraRegistryValueCandidates -Definition $definition)
        if ($definition.Mode -eq 'Mixed' -and (Test-Path -LiteralPath $definition.Path)) {
            foreach ($subKey in @(Get-ChildItem -LiteralPath $definition.Path -ErrorAction SilentlyContinue)) {
                $all += [pscustomobject]@{
                    Category  = $definition.Category
                    Items     = 1
                    SizeBytes = [int64]512
                    Path      = $subKey.PSPath
                    Note      = $definition.Note
                    Preview   = $definition.Preview
                }
            }
        }
    }
    return $all
}

function Get-KobraRegistryPreview {
    if (-not (Test-KobraRegistryCleanupSelected)) { return @() }
    $candidates = @(Get-KobraRegistryCandidates)
    if ($candidates.Count -eq 0) { return @() }

    $preview = @()
    foreach ($group in @($candidates | Group-Object Category)) {
        $groupItems = @($group.Group)
        $firstItem = $groupItems | Select-Object -First 1
        $note = if ($null -ne $firstItem -and ($firstItem.PSObject.Properties.Name -contains 'Note')) { [string]$firstItem.Note } else { 'Safe user traces only' }
        $preview += [pscustomobject]@{
            Name      = $group.Name
            Items     = $groupItems.Count
            SizeBytes = Get-KobraSafeMeasureSum -Items $groupItems -PropertyName 'SizeBytes'
            SizeMB    = [Math]::Round(((Get-KobraSafeMeasureSum -Items $groupItems -PropertyName 'SizeBytes') / 1MB), 4)
            Note      = $note
        }
    }
    return $preview
}

function Test-KobraRegistryBackupCoversDefinition {
    param(
        [Parameter(Mandatory)][psobject]$BackupInfo,
        [Parameter(Mandatory)][psobject]$Definition
    )

    if ($null -eq $BackupInfo) {
        return $false
    }

    if (-not ($BackupInfo.PSObject.Properties.Name -contains 'SuccessfulPaths')) {
        return $false
    }

    return (@($BackupInfo.SuccessfulPaths) -contains $Definition.Path)
}

function Invoke-KobraRegistryTraceCleanup {

    param(
        [scriptblock]$Log,
        [psobject]$BackupInfo
    )

    $results = @()

    foreach ($definition in @(Get-KobraRegistryTraceDefinitions)) {
        $foundCount = 0
        $attemptedCount = 0
        $removedCount = 0
        $skippedCount = 0
        $failedCount = 0
        $lockedCount = 0

        if (-not (Test-KobraRegistryBackupCoversDefinition -BackupInfo $BackupInfo -Definition $definition)) {
            $results += @(New-KobraExecutionCategoryResult -Section 'Registry' -Category $definition.Category -SkippedCount 1 -Reason 'no backup')
            if ($Log) { & $Log ("Registry clean skipped: {0} - no backup available" -f $definition.Path) }
            continue
        }

        if (-not (Test-Path -LiteralPath $definition.Path)) {
            $results += @(New-KobraExecutionCategoryResult -Section 'Registry' -Category $definition.Category -SkippedCount 1 -Reason 'missing key')
            if ($Log) { & $Log ("Registry clean skipped: {0} - missing key" -f $definition.Path) }
            continue
        }

        if ($Log) { & $Log ("Registry key found for cleanup: {0}" -f $definition.Path) }

        $propertyCandidates = @(Get-KobraRegistryValueCandidates -Definition $definition)
        foreach ($candidate in $propertyCandidates) {
            if ($candidate.Path -match '\[(.+)\]$') {
                $valueName = $matches[1]
                $foundCount++
                $attemptedCount++
                try {
                    Remove-ItemProperty -LiteralPath $definition.Path -Name $valueName -ErrorAction Stop
                    $removedCount++
                    if ($Log) { & $Log ("Registry key cleaned: {0} [{1}]" -f $definition.Path, $valueName) }
                }
                catch {
                    $failedCount++
                    if ($Log) { & $Log ("Registry clean failed: {0} [{1}] - {2}" -f $definition.Path, $valueName, $_.Exception.Message) }
                }
            }
        }

        if ($definition.Mode -eq 'Mixed') {
            foreach ($subKey in @(Get-ChildItem -LiteralPath $definition.Path -ErrorAction SilentlyContinue)) {
                $foundCount++
                $attemptedCount++
                try {
                    Remove-Item -LiteralPath $subKey.PSPath -Recurse -Force -ErrorAction Stop
                    $removedCount++
                    if ($Log) { & $Log ("Registry key cleaned: {0}" -f $subKey.PSPath) }
                }
                catch {
                    $failedCount++
                    if ($Log) { & $Log ("Registry clean failed: {0} - {1}" -f $subKey.PSPath, $_.Exception.Message) }
                }
            }
        }

        if (($foundCount -eq 0) -and ($attemptedCount -eq 0)) {
            $skippedCount++
        }

        $results += @(New-KobraExecutionCategoryResult -Section 'Registry' -Category $definition.Category -FoundCount $foundCount -AttemptedCount $attemptedCount -RemovedCount $removedCount -SkippedCount $skippedCount -FailedCount $failedCount -LockedCount $lockedCount)
    }

    if ($Log) { & $Log 'Registry trace cleanup complete.' }
    return $results
}

function Set-KobraQuickScanPreset {
    $script:ChkUserTemp.IsChecked    = $true
    $script:ChkSystemTemp.IsChecked  = $true
    $script:ChkWinUpdate.IsChecked   = $true
    $script:ChkThumbCache.IsChecked  = $true
    $script:ChkShaderCache.IsChecked = $false
    $script:ChkRecycleBin.IsChecked  = $true

    $script:ChkChrome.IsChecked         = $true
    $script:ChkEdge.IsChecked           = $true
    $script:ChkFirefox.IsChecked        = $true
    if ($null -ne $script:ChkOpera) { $script:ChkOpera.IsChecked = $true }
    if ($null -ne $script:ChkBrave) { $script:ChkBrave.IsChecked = $true }
    if ($null -ne $script:ChkVivaldi) { $script:ChkVivaldi.IsChecked = $true }
    $script:ChkBrowserCookies.IsChecked = $false
    if ($null -ne $script:ChkBrowserHistory) { $script:ChkBrowserHistory.IsChecked = $true }
    if ($null -ne $script:ChkRegistryBackup) { $script:ChkRegistryBackup.IsChecked = $true }

    Update-KobraSelectionSummary
    Update-KobraDashboard
}

function Get-KobraSelectedDnsProfile {
    switch ($script:CmbDnsProvider.SelectedIndex) {
        0 { return 'Cloudflare' }
        1 { return 'Google' }
        2 { return 'Automatic' }
        default { return 'Cloudflare' }
    }
}


function Get-KobraResultNote {
    param(
        [string]$Category,
        [string]$ExistingNote = ''
    )

    if (-not [string]::IsNullOrWhiteSpace($ExistingNote)) {
        return $ExistingNote
    }

    switch -Wildcard ($Category) {
        '*Shader Cache' { return 'Rebuilds on next game launch' }
        '*Windows Update Cache' { return 'Windows may rebuild cache later' }
        '*Cookies' { return 'May sign you out of websites' }
        '*History' { return 'Browsing and download history where safe to remove' }
        '*Chrome Cache' { return 'Bookmarks and saved passwords are preserved' }
        '*Edge Cache' { return 'Bookmarks and saved passwords are preserved' }
        '*Firefox Cache' { return 'Bookmarks and saved passwords are preserved' }
        '*Opera Cache' { return 'Bookmarks and saved passwords are preserved' }
        '*Brave* Cache' { return 'Bookmarks and saved passwords are preserved' }
        '*Vivaldi Cache' { return 'Bookmarks and saved passwords are preserved' }
        'Registry -*' { return 'Safe user traces only; grouped for review before cleaning' }
        default { return '' }
    }
}


function Get-KobraSectionPlan {
    param(
        [ValidateSet('System','Browser','Registry')]
        [string]$Scope
    )

    switch ($Scope) {
        'System' {
            return [pscustomobject]@{
                Scope                = 'System'
                CleanupTargets       = @(Get-KobraCleanupTargetsFromUi)
                BrowserTargets       = @()
                BrowserComponents    = @('Cache')
                IncludeRegistryTraces = $false
                CreateRestorePoint   = ($null -ne $script:ChkSystemRestorePoint -and [bool]$script:ChkSystemRestorePoint.IsChecked)
                CreateBackupBundle   = $false
                CreateRegistryBackup = $false
                Label                = 'System cleanup'
            }
        }
        'Browser' {
            return [pscustomobject]@{
                Scope                = 'Browser'
                CleanupTargets       = @()
                BrowserTargets       = @(Get-KobraBrowserTargetsFromUi)
                BrowserComponents    = @(Get-KobraBrowserComponentsFromUi)
                IncludeRegistryTraces = $false
                CreateRestorePoint   = $false
                CreateBackupBundle   = ($null -ne $script:ChkBrowserBackupBundle -and [bool]$script:ChkBrowserBackupBundle.IsChecked)
                CreateRegistryBackup = $false
                Label                = 'Browser cleanup'
            }
        }
        'Registry' {
            return [pscustomobject]@{
                Scope                = 'Registry'
                CleanupTargets       = @()
                BrowserTargets       = @()
                BrowserComponents    = @('Cache')
                IncludeRegistryTraces = (Test-KobraRegistryCleanupSelected)
                CreateRestorePoint   = $false
                CreateBackupBundle   = $false
                CreateRegistryBackup = (Test-KobraRegistryBackupSelected)
                Label                = 'Registry cleanup'
            }
        }
    }
}

function Reset-KobraSectionReadiness {
    $script:SectionScanReady['System'] = $false
    $script:SectionScanReady['Browser'] = $false
    $script:SectionScanReady['Registry'] = $false
    Set-KobraSectionActionState
}

function Set-KobraSectionActionState {
    if ($null -ne $script:BtnSystemClean) { $script:BtnSystemClean.IsEnabled = [bool]$script:SectionScanReady['System'] }
    if ($null -ne $script:BtnBrowserClean) { $script:BtnBrowserClean.IsEnabled = [bool]$script:SectionScanReady['Browser'] }
    if ($null -ne $script:BtnRegistryClean) { $script:BtnRegistryClean.IsEnabled = [bool]$script:SectionScanReady['Registry'] }
}

function Invoke-KobraScopeAnalyze {
    param(
        [ValidateSet('System','Browser','Registry')]
        [string]$Scope
    )

    $plan = Get-KobraSectionPlan -Scope $Scope
    $cleanupTargets = @($plan.CleanupTargets)
    $browserTargets = @($plan.BrowserTargets)
    $browserComponents = @($plan.BrowserComponents)
    $includeRegistry = [bool]$plan.IncludeRegistryTraces

    if ((Get-KobraSafeCount $cleanupTargets) -eq 0 -and (Get-KobraSafeCount $browserTargets) -eq 0 -and -not $includeRegistry) {
        Write-KobraUiLog -Message ("Nothing selected in {0}. Choose at least one option first." -f $plan.Label) -BlankLine
        return
    }

    Set-KobraButtonsEnabled -Enabled $false
    Show-KobraOperationView -Status (("Scanning {0}") -f $plan.Label) -Detail 'Preparing the selected area for review.' -OriginView 'CustomClean' -Value 8

    try {
        Set-KobraAnalyzeStatus -Title ("Analyzing {0}" -f $plan.Label.ToLower()) -SubTitle ("Scanning only the options selected in {0}." -f $plan.Label.ToLower())
        Write-KobraUiLog -Message (("Starting {0} scan...") -f $plan.Label) -BlankLine

        $cleanupPreview = @()
        if ((Get-KobraSafeCount $cleanupTargets) -gt 0) {
            $cleanupPreview = @(Get-KobraCleanupPreview -Targets $cleanupTargets)
        }
        $browserPreview = @()
        if ((Get-KobraSafeCount $browserTargets) -gt 0) {
            $browserPreview = @(Get-KobraBrowserPreview -Browsers $browserTargets -Components $browserComponents)
        }
        $registryPreview = @()
        if ($includeRegistry) {
            $registryPreview = @(Get-KobraRegistryPreview)
            Write-KobraDebug -Message ('Analyze registryPreviewCount=' + (@($registryPreview).Count))
        }

        [int64]$totalBytes = 0

        if ((Get-KobraSafeCount $cleanupPreview) -gt 0) {
            Update-KobraOperationView -Status (("Scanning {0}") -f $plan.Label) -Detail 'Inspecting system junk, temp files, and cache locations.' -Value 28
            Write-KobraUiLog -Message 'Cleanup targets:' -NoTimestamp
            foreach ($item in $cleanupPreview) {
                if ($null -eq $item) { continue }
                $totalBytes += Get-KobraSafeInt64 $item.SizeBytes
                Write-KobraUiLog -Message (("  {0}: {1:N2} MB ({2} items)") -f $item.Name, [double]$item.SizeMB, (Get-KobraSafeInt64 $item.Items)) -NoTimestamp
            }
        }

        if ((Get-KobraSafeCount $browserPreview) -gt 0) {
            Update-KobraOperationView -Status (("Scanning {0}") -f $plan.Label) -Detail 'Counting browser cache, cookie, and history targets.' -Value 55
            Write-KobraUiLog -Message 'Browser targets:' -NoTimestamp
            foreach ($item in $browserPreview) {
                if ($null -eq $item) { continue }
                $totalBytes += Get-KobraSafeInt64 $item.SizeBytes
                Write-KobraUiLog -Message (("  {0}: {1:N2} MB ({2:N0} records)") -f $item.Name, [double]$item.SizeMB, (Get-KobraSafeInt64 $item.Items)) -NoTimestamp
            }
            Write-KobraUiLog -Message 'Saved passwords and bookmarks are preserved. Cookie cleaning may sign you out of sites.' -NoTimestamp
        }

        if ((Get-KobraSafeCount $registryPreview) -gt 0) {
            Update-KobraOperationView -Status (("Scanning {0}") -f $plan.Label) -Detail 'Reviewing safe MRU and typed-history registry traces.' -Value 72
            Write-KobraUiLog -Message 'Registry targets:' -NoTimestamp
            foreach ($item in $registryPreview) {
                if ($null -eq $item) { continue }
                $totalBytes += Get-KobraSafeInt64 $item.SizeBytes
                Write-KobraUiLog -Message (("  {0}: {1:N2} MB ({2:N0} records)") -f $item.Name, [double]$item.SizeMB, (Get-KobraSafeInt64 $item.Items)) -NoTimestamp
            }
        }

        $manifest = New-KobraDeleteManifest -CleanupTargets $cleanupTargets -BrowserTargets $browserTargets -BrowserComponents $browserComponents -IncludeRegistryTraces $includeRegistry
        Write-KobraDebug -Message ('Analyze manifest candidateCount=' + $manifest.CandidateCount + '; totalBytes=' + $manifest.TotalBytes + '; manifest=' + $manifest.ManifestPath)
        foreach ($item in @($manifest.CategorySummary)) { try { Write-KobraDebug -Message ('Analyze category=' + $item.Category + '; items=' + $item.Items + '; bytes=' + $item.SizeBytes) } catch {} }
        $script:LastAnalyzeManifest = $manifest
        $script:LastCleanupSummary = $null
        $script:LastAnalyzeTime = Get-Date
        $script:LastAnalyzeSelectionSignature = Get-KobraCurrentAnalyzeSelectionSignature -Scope $Scope
        $script:LastAnalyzeScope = $Scope
        $script:CurrentResultsScope = $Scope
        $script:SectionScanReady[$Scope] = ($manifest.CandidateCount -gt 0)
        Set-KobraSectionActionState
        Update-KobraResultsPanel -CategorySummary $manifest.CategorySummary -TotalRecords $manifest.CandidateCount -TotalBytes $manifest.TotalBytes
        Update-KobraLastActionStatus -Status 'Scan completed'
        Write-KobraUiLog -Message (("Delete manifest written: {0}") -f $manifest.ManifestPath) -NoTimestamp
        Write-KobraUiLog -Message (("Estimated reclaimable space: {0:N2} MB") -f ($totalBytes / 1MB))
        Write-KobraUiLog -Message (("Estimated removable records: {0:N0}") -f $manifest.CandidateCount)
        Set-KobraAnalyzeStatus -Title ("{0} scan complete" -f $plan.Label) -SubTitle 'Your results are ready to review or clean.'
        Complete-KobraOperationView -Status (("{0} scan complete" -f $plan.Label)) -Detail 'Your results are ready to review or clean.' -NextView 'Results'
    }
    catch {
        Write-KobraDebug -Message ('Analysis exception: ' + $_.Exception.Message)
        try { Write-KobraDebug -Message ('Analysis scriptstack: ' + $_.ScriptStackTrace) } catch {}
        Set-KobraAnalyzeStatus -Title 'Analysis failed' -SubTitle $_.Exception.Message
        Write-KobraUiLog -Message (("Analysis failed: {0}") -f $_.Exception.Message)
        Update-KobraOperationView -Status 'Analysis failed' -Detail $_.Exception.Message -Value 0
    }
    finally {
        Set-KobraButtonsEnabled -Enabled $true
    }
}

function Invoke-KobraScopeClean {
    param(
        [ValidateSet('System','Browser','Registry')]
        [string]$Scope
    )

    $plan = Get-KobraSectionPlan -Scope $Scope
    $cleanupTargets = @($plan.CleanupTargets)
    $browserTargets = @($plan.BrowserTargets)
    $browserComponents = @($plan.BrowserComponents)
    $includeRegistry = [bool]$plan.IncludeRegistryTraces

    if ((Get-KobraSafeCount $cleanupTargets) -eq 0 -and (Get-KobraSafeCount $browserTargets) -eq 0 -and -not $includeRegistry) {
        Write-KobraUiLog -Message (("Nothing selected in {0}. Choose at least one option first.") -f $plan.Label) -BlankLine
        return
    }

    Set-KobraButtonsEnabled -Enabled $false
    Show-KobraOperationView -Status (("Cleaning {0}") -f $plan.Label) -Detail 'Preparing the selected area for cleanup.' -OriginView 'CustomClean' -Value 5
    try {
        if (-not (Test-KobraAnalyzeManifestMatchesCurrentSelection -Scope $Scope)) {
            Write-KobraUiLog -Message (("Rescan required before {0}. The current selections no longer match the last scan manifest.") -f $plan.Label) -BlankLine
            Update-KobraOperationView -Status (("Cleaning {0}") -f $plan.Label) -Detail 'Rescan this section before cleaning.' -Value 0
            return
        }

        $stepCount = 0
        if ($plan.CreateRestorePoint) { $stepCount++ }
        if ($plan.CreateBackupBundle) { $stepCount++ }
        if ($plan.CreateRegistryBackup) { $stepCount++ }
        if ((Get-KobraSafeCount $cleanupTargets) -gt 0) { $stepCount++ }
        if ((Get-KobraSafeCount $browserTargets) -gt 0) { $stepCount++ }
        if ($includeRegistry) { $stepCount++ }
        if ($stepCount -eq 0) {
            Write-KobraUiLog -Message (("Nothing to clean for {0}.") -f $plan.Label) -BlankLine
            Set-KobraProgress 0
            return
        }

        $manifest = $script:LastAnalyzeManifest
        $confirmed = Show-KobraExecutionConfirmation -CleanupTargets $cleanupTargets -BrowserTargets $browserTargets -BrowserComponents $browserComponents -DoTls $false -DoNetwork $false -DoDnsFlush $false -DoDnsProfile $false -DoHpDebloat $false -DoRestore $plan.CreateRestorePoint -DoRegistryTraces $includeRegistry -DnsProfile '' -ManifestInfo $manifest
        if (-not $confirmed) {
            Write-KobraUiLog -Message (("{0} cleanup canceled by user.") -f $plan.Label) -BlankLine
            Set-KobraProgress 0
            return
        }

        $currentStep = 0
        $executionResults = @()
        $registryBackup = $script:LastRegistryBackup
        Write-KobraUiLog -Message (("Executing {0}...") -f $plan.Label) -BlankLine
        Write-KobraUiLog -Message (("Delete manifest: {0}") -f $manifest.ManifestPath)
        Update-KobraOperationView -Status (("Cleaning {0}") -f $plan.Label) -Detail 'Starting the selected cleanup tasks.' -Value 8

        if ($plan.CreateRestorePoint) {
            Update-KobraOperationView -Status (("Cleaning {0}") -f $plan.Label) -Detail 'Creating a restore point before cleanup.'
            Write-KobraUiLog -Message 'Creating restore point before system cleanup...'
            Invoke-KobraGuard -Log ${function:Write-KobraUiLog}
            $currentStep++
            Set-KobraProgress ([math]::Round(($currentStep / $stepCount) * 100))
        }
        if ($plan.CreateBackupBundle) {
            Update-KobraOperationView -Status (("Cleaning {0}") -f $plan.Label) -Detail 'Creating a system backup bundle before browser cleanup.'
            Write-KobraUiLog -Message 'Creating system backup bundle before browser cleanup...'
            $backup = Invoke-KobraSettingsBackup -BackupRoot $script:BackupRoot -Log ${function:Write-KobraUiLog}
            $currentStep++
            Set-KobraProgress ([math]::Round(($currentStep / $stepCount) * 100))
        }
        if ($plan.CreateRegistryBackup) {
            Update-KobraOperationView -Status 'Preparing registry backup' -Detail 'Exporting registry backups before cleanup.'
            Write-KobraUiLog -Message 'Creating registry backup bundle before registry cleanup...'
            $registryBackup = New-KobraRegistryTraceBackup -Log ${function:Write-KobraUiLog}
            $script:LastRegistryBackup = $registryBackup
            Update-KobraRegistryBackupStatusUi
            try { Write-KobraDebug -Message ('Registry backup created: exportCount=' + $registryBackup.ExportCount + '; backupPath=' + $registryBackup.BackupPath) } catch {}
            $currentStep++
            Set-KobraProgress ([math]::Round(($currentStep / $stepCount) * 100))
        }
        if ((Get-KobraSafeCount $cleanupTargets) -gt 0) {
            Update-KobraOperationView -Status 'Cleaning selected items' -Detail 'Removing selected system cleanup targets.'
            $executionResults += @(Invoke-KobraShed -Targets $cleanupTargets -Log ${function:Write-KobraUiLog})
            $currentStep++
            Set-KobraProgress ([math]::Round(($currentStep / $stepCount) * 100))
        }
        if ((Get-KobraSafeCount $browserTargets) -gt 0) {
            Update-KobraOperationView -Status 'Cleaning selected items' -Detail 'Removing selected browser cache data.'
            $executionResults += @(Invoke-KobraBrowserCleanup -Browsers $browserTargets -Components $browserComponents -Log ${function:Write-KobraUiLog})
            $currentStep++
            Set-KobraProgress ([math]::Round(($currentStep / $stepCount) * 100))
        }
        if ($includeRegistry) {
            Update-KobraOperationView -Status 'Cleaning selected items' -Detail 'Cleaning the selected registry trace groups.'
            $executionResults += @(Invoke-KobraRegistryTraceCleanup -Log ${function:Write-KobraUiLog} -BackupInfo $registryBackup)
            $currentStep++
            Set-KobraProgress ([math]::Round(($currentStep / $stepCount) * 100))
        }

        Write-KobraUiLog -Message (("{0} complete.") -f $plan.Label)
        Show-KobraCleanupSummary -ExecutionResults $executionResults -ScopeLabel $plan.Label
        Complete-KobraOperationView -Status (("{0} complete" -f $plan.Label)) -Detail 'Cleanup finished. Review the cleanup summary.' -NextView 'Results'
        $script:SectionScanReady[$Scope] = $false
        Set-KobraSectionActionState
        $script:CurrentResultsScope = 'All'
    }
    catch {
        Write-KobraUiLog -Message (("{0} failed: {1}") -f $plan.Label, $_.Exception.Message)
        Update-KobraOperationView -Status (("{0} failed" -f $plan.Label)) -Detail $_.Exception.Message -Value 0
    }
    finally {
        Set-KobraButtonsEnabled -Enabled $true
    }
}

function Update-KobraResultsPanel {
    param(
        [object[]]$CategorySummary,
        [int]$TotalRecords = 0,
        [int64]$TotalBytes = 0
    )

    $bindingList = New-Object System.Collections.ArrayList
    foreach ($item in @($CategorySummary | Sort-Object SizeBytes -Descending)) {
        if ($null -eq $item) { continue }
        [void]$bindingList.Add([pscustomobject]@{
            Category = $item.Category
            ItemsText = ('{0:N0}' -f $item.Items)
            SizeText  = ('{0:N2} MB' -f ($item.SizeBytes / 1MB))
            Note      = (Get-KobraResultNote -Category $item.Category -ExistingNote ($item.Note -as [string]))
        })
    }

    $script:ResultsList.ItemsSource = $null
    $script:ResultsList.ItemsSource = $bindingList
    $script:ResultsMode = 'ScanResults'

    if ($TotalRecords -gt 0) {
        $script:TxtResultsHeadline.Text = ('{0:N0} records can be removed' -f $TotalRecords)
        $script:TxtResultsSubHeadline.Text = ('Estimated reclaimable space: {0:N2} MB across {1} categories.' -f ($TotalBytes / 1MB), $bindingList.Count)
    }
    else {
        $script:TxtResultsHeadline.Text = 'No removable records detected'
        $script:TxtResultsSubHeadline.Text = 'Analyze scanned the current selections but did not find removable records.'
    }

    if ($null -ne $script:TxtResultsMode) {
        $script:TxtResultsMode.Text = 'Scan completed'
    }
    if ($null -ne $script:BtnRunSelected) {
        $script:BtnRunSelected.Visibility = 'Visible'
        $script:BtnRunSelected.IsEnabled = ($TotalRecords -gt 0)
        $script:BtnRunSelected.Content = 'Clean Selected'
    }
    if ($null -ne $script:BtnResultsRescan) { $script:BtnResultsRescan.Visibility = 'Visible' }
    if ($null -ne $script:BtnResultsBackCustom) { $script:BtnResultsBackCustom.Visibility = 'Visible' }
    if ($null -ne $script:BtnResultsOpenDebugLog) { $script:BtnResultsOpenDebugLog.Visibility = 'Visible' }

    Update-KobraSelectionSummary
    Update-KobraDashboard
    Invoke-KobraUiRefresh
}

function Show-KobraCleanupSummary {
    param(
        [Parameter(Mandatory)][object[]]$ExecutionResults,
        [string]$ScopeLabel = 'Selected tasks'
    )

    $bindingList = New-Object System.Collections.ArrayList
    $sectionMap = [ordered]@{
        System   = @{ Found = 0; Removed = 0; Skipped = 0; Failed = 0; Locked = 0; RemovedBytes = [int64]0 }
        Browser  = @{ Found = 0; Removed = 0; Skipped = 0; Failed = 0; Locked = 0; RemovedBytes = [int64]0 }
        Registry = @{ Found = 0; Removed = 0; Skipped = 0; Failed = 0; Locked = 0; RemovedBytes = [int64]0 }
    }

    $cleanedItems = 0
    $skippedItems = 0
    $failedItems = 0
    $lockedItems = 0
    [int64]$bytesFreed = 0

    foreach ($item in @($ExecutionResults)) {
        if ($null -eq $item -or [string]::IsNullOrWhiteSpace(($item.Category -as [string]))) { continue }

        $section = if ($sectionMap.Contains($item.Section)) { $item.Section } else { 'System' }
        $sectionMap[$section].Found += [int](Get-KobraSafeInt64 $item.FoundCount)
        $sectionMap[$section].Removed += [int](Get-KobraSafeInt64 $item.RemovedCount)
        $sectionMap[$section].Skipped += [int](Get-KobraSafeInt64 $item.SkippedCount)
        $sectionMap[$section].Failed += [int](Get-KobraSafeInt64 $item.FailedCount)
        $sectionMap[$section].Locked += [int](Get-KobraSafeInt64 $item.LockedCount)
        $sectionMap[$section].RemovedBytes += Get-KobraSafeInt64 $item.RemovedBytes

        $cleanedItems += [int](Get-KobraSafeInt64 $item.RemovedCount)
        $skippedItems += [int](Get-KobraSafeInt64 $item.SkippedCount)
        $failedItems += [int](Get-KobraSafeInt64 $item.FailedCount)
        $lockedItems += [int](Get-KobraSafeInt64 $item.LockedCount)
        $bytesFreed += Get-KobraSafeInt64 $item.RemovedBytes

        $noteParts = @()
        if ((Get-KobraSafeInt64 $item.RemovedCount) -gt 0) { $noteParts += ('cleaned {0:N0}' -f $item.RemovedCount) }
        if ((Get-KobraSafeInt64 $item.SkippedCount) -gt 0) { $noteParts += ('skipped {0:N0}' -f $item.SkippedCount) }
        if ((Get-KobraSafeInt64 $item.LockedCount) -gt 0) { $noteParts += ('locked {0:N0}' -f $item.LockedCount) }
        if ((Get-KobraSafeInt64 $item.FailedCount) -gt 0) { $noteParts += ('failed {0:N0}' -f $item.FailedCount) }
        if (-not [string]::IsNullOrWhiteSpace(($item.Reason -as [string]))) { $noteParts += ($item.Reason -as [string]) }
        if ($noteParts.Count -eq 0) { $noteParts += 'no changes needed' }

        [void]$bindingList.Add([pscustomobject]@{
            Category = $item.Category
            ItemsText = ('{0:N0}' -f ((Get-KobraSafeInt64 $item.RemovedCount) + (Get-KobraSafeInt64 $item.SkippedCount) + (Get-KobraSafeInt64 $item.FailedCount)))
            SizeText = ('{0:N2} MB' -f ((Get-KobraSafeInt64 $item.RemovedBytes) / 1MB))
            Note = ($noteParts -join '; ')
            SortBytes = (Get-KobraSafeInt64 $item.RemovedBytes)
        })
    }

    $sortedRows = @($bindingList | Sort-Object -Property SortBytes, Category -Descending)
    $resultsBinding = New-Object System.Collections.ArrayList
    foreach ($row in $sortedRows) {
        [void]$resultsBinding.Add([pscustomobject]@{
            Category = $row.Category
            ItemsText = $row.ItemsText
            SizeText = $row.SizeText
            Note = $row.Note
        })
    }

    $script:ResultsList.ItemsSource = $null
    $script:ResultsList.ItemsSource = $resultsBinding
    $script:ResultsMode = 'CleanupSummary'

    $statusText = if ($failedItems -gt 0) {
        'Cleanup partially failed'
    }
    elseif ($skippedItems -gt 0 -or $lockedItems -gt 0) {
        'Cleanup completed with skips'
    }
    else {
        'Cleanup completed'
    }

    $script:TxtResultsHeadline.Text = $statusText
    $script:TxtResultsSubHeadline.Text = ('Cleaned: {0:N0} | Skipped: {1:N0} | Failed: {2:N0} | Locked: {3:N0} | Bytes freed: {4:N2} MB' -f $cleanedItems, $skippedItems, $failedItems, $lockedItems, ($bytesFreed / 1MB))
    if ($null -ne $script:TxtResultsMode) {
        $script:TxtResultsMode.Text = $statusText
    }

    if ($null -ne $script:BtnRunSelected) {
        $script:BtnRunSelected.Visibility = 'Collapsed'
    }
    if ($null -ne $script:BtnResultsRescan) { $script:BtnResultsRescan.Visibility = 'Visible' }
    if ($null -ne $script:BtnResultsBackCustom) { $script:BtnResultsBackCustom.Visibility = 'Visible' }
    if ($null -ne $script:BtnResultsOpenDebugLog) { $script:BtnResultsOpenDebugLog.Visibility = 'Visible' }

    $sectionSummaryText = @()
    foreach ($sectionName in $sectionMap.Keys) {
        $section = $sectionMap[$sectionName]
        if (($section.Found + $section.Removed + $section.Skipped + $section.Failed + $section.Locked) -eq 0) { continue }
        $sectionSummaryText += ('{0}: removed {1:N0}, skipped {2:N0}, failed {3:N0}, locked {4:N0}, freed {5:N2} MB' -f $sectionName, $section.Removed, $section.Skipped, $section.Failed, $section.Locked, ($section.RemovedBytes / 1MB))
    }

    $script:LastCleanupSummary = [pscustomobject]@{
        ScopeLabel    = $ScopeLabel
        CleanedItems  = $cleanedItems
        SkippedItems  = $skippedItems
        FailedItems   = $failedItems
        LockedItems   = $lockedItems
        BytesFreed    = $bytesFreed
        SectionTotals = $sectionMap
        Rows          = @($ExecutionResults)
    }

    Update-KobraLastActionStatus -Status $statusText
    Write-KobraDebug -Message ('Cleanup summary scope=' + $ScopeLabel + '; cleaned=' + $cleanedItems + '; skipped=' + $skippedItems + '; failed=' + $failedItems + '; locked=' + $lockedItems + '; bytesFreed=' + $bytesFreed)
    Write-KobraUiLog -Message 'Cleanup summary:' -NoTimestamp
    Write-KobraUiLog -Message (('  Cleaned items: {0:N0}' -f $cleanedItems)) -NoTimestamp
    Write-KobraUiLog -Message (('  Skipped items: {0:N0}' -f $skippedItems)) -NoTimestamp
    Write-KobraUiLog -Message (('  Failed items: {0:N0}' -f $failedItems)) -NoTimestamp
    Write-KobraUiLog -Message (('  Locked items: {0:N0}' -f $lockedItems)) -NoTimestamp
    Write-KobraUiLog -Message (('  Bytes freed: {0:N2} MB' -f ($bytesFreed / 1MB))) -NoTimestamp
    foreach ($line in $sectionSummaryText) {
        Write-KobraUiLog -Message ('  ' + $line) -NoTimestamp
    }

    Update-KobraSelectionSummary
    Update-KobraDashboard
    Invoke-KobraUiRefresh
}

function Resolve-KobraCategorySelected {
    param([string]$Category)

    switch -Wildcard ($Category) {
        'User Temp' { return [bool]$script:ChkUserTemp.IsChecked }
        'System Temp' { return [bool]$script:ChkSystemTemp.IsChecked }
        'Windows Update Cache' { return [bool]$script:ChkWinUpdate.IsChecked }
        'Thumbnail / Icon Cache' { return [bool]$script:ChkThumbCache.IsChecked }
        'DirectX Shader Cache' { return [bool]$script:ChkShaderCache.IsChecked }
        'Recycle Bin' { return [bool]$script:ChkRecycleBin.IsChecked }
        '*Chrome Cache' { return [bool]$script:ChkChrome.IsChecked }
        '*Edge Cache' { return [bool]$script:ChkEdge.IsChecked }
        '*Firefox Cache' { return [bool]$script:ChkFirefox.IsChecked }
        '*Opera* Cache' { return ($null -ne $script:ChkOpera -and [bool]$script:ChkOpera.IsChecked) }
        '*Brave* Cache' { return ($null -ne $script:ChkBrave -and [bool]$script:ChkBrave.IsChecked) }
        '*Vivaldi Cache' { return ($null -ne $script:ChkVivaldi -and [bool]$script:ChkVivaldi.IsChecked) }
        '*Chrome Cookies' { return ([bool]$script:ChkChrome.IsChecked -and [bool]$script:ChkBrowserCookies.IsChecked) }
        '*Edge Cookies' { return ([bool]$script:ChkEdge.IsChecked -and [bool]$script:ChkBrowserCookies.IsChecked) }
        '*Firefox Cookies' { return ([bool]$script:ChkFirefox.IsChecked -and [bool]$script:ChkBrowserCookies.IsChecked) }
        '*Opera* Cookies' { return ($null -ne $script:ChkOpera -and [bool]$script:ChkOpera.IsChecked -and [bool]$script:ChkBrowserCookies.IsChecked) }
        '*Brave* Cookies' { return ($null -ne $script:ChkBrave -and [bool]$script:ChkBrave.IsChecked -and [bool]$script:ChkBrowserCookies.IsChecked) }
        '*Vivaldi Cookies' { return ($null -ne $script:ChkVivaldi -and [bool]$script:ChkVivaldi.IsChecked -and [bool]$script:ChkBrowserCookies.IsChecked) }
        '*Chrome History' { return ([bool]$script:ChkChrome.IsChecked -and $null -ne $script:ChkBrowserHistory -and [bool]$script:ChkBrowserHistory.IsChecked) }
        '*Edge History' { return ([bool]$script:ChkEdge.IsChecked -and $null -ne $script:ChkBrowserHistory -and [bool]$script:ChkBrowserHistory.IsChecked) }
        '*Firefox History' { return ([bool]$script:ChkFirefox.IsChecked -and $null -ne $script:ChkBrowserHistory -and [bool]$script:ChkBrowserHistory.IsChecked) }
        '*Opera* History' { return ($null -ne $script:ChkOpera -and [bool]$script:ChkOpera.IsChecked -and $null -ne $script:ChkBrowserHistory -and [bool]$script:ChkBrowserHistory.IsChecked) }
        '*Brave* History' { return ($null -ne $script:ChkBrave -and [bool]$script:ChkBrave.IsChecked -and $null -ne $script:ChkBrowserHistory -and [bool]$script:ChkBrowserHistory.IsChecked) }
        '*Vivaldi History' { return ($null -ne $script:ChkVivaldi -and [bool]$script:ChkVivaldi.IsChecked -and $null -ne $script:ChkBrowserHistory -and [bool]$script:ChkBrowserHistory.IsChecked) }
        'Registry traces' { return (Test-KobraRegistryCleanupSelected) }
        'Registry -*' { return (Test-KobraRegistryCleanupSelected) }
        default { return $true }
    }
}

function Get-KobraLatestBackupPath {
    if (-not (Test-Path -LiteralPath $script:BackupRoot)) { return $null }
    $latest = Get-ChildItem -LiteralPath $script:BackupRoot -Directory -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($latest) { return $latest.FullName }
    return $null
}

function Update-KobraRecentActivity {
    if ($null -eq $script:TxtRecentActivity -or $null -eq $script:StatusTextBox) { return }
    $lines = @($script:StatusTextBox.Text -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($lines.Count -eq 0) {
        $script:TxtRecentActivity.Text = 'No recent activity yet.'
        return
    }
    $script:TxtRecentActivity.Text = (($lines | Select-Object -Last $script:RecentActivityMax) -join "`r`n")
}

function Update-KobraSelectionSummary {
    if ($null -eq $script:TxtSelectedCategoryCount) { return }

    if ($script:ResultsMode -eq 'CleanupSummary' -and $null -ne $script:LastCleanupSummary) {
        $sectionCount = 0
        foreach ($sectionName in $script:LastCleanupSummary.SectionTotals.Keys) {
            $section = $script:LastCleanupSummary.SectionTotals[$sectionName]
            if (($section.Removed + $section.Skipped + $section.Failed + $section.Locked) -gt 0) {
                $sectionCount++
            }
        }

        $script:TxtSelectedCategoryCount.Text = ('{0} sections summarized' -f $sectionCount)
        $script:TxtSelectedRecordCount.Text = ('{0:N0} cleaned | {1:N0} skipped | {2:N0} failed | {3:N0} locked' -f $script:LastCleanupSummary.CleanedItems, $script:LastCleanupSummary.SkippedItems, $script:LastCleanupSummary.FailedItems, $script:LastCleanupSummary.LockedItems)
        $script:TxtSelectedBytes.Text = ('{0:N2} MB freed' -f ($script:LastCleanupSummary.BytesFreed / 1MB))
        $script:TxtSelectedWarnings.Text = 'Use Rescan to rebuild results after cleanup, or go back to Custom Clean to adjust categories.'
        return
    }

    $selectedCategories = 0
    $selectedRecords = 0
    [int64]$selectedBytes = 0
    $warnings = New-Object System.Collections.Generic.List[string]

    if ($script:LastAnalyzeManifest -and $script:LastAnalyzeManifest.CategorySummary) {
        foreach ($item in @($script:LastAnalyzeManifest.CategorySummary)) {
            if ($null -eq $item) { continue }
            if (Resolve-KobraCategorySelected -Category ($item.Category -as [string])) {
                $selectedCategories++
                $selectedRecords += [int](Get-KobraSafeInt64 $item.Items)
                $selectedBytes += Get-KobraSafeInt64 $item.SizeBytes
                $note = Get-KobraResultNote -Category ($item.Category -as [string]) -ExistingNote ($item.Note -as [string])
                if (-not [string]::IsNullOrWhiteSpace($note) -and -not $warnings.Contains($note)) {
                    $warnings.Add($note)
                }
            }
        }
    }

    if ($selectedCategories -eq 0) {
        $selectedCategories = (@(Get-KobraCleanupTargetsFromUi) + @(Get-KobraBrowserTargetsFromUi)).Count
        if (Test-KobraRegistryCleanupSelected) { $selectedCategories++ }
    }

    $script:TxtSelectedCategoryCount.Text = ('{0} categories selected' -f $selectedCategories)
    $script:TxtSelectedRecordCount.Text = ('{0:N0} records selected' -f $selectedRecords)
    $script:TxtSelectedBytes.Text = ('{0:N2} MB selected' -f ($selectedBytes / 1MB))
    $script:TxtSelectedWarnings.Text = if ($warnings.Count -gt 0) { ($warnings -join ' | ') } else { 'Review the results, then run Clean Selected when you are ready.' }
}

function Update-KobraDashboard {
    if ($null -eq $script:TxtDashLastScan) { return }

    if ($script:LastAnalyzeManifest) {
        $script:TxtDashLastScan.Text = if ($script:LastAnalyzeTime) { $script:LastAnalyzeTime.ToString('MMM d, h:mm tt') } else { 'Moments ago' }
        $script:TxtDashReclaimable.Text = ('{0:N2} MB' -f ($script:LastAnalyzeManifest.TotalBytes / 1MB))
        $script:TxtDashRecords.Text = ('{0:N0}' -f $script:LastAnalyzeManifest.CandidateCount)
        $script:TxtDashCategories.Text = ('{0}' -f (@($script:LastAnalyzeManifest.CategorySummary).Count))
    }
    else {
        $script:TxtDashLastScan.Text = 'No scan yet'
        $script:TxtDashReclaimable.Text = '--'
        $script:TxtDashRecords.Text = '--'
        $script:TxtDashCategories.Text = '--'
    }

    $selectedCount = (@(Get-KobraCleanupTargetsFromUi) + @(Get-KobraBrowserTargetsFromUi)).Count
    if (Test-KobraRegistryCleanupSelected) { $selectedCount++ }
    $script:TxtDashSelections.Text = if ($selectedCount -gt 0) { "$selectedCount active categories are ready for analysis." } else { 'No cleanup categories selected yet.' }

    $latestBackup = Get-KobraLatestBackupPath
    $script:TxtDashSafety.Text = if ($latestBackup) {
        'Latest backup: ' + (Split-Path -Leaf $latestBackup)
    }
    else {
        'No backup bundle yet. Create one before deep system changes.'
    }

    Update-KobraRecentActivity
}

function Set-KobraAnalyzeStatus {
    param(
        [string]$Title,
        [string]$SubTitle = 'Scanning your selected categories and building a reviewable results set.'
    )

    if ($null -ne $script:TxtAnalyzeStage) { $script:TxtAnalyzeStage.Text = $Title }
    if ($null -ne $script:TxtAnalyzeSubStage) { $script:TxtAnalyzeSubStage.Text = $SubTitle }
    Invoke-KobraUiRefresh
}

function Set-KobraNavState {
    param([string]$ActiveView)

    $normalBackground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString('#3B4B67'))
    $normalBorder = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString('#3B4B67'))
    $activeBackground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString('#16E06E'))
    $activeBorder = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString('#16E06E'))
    $activeForeground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString('#102414'))
    $normalForeground = [System.Windows.Media.Brushes]::White

    $navMap = @{
        Dashboard   = $script:BtnNavDashboard
        Analyze     = $script:BtnNavAnalyze
        CustomClean = $script:BtnNavCustomClean
        Results     = $script:BtnNavResults
        OperationProgress = $null
        Tools       = $script:BtnNavTools
        Startup     = $script:BtnNavStartup
        Utilities   = $script:BtnNavUtilities
        About       = $script:BtnNavAbout
    }

    foreach ($entry in $navMap.GetEnumerator()) {
        if ($null -eq $entry.Value) { continue }
        if ($entry.Key -eq $ActiveView) {
            $entry.Value.Background = $activeBackground
            $entry.Value.BorderBrush = $activeBorder
            $entry.Value.Foreground = $activeForeground
        }
        else {
            $entry.Value.Background = $normalBackground
            $entry.Value.BorderBrush = $normalBorder
            $entry.Value.Foreground = $normalForeground
        }
    }
}

function Switch-KobraView {
    param([string]$ViewName)

        $views = @{
        Dashboard   = $script:ViewDashboard
        Analyze     = $script:ViewAnalyze
        CustomClean = $script:ViewCustomClean
        Results     = $script:ViewResults
        OperationProgress = $script:ViewOperationProgress
        Tools       = $script:ViewTools
        Startup     = $script:ViewStartup
        Utilities   = $script:ViewUtilities
        About       = $script:ViewAbout
    }

    foreach ($view in $views.Values) {
        if ($null -ne $view) { $view.Visibility = 'Collapsed' }
    }

    if ($views.ContainsKey($ViewName) -and $null -ne $views[$ViewName]) {
        $views[$ViewName].Visibility = 'Visible'
        $script:CurrentView = $ViewName
    }

    $activeNavView = if ($ViewName -eq 'OperationProgress' -and -not [string]::IsNullOrWhiteSpace($script:OperationOriginView)) { $script:OperationOriginView } else { $ViewName }
    Set-KobraNavState -ActiveView $activeNavView
    Update-KobraDashboard
    Update-KobraSelectionSummary
    Invoke-KobraUiRefresh
}

function Convert-KobraWhiteToTransparent {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [byte]$Threshold = 245
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    try {
        $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
        $bitmap.BeginInit()
        $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bitmap.UriSource = [Uri]$Path
        $bitmap.EndInit()
        $bitmap.Freeze()

        $formatted = New-Object System.Windows.Media.Imaging.FormatConvertedBitmap
        $formatted.BeginInit()
        $formatted.Source = $bitmap
        $formatted.DestinationFormat = [System.Windows.Media.PixelFormats]::Bgra32
        $formatted.EndInit()
        $formatted.Freeze()

        $width  = $formatted.PixelWidth
        $height = $formatted.PixelHeight
        $stride = $width * 4
        $pixels = New-Object byte[] ($stride * $height)
        $formatted.CopyPixels($pixels, $stride, 0)

        for ($i = 0; $i -lt $pixels.Length; $i += 4) {
            $blue  = $pixels[$i]
            $green = $pixels[$i + 1]
            $red   = $pixels[$i + 2]

            if ($red -ge $Threshold -and $green -ge $Threshold -and $blue -ge $Threshold) {
                $pixels[$i + 3] = 0
            }
        }

        $writeable = New-Object System.Windows.Media.Imaging.WriteableBitmap(
            $width,
            $height,
            $formatted.DpiX,
            $formatted.DpiY,
            [System.Windows.Media.PixelFormats]::Bgra32,
            $null
        )

        $rect = New-Object System.Windows.Int32Rect(0, 0, $width, $height)
        $writeable.WritePixels($rect, $pixels, $stride, 0)
        $writeable.Freeze()
        return $writeable
    }
    catch {
        return $null
    }
}

function Get-KobraStartupEntriesForUi {
    param([switch]$IncludeMicrosoft)
    return @(Get-KobraStartupEntries -IncludeMicrosoft:$IncludeMicrosoft)
}

function Refresh-KobraStartupList {
    $includeMs = [bool]$script:ChkStartupShowMicrosoft.IsChecked
    $items = @(Get-KobraStartupEntriesForUi -IncludeMicrosoft:$includeMs)
    $bindingList = New-Object System.Collections.ArrayList
    foreach ($item in $items) {
        [void]$bindingList.Add($item)
    }
    $script:StartupList.ItemsSource = $null
    $script:StartupList.ItemsSource = $bindingList
    Invoke-KobraUiRefresh
    Write-KobraUiLog -Message ("Startup entries loaded: {0}" -f $bindingList.Count)
}

function Invoke-KobraOpenUri {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$FriendlyName = $Uri
    )

    try {
        Start-Process $Uri | Out-Null
        Write-KobraUiLog -Message ("Opened: {0}" -f $FriendlyName)
    }
    catch {
        Write-KobraUiLog -Message ("Could not open {0}: {1}" -f $FriendlyName, $_.Exception.Message)
    }
}

function Invoke-KobraOpenWindowsUpdate { Invoke-KobraOpenUri -Uri 'ms-settings:windowsupdate-action' -FriendlyName 'Windows Update' }
function Invoke-KobraOpenWindowsUpdateSettings { Invoke-KobraOpenUri -Uri 'ms-settings:windowsupdate' -FriendlyName 'Windows Update settings' }
function Invoke-KobraOpenWindowsStorage { Invoke-KobraOpenUri -Uri 'ms-settings:storagesense' -FriendlyName 'Storage settings' }
function Invoke-KobraOpenWindowsApps { Invoke-KobraOpenUri -Uri 'ms-settings:appsfeatures' -FriendlyName 'Apps & Features' }
function Invoke-KobraOpenWindowsStartupSettings { Invoke-KobraOpenUri -Uri 'ms-settings:startupapps' -FriendlyName 'Startup Apps settings' }
function Invoke-KobraOpenWindowsGameMode { Invoke-KobraOpenUri -Uri 'ms-settings:gaming-gamemode' -FriendlyName 'Game Mode settings' }
function Invoke-KobraOpenWindowsGraphics { Invoke-KobraOpenUri -Uri 'ms-settings:display-advancedgraphics' -FriendlyName 'Graphics settings' }
function Invoke-KobraOpenWindowsPower { Invoke-KobraOpenUri -Uri 'ms-settings:powersleep' -FriendlyName 'Power & sleep settings' }

function Invoke-KobraDonate {
    if ([string]::IsNullOrWhiteSpace($script:DonationUrl)) {
        [System.Windows.MessageBox]::Show(
            "Support link not configured yet.`r`n`r`nSet `$script:DonationUrl in Main.ps1 before publishing the public build.",
            'KobraOptimizer - Support Development',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        Write-KobraUiLog -Message 'Support link clicked, but no donation URL is configured yet.'
        return
    }

    Invoke-KobraOpenUri -Uri $script:DonationUrl -FriendlyName 'Donate'
}

function Show-KobraDisclaimer {
    $message = @'
KobraOptimizer can modify caches, startup entries, DNS settings, and other system behavior.

Use it carefully and review the Analyze results before running changes.
Restore points and backup bundles reduce risk, but no warranty is provided.
Save your work and consider a full system backup before applying advanced tweaks.

Please consider donating $1, $5, or whatever you can to help us keep adding new features and support development.
'@

    [System.Windows.MessageBox]::Show(
        $message,
        'KobraOptimizer - Disclaimer',
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    ) | Out-Null
}

function Show-KobraAboutMe {
    $message = @'
KobraOptimizer v{0}

KobraOptimizer is a free Windows utility built to help everyday users clean clutter, manage startup items, review what will be removed before cleanup, create backups before deeper changes, and quickly reach useful Windows performance settings from one place.

Its goal is simple: make PCs easier to maintain, easier to understand, and easier to keep responsive without burying people in confusing menus or risky mystery tweaks.

Key benefits include:
- system cleanup with analysis and delete manifests
- browser cache cleanup that preserves saved logins and bookmarks
- startup management for reducing boot clutter
- backup and restore-point friendly workflow before major changes
- quick access to Windows tools for updates, storage, startup apps, graphics, and power settings

About the creator:
This project was created by an MCSE with more than 30 years of experience in the technology world. Beyond IT and systems work, he enjoys coding, building useful tools, developing ideas into real products, and spending time appreciating nature.

KobraOptimizer is being developed as a practical, community-friendly utility that stays free to use while continuing to grow with new features and improvements over time.
'@ -f $script:AppVersion

    [System.Windows.MessageBox]::Show(
        $message,
        'KobraOptimizer - About Me',
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    ) | Out-Null
}

function Toggle-KobraLogPanel {
    if ($null -eq $script:LogRow) { return }

    $current = $script:LogRow.Height
    if ($current.Value -le 0) {
        $script:LogRow.Height = New-Object System.Windows.GridLength(220)
        $script:BtnToggleLog.Content = 'Hide advanced log'
        Write-KobraUiLog -Message 'Log panel expanded.'
    }
    else {
        $script:LogRow.Height = New-Object System.Windows.GridLength(0)
        $script:BtnToggleLog.Content = 'Show advanced log'
        Write-KobraUiLog -Message 'Log panel collapsed.'
    }
    Invoke-KobraUiRefresh
}

Set-KobraWindowBounds
$fontUsed = Set-KobraPreferredFont

$logoSource = Convert-KobraWhiteToTransparent -Path $script:LogoPath
if ($null -ne $logoSource) {
    $script:KobraLogo.Source = $logoSource
}
elseif (Test-Path -LiteralPath $script:LogoPath) {
    $uri = New-Object Uri($script:LogoPath)
    $script:KobraLogo.Source = New-Object System.Windows.Media.Imaging.BitmapImage($uri)
}

if ($null -ne $script:CmbDnsProvider -and $script:CmbDnsProvider.PSObject.Properties.Name -contains 'SelectedIndex') { $script:CmbDnsProvider.SelectedIndex = 0 }
$script:LogRow.Height = New-Object System.Windows.GridLength(0)
if ($null -ne $script:BtnToggleLog) { $script:BtnToggleLog.Content = 'Show advanced log' }
Update-KobraDnsControls
Update-KobraRegistryBackupStatusUi
$script:StatusTextBox.Clear()
Write-KobraUiLog -Message 'KobraOptimizer ready.'
Write-KobraUiLog -Message ("Version: {0}" -f $script:AppVersion)
Write-KobraUiLog -Message ("Log file: {0}" -f $script:LogFile)
Write-KobraUiLog -Message ("Backup root: {0}" -f $script:BackupRoot)
Write-KobraUiLog -Message ("Manifest root: {0}" -f $script:ManifestRoot)
Write-KobraUiLog -Message ("PowerShell: {0}" -f $PSVersionTable.PSVersion)
Write-KobraUiLog -Message ("Admin: {0}" -f (Test-KobraAdministrator))
Write-KobraUiLog -Message ("Project root: {0}" -f $script:ProjectRoot)
Write-KobraUiLog -Message ("Font: {0}" -f $fontUsed)
Write-KobraUiLog -Message 'Tip: Analyze first to estimate reclaimable space.' -NoTimestamp
Write-KobraUiLog -Message 'Tip: Startup Manager works on Run keys and Startup folders.' -NoTimestamp
Write-KobraUiLog -Message 'Tip: Analyze writes a delete manifest to C:\Temp\KobraOptimizer\Manifests.' -NoTimestamp
Update-KobraResultsPanel -CategorySummary @() -TotalRecords 0 -TotalBytes 0
Update-KobraLastActionStatus -Status 'No scan yet'
Set-KobraAnalyzeStatus -Title 'Quick Scan is ready' -SubTitle 'Run the recommended scan here, or open Custom Clean to fine-tune exactly what Kobra reviews.'

if ($null -ne $script:ChkDnsProfile) { $script:ChkDnsProfile.Add_Click({ Update-KobraDnsControls }) }
if ($null -ne $script:ChkStartupShowMicrosoft) { $script:ChkStartupShowMicrosoft.Add_Click({ Refresh-KobraStartupList }) }
if ($null -ne $script:BtnToggleLog) { $script:BtnToggleLog.Add_Click({ Toggle-KobraLogPanel }) }

$script:BtnNavDashboard.Add_Click({ Switch-KobraView -ViewName 'Dashboard' })
$script:BtnNavAnalyze.Add_Click({ Switch-KobraView -ViewName 'Analyze' })
$script:BtnNavCustomClean.Add_Click({ Switch-KobraView -ViewName 'CustomClean' })
$script:BtnNavResults.Add_Click({ Switch-KobraView -ViewName 'Results' })
$script:BtnNavTools.Add_Click({ Switch-KobraView -ViewName 'Tools' })
$script:BtnNavStartup.Add_Click({ Switch-KobraView -ViewName 'Startup' })
$script:BtnNavUtilities.Add_Click({ Switch-KobraView -ViewName 'Utilities' })
$script:BtnNavAbout.Add_Click({ Switch-KobraView -ViewName 'About' })

$script:BtnDashboardQuickScan.Add_Click({ Switch-KobraView -ViewName 'Analyze' })
$script:BtnDashboardCustomClean.Add_Click({ Switch-KobraView -ViewName 'CustomClean' })
$script:BtnDashboardPerformance.Add_Click({ Switch-KobraView -ViewName 'Tools' })
$script:BtnDashboardStartup.Add_Click({ Switch-KobraView -ViewName 'Startup' })

$script:BtnQuickScanCustomize.Add_Click({ Switch-KobraView -ViewName 'CustomClean' })
$script:BtnQuickScanBackup.Add_Click({ Invoke-KobraButtonClick -Button $script:BtnCreateBackup })
$script:BtnCustomAnalyze.Add_Click({ $script:CurrentResultsScope = 'All'; Switch-KobraView -ViewName 'Analyze'; Invoke-KobraButtonClick -Button $script:BtnAnalyze })
$script:BtnPerformanceApply.Add_Click({ Invoke-KobraButtonClick -Button $script:BtnRunSelected })
$script:BtnPerformanceBackup.Add_Click({ Invoke-KobraButtonClick -Button $script:BtnCreateBackup })
if ($null -ne $script:BtnSystemScan) { $script:BtnSystemScan.Add_Click({ Invoke-KobraScopeAnalyze -Scope 'System' }) }
if ($null -ne $script:BtnSystemClean) { $script:BtnSystemClean.Add_Click({ Invoke-KobraScopeClean -Scope 'System' }) }
if ($null -ne $script:BtnBrowserScan) { $script:BtnBrowserScan.Add_Click({ Invoke-KobraScopeAnalyze -Scope 'Browser' }) }
if ($null -ne $script:BtnBrowserClean) { $script:BtnBrowserClean.Add_Click({ Invoke-KobraScopeClean -Scope 'Browser' }) }
if ($null -ne $script:BtnRegistryScan) { $script:BtnRegistryScan.Add_Click({ Invoke-KobraScopeAnalyze -Scope 'Registry' }) }
if ($null -ne $script:BtnRegistryBackup) {
    $script:BtnRegistryBackup.Add_Click({
        Set-KobraButtonsEnabled -Enabled $false
        Show-KobraOperationView -Status 'Preparing registry backup' -Detail 'Exporting safe registry trace keys.' -OriginView 'CustomClean' -Indeterminate

        try {
            Write-KobraUiLog -Message 'Creating registry backup bundle...' -BlankLine
            $registryBackup = New-KobraRegistryTraceBackup -Log ${function:Write-KobraUiLog}
            $script:LastRegistryBackup = $registryBackup
            Update-KobraRegistryBackupStatusUi

            if ($null -ne $registryBackup) {
                Write-KobraDebug -Message ('Manual registry backup created: exportCount=' + $registryBackup.ExportCount + '; skippedCount=' + $registryBackup.SkippedCount + '; backupPath=' + $registryBackup.BackupPath)
                if (Test-Path -LiteralPath $registryBackup.BackupPath) {
                    Start-Process explorer.exe $registryBackup.BackupPath
                }
                $backupStatus = if ($registryBackup.SkippedCount -gt 0) { 'Registry backup partially completed' } else { 'Registry backup completed' }
                Update-KobraLastActionStatus -Status $backupStatus
                Complete-KobraOperationView -Status 'Registry backup complete' -Detail 'Registry trace backup finished. Missing or unavailable keys were skipped.' -NextView 'CustomClean'
            }
            else {
                Update-KobraLastActionStatus -Status 'No backup yet'
                Complete-KobraOperationView -Status 'Registry backup skipped' -Detail 'No matching registry trace keys exist right now.' -NextView 'CustomClean'
            }
        }
        catch {
            Write-KobraDebug -Message ('Manual registry backup exception: ' + $_.Exception.Message)
            try { Write-KobraDebug -Message ('Manual registry backup scriptstack: ' + $_.ScriptStackTrace) } catch {}
            Write-KobraUiLog -Message ("Registry backup failed: {0}" -f $_.Exception.Message)
            Update-KobraOperationView -Status 'Registry backup failed' -Detail $_.Exception.Message -Value 0
        }
        finally {
            Set-KobraButtonsEnabled -Enabled $true
        }
    })
}
if ($null -ne $script:BtnRegistryClean) { $script:BtnRegistryClean.Add_Click({ Invoke-KobraScopeClean -Scope 'Registry' }) }

$selectionControls = @(
    $script:ChkUserTemp,$script:ChkSystemTemp,$script:ChkWinUpdate,$script:ChkThumbCache,$script:ChkShaderCache,$script:ChkRecycleBin,
$script:ChkChrome,$script:ChkEdge,$script:ChkFirefox,$script:ChkOpera,$script:ChkBrave,$script:ChkVivaldi,$script:ChkBrowserCookies,$script:ChkBrowserHistory,$script:ChkBrowserBackupBundle,$script:ChkSystemRestorePoint,
    $script:ChkRegistryClean,$script:ChkRegistryBackup,
    $script:ChkRestorePoint,$script:ChkRegistry,$script:ChkNetwork,$script:ChkDnsFlush,$script:ChkDnsProfile,$script:ChkHPDebloat
)
foreach ($control in $selectionControls) {
    if ($null -ne $control) {
        $control.Add_Click({ Write-KobraDebug -Message ('Selection changed: ' + (Get-KobraDebugSelectionState)); Update-KobraSelectionSummary; Update-KobraDashboard; Update-KobraDnsControls; Update-KobraRegistryBackupStatusUi; Reset-KobraSectionReadiness })
    }
}

$script:BtnAnalyze.Add_Click({
    Write-KobraDebug -Message 'BtnAnalyze clicked.'
    Write-KobraDebug -Message ('Selections at Analyze: ' + (Get-KobraDebugSelectionState))
    Set-KobraButtonsEnabled -Enabled $false
    $script:CurrentResultsScope = 'All'
    Show-KobraOperationView -Status 'Kobra is analyzing your system...' -Detail 'Preparing the recommended scan path.' -OriginView 'Analyze' -Value 8

    try {
        Set-KobraAnalyzeStatus -Title 'Analyzing your system' -SubTitle 'Reviewing selected cleanup categories and browser data.'
        Write-KobraUiLog -Message 'Starting analysis scan...' -BlankLine

        $cleanupTargets = @(Get-KobraCleanupTargetsFromUi)
        Write-KobraDebug -Message ('RunSelected cleanupTargets=' + $(if ((Get-KobraSafeCount $cleanupTargets) -gt 0) { $cleanupTargets -join ',' } else { '<none>' }))
        Write-KobraDebug -Message ('Analyze cleanupTargets=' + $(if ((Get-KobraSafeCount $cleanupTargets) -gt 0) { $cleanupTargets -join ',' } else { '<none>' }))
        if ((Get-KobraSafeCount $cleanupTargets) -eq 0) {
            $cleanupTargets = @('UserTemp','SystemTemp','WindowsUpdate','ThumbnailCache','RecycleBin')
        }

        $cleanupPreview = @(Get-KobraCleanupPreview -Targets $cleanupTargets)
        $browserTargets = @(Get-KobraBrowserTargetsFromUi)
        Write-KobraDebug -Message ('RunSelected browserTargets=' + $(if ((Get-KobraSafeCount $browserTargets) -gt 0) { $browserTargets -join ',' } else { '<none>' }))
        Write-KobraDebug -Message ('Analyze browserTargets=' + $(if ((Get-KobraSafeCount $browserTargets) -gt 0) { $browserTargets -join ',' } else { '<none>' }))
        $browserComponents = @(Get-KobraBrowserComponentsFromUi)
        Write-KobraDebug -Message ('RunSelected browserComponents=' + $(if ((Get-KobraSafeCount $browserComponents) -gt 0) { $browserComponents -join ',' } else { '<none>' }))
        Write-KobraDebug -Message ('Analyze browserComponents=' + $(if ((Get-KobraSafeCount $browserComponents) -gt 0) { $browserComponents -join ',' } else { '<none>' }))
        $browserPreview = @()
        if ((Get-KobraSafeCount $browserTargets) -gt 0) {
            $browserPreview = @(Get-KobraBrowserPreview -Browsers $browserTargets -Components $browserComponents)
        }
        $registryPreview = @(Get-KobraRegistryPreview)

        $totalBytes = [int64]0

        Set-KobraAnalyzeStatus -Title 'Scanning cleanup targets' -SubTitle 'Inspecting temp files, caches, and recycle locations.'
        Update-KobraOperationView -Status 'Kobra is analyzing your system...' -Detail 'Inspecting system junk, temp files, and recycle locations.' -Value 30
        Write-KobraUiLog -Message 'Cleanup targets:' -NoTimestamp
        foreach ($item in $cleanupPreview) {
            if ($null -eq $item) { continue }
            $totalBytes += Get-KobraSafeInt64 $item.SizeBytes
            Write-KobraUiLog -Message ("  {0}: {1:N2} MB ({2} items)" -f $item.Name, [double]$item.SizeMB, (Get-KobraSafeInt64 $item.Items)) -NoTimestamp
        }

        if ((Get-KobraSafeCount $browserPreview) -gt 0) {
        Set-KobraAnalyzeStatus -Title 'Scanning browser data' -SubTitle 'Counting browser cache, cookie, and history targets.'
        Update-KobraOperationView -Status 'Kobra is analyzing your system...' -Detail 'Counting browser cache, cookie, and history targets.' -Value 58
            Write-KobraUiLog -Message 'Browser targets:' -NoTimestamp
            foreach ($item in $browserPreview) {
                if ($null -eq $item) { continue }
                $totalBytes += Get-KobraSafeInt64 $item.SizeBytes
                Write-KobraUiLog -Message ("  {0}: {1:N2} MB ({2:N0} records)" -f $item.Name, [double]$item.SizeMB, (Get-KobraSafeInt64 $item.Items)) -NoTimestamp
                if (-not [string]::IsNullOrWhiteSpace(($item.Note -as [string]))) {
                    Write-KobraUiLog -Message ("    Note: {0}" -f $item.Note) -NoTimestamp
                }
            }
            Write-KobraUiLog -Message 'Saved passwords and bookmarks are preserved. Cookie cleaning may sign you out of sites.' -NoTimestamp
        }

        if ((Get-KobraSafeCount $registryPreview) -gt 0) {
            Set-KobraAnalyzeStatus -Title 'Scanning registry traces' -SubTitle 'Counting safe MRU and typed-history registry items for review.'
            Update-KobraOperationView -Status 'Kobra is analyzing your system...' -Detail 'Reviewing safe MRU and typed-history registry traces.' -Value 74
            Write-KobraUiLog -Message 'Registry targets:' -NoTimestamp
            foreach ($item in $registryPreview) {
                if ($null -eq $item) { continue }
                $totalBytes += Get-KobraSafeInt64 $item.SizeBytes
                Write-KobraUiLog -Message (("  {0}: {1:N2} MB ({2:N0} records)") -f $item.Name, [double]$item.SizeMB, (Get-KobraSafeInt64 $item.Items)) -NoTimestamp
                if (-not [string]::IsNullOrWhiteSpace(($item.Note -as [string]))) {
                    Write-KobraUiLog -Message (("    Note: {0}") -f $item.Note) -NoTimestamp
                }
            }
        }

        if ($script:ChkDnsProfile.IsChecked) {
            $profileName = Get-KobraSelectedDnsProfile
            $dnsSummary = Get-KobraDnsPlan -ProfileName $profileName
            Write-KobraUiLog -Message ("DNS profile queued: {0}" -f $dnsSummary.DisplayName) -NoTimestamp
            if ((Get-KobraSafeCount $dnsSummary.IPv4) -gt 0) {
                Write-KobraUiLog -Message ("  IPv4: {0}" -f ($dnsSummary.IPv4 -join ', ')) -NoTimestamp
            }
            else {
                Write-KobraUiLog -Message '  IPv4: Automatic (DHCP)' -NoTimestamp
            }
        }

        $manifest = New-KobraDeleteManifest -CleanupTargets $cleanupTargets -BrowserTargets $browserTargets -BrowserComponents $browserComponents -IncludeRegistryTraces (Test-KobraRegistryCleanupSelected)
        Write-KobraDebug -Message ('RunSelected manifest candidateCount=' + $manifest.CandidateCount + '; totalBytes=' + $manifest.TotalBytes + '; manifest=' + $manifest.ManifestPath)
        $script:LastAnalyzeManifest = $manifest
        $script:LastCleanupSummary = $null
        $script:LastAnalyzeTime = Get-Date
        $script:LastAnalyzeSelectionSignature = Get-KobraCurrentAnalyzeSelectionSignature -Scope 'All'
        $script:LastAnalyzeScope = 'All'
        Update-KobraResultsPanel -CategorySummary $manifest.CategorySummary -TotalRecords $manifest.CandidateCount -TotalBytes $manifest.TotalBytes
        Update-KobraLastActionStatus -Status 'Scan completed'
        Write-KobraUiLog -Message ("Delete manifest written: {0}" -f $manifest.ManifestPath) -NoTimestamp
        Write-KobraUiLog -Message ("Estimated reclaimable space: {0:N2} MB" -f ($totalBytes / 1MB))
        Write-KobraUiLog -Message ("Estimated removable records: {0:N0}" -f $manifest.CandidateCount)
        Set-KobraAnalyzeStatus -Title 'Analysis complete' -SubTitle 'Your scan results are ready to review.'
        Reset-KobraSectionReadiness
        Complete-KobraOperationView -Status 'Analysis complete' -Detail 'Your scan results are ready to review.' -NextView 'Results'
    }
    catch {
        Set-KobraAnalyzeStatus -Title 'Analysis failed' -SubTitle $_.Exception.Message
        Write-KobraUiLog -Message ("Analysis failed: {0}" -f $_.Exception.Message)
        Update-KobraOperationView -Status 'Analysis failed' -Detail $_.Exception.Message -Value 0
    }
    finally {
        Set-KobraButtonsEnabled -Enabled $true
    }
})

$script:BtnQuickShed.Add_Click({
    Set-KobraButtonsEnabled -Enabled $false
    Show-KobraOperationView -Status 'Quick Shed in Progress' -Detail 'Preparing the default cleanup path.' -OriginView 'Dashboard' -Indeterminate

    try {
        $defaultTargets = @('UserTemp','SystemTemp','ThumbnailCache','ShaderCache','RecycleBin')
        $manifest = New-KobraDeleteManifest -CleanupTargets $defaultTargets -BrowserTargets @() -BrowserComponents @('Cache')

        $confirmed = Show-KobraExecutionConfirmation -CleanupTargets $defaultTargets -BrowserTargets @() -BrowserComponents @('Cache') -DoTls $false -DoNetwork $false -DoDnsFlush $false -DoDnsProfile $false -DoHpDebloat $false -DoRestore $false -DoRegistryTraces $false -DnsProfile '' -ManifestInfo $manifest
        if (-not $confirmed) {
            Write-KobraUiLog -Message 'Quick Shed canceled by user.' -BlankLine
            Set-KobraProgress 0
            return
        }

        Update-KobraOperationView -Status 'Quick Shed in Progress' -Detail 'Removing default junk and cache targets.' -Indeterminate
        Write-KobraUiLog -Message 'Quick Shed started...' -BlankLine
        Write-KobraUiLog -Message ("Delete manifest: {0}" -f $manifest.ManifestPath)
        Invoke-KobraShed -Targets $defaultTargets -Log ${function:Write-KobraUiLog}
        Write-KobraUiLog -Message 'Quick Shed complete.'
        Complete-KobraOperationView -Status 'Quick Shed complete' -Detail 'Default cleanup finished.' -NextView 'Results'
    }
    catch {
        Write-KobraUiLog -Message ("Quick Shed failed: {0}" -f $_.Exception.Message)
        Update-KobraOperationView -Status 'Quick Shed failed' -Detail $_.Exception.Message -Value 0
    }
    finally {
        Set-KobraButtonsEnabled -Enabled $true
    }
})

$script:BtnCreateBackup.Add_Click({
    Set-KobraButtonsEnabled -Enabled $false
    Show-KobraOperationView -Status 'Creating backup bundle' -Detail 'Preparing backup files and snapshots.' -OriginView 'Utilities' -Indeterminate

    try {
        Write-KobraUiLog -Message 'Creating manual backup bundle...' -BlankLine
        $backup = Invoke-KobraSettingsBackup -BackupRoot $script:BackupRoot -Log ${function:Write-KobraUiLog}
        if ($null -ne $backup -and $backup.PSObject.Properties.Name -contains 'BackupPath') {
            Write-KobraUiLog -Message ("Backup ready: {0}" -f $backup.BackupPath)
            Update-KobraDashboard
            if (Test-Path -LiteralPath $backup.BackupPath) {
                Start-Process explorer.exe $backup.BackupPath
            }
        }
        Complete-KobraOperationView -Status 'Backup complete' -Detail 'Your backup bundle is ready.' -NextView 'Utilities'
    }
    catch {
        Write-KobraUiLog -Message ("Backup failed: {0}" -f $_.Exception.Message)
        Update-KobraOperationView -Status 'Backup failed' -Detail $_.Exception.Message -Value 0
    }
    finally {
        Set-KobraButtonsEnabled -Enabled $true
    }
})

$script:BtnOpenLogs.Add_Click({
    if (-not (Test-Path -LiteralPath $script:LogRoot)) {
        $null = New-Item -Path $script:LogRoot -ItemType Directory -Force
    }
    Start-Process explorer.exe $script:LogRoot
})

$script:BtnOpenManifests.Add_Click({
    if (-not (Test-Path -LiteralPath $script:ManifestRoot)) {
        $null = New-Item -Path $script:ManifestRoot -ItemType Directory -Force
    }
    Start-Process explorer.exe $script:ManifestRoot
})

$script:BtnExit.Add_Click({
    Write-KobraUiLog -Message 'Exit requested by user.'
    $script:Window.Close()
})

$script:BtnDonate.Add_Click({
    Invoke-KobraDonate
})

if ($null -ne $script:BtnDonateSidebar) {
    $script:BtnDonateSidebar.Add_Click({
        Invoke-KobraDonate
    })
}

$script:BtnDisclaimer.Add_Click({
    Show-KobraDisclaimer
})

$script:BtnAboutMe.Add_Click({
    Show-KobraAboutMe
})

if ($null -ne $script:BtnCancelScan) {
    $script:BtnCancelScan.Add_Click({
        [System.Windows.MessageBox]::Show('Safe cancel is not available yet during active operations. Kobra will finish the current task.', 'KobraOptimizer', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
    })
}

if ($null -ne $script:BtnResultsRescan) {
    $script:BtnResultsRescan.Add_Click({
        Write-KobraDebug -Message ('Rescan requested. LastAnalyzeScope=' + $script:LastAnalyzeScope)
        switch ($script:LastAnalyzeScope) {
            'System' { Invoke-KobraScopeAnalyze -Scope 'System' }
            'Browser' { Invoke-KobraScopeAnalyze -Scope 'Browser' }
            'Registry' { Invoke-KobraScopeAnalyze -Scope 'Registry' }
            default { Invoke-KobraButtonClick -Button $script:BtnAnalyze }
        }
    })
}
if ($null -ne $script:BtnResultsBackCustom) {
    $script:BtnResultsBackCustom.Add_Click({
        Switch-KobraView -ViewName 'CustomClean'
    })
}
if ($null -ne $script:BtnResultsOpenDebugLog) {
    $script:BtnResultsOpenDebugLog.Add_Click({
        if (-not (Test-Path -LiteralPath $script:DebugLogFile)) {
            Write-KobraUiLog -Message 'Debug log does not exist yet.'
            return
        }
        Start-Process notepad.exe $script:DebugLogFile
    })
}

$script:BtnRunSelected.Add_Click({
    Write-KobraDebug -Message 'BtnRunSelected clicked.'
    Write-KobraDebug -Message ('Selections at RunSelected: ' + (Get-KobraDebugSelectionState))
    if ($script:CurrentResultsScope -in @('System','Browser','Registry')) {
        Invoke-KobraScopeClean -Scope $script:CurrentResultsScope
        return
    }

    Set-KobraButtonsEnabled -Enabled $false
    Show-KobraOperationView -Status 'Executing selected tasks' -Detail 'Preparing the selected areas for cleanup.' -OriginView 'Results' -Value 5

    try {
        $cleanupTargets = @(Get-KobraCleanupTargetsFromUi)
        $browserTargets = @(Get-KobraBrowserTargetsFromUi)
        $browserComponents = @(Get-KobraBrowserComponentsFromUi)
        $doTls         = [bool]$script:ChkRegistry.IsChecked
        $doNetwork     = [bool]$script:ChkNetwork.IsChecked
        $doDnsFlush    = [bool]$script:ChkDnsFlush.IsChecked
        $doDnsProfile  = [bool]$script:ChkDnsProfile.IsChecked
        $doHpDebloat   = [bool]$script:ChkHPDebloat.IsChecked
        $doRestore     = [bool]$script:ChkRestorePoint.IsChecked
        $doRegistryTraces = Test-KobraRegistryCleanupSelected
        Write-KobraDebug -Message ('RunSelected doRegistryTraces=' + $doRegistryTraces + '; registryBackupSelected=' + (Test-KobraRegistryBackupSelected))
        $dnsProfile    = Get-KobraSelectedDnsProfile

        if (((Get-KobraSafeCount $cleanupTargets) -gt 0 -or (Get-KobraSafeCount $browserTargets) -gt 0 -or $doRegistryTraces) -and -not (Test-KobraAnalyzeManifestMatchesCurrentSelection -Scope 'All')) {
            Write-KobraUiLog -Message 'Rescan required before Clean Selected. The current selections no longer match the last scan manifest.' -BlankLine
            Update-KobraOperationView -Status 'Execution paused' -Detail 'Rescan the selected items before cleaning.' -Value 0
            return
        }

        $stepCount = 0
        if ((Get-KobraSafeCount $cleanupTargets) -gt 0) { $stepCount++ }
        if ((Get-KobraSafeCount $browserTargets) -gt 0) { $stepCount++ }
        if ($doTls)        { $stepCount++ }
        if ($doNetwork)    { $stepCount++ }
        if ($doDnsFlush)   { $stepCount++ }
        if ($doDnsProfile) { $stepCount++ }
        if ($doHpDebloat)  { $stepCount++ }
        if ($doRegistryTraces) { $stepCount++ }
        if (Test-KobraRegistryBackupSelected) { $stepCount++ }
        if ($doRestore -and ($doTls -or $doNetwork -or $doHpDebloat)) { $stepCount++ }

        if ($stepCount -eq 0) {
            Write-KobraUiLog -Message 'Nothing selected. Check at least one option first.' -BlankLine
            Set-KobraProgress 0
            return
        }

        $manifest = $null
        if ((Get-KobraSafeCount $cleanupTargets) -gt 0 -or (Get-KobraSafeCount $browserTargets) -gt 0 -or $doRegistryTraces) {
            $manifest = $script:LastAnalyzeManifest
        }
        $confirmed = Show-KobraExecutionConfirmation -CleanupTargets $cleanupTargets -BrowserTargets $browserTargets -BrowserComponents $browserComponents -DoTls $doTls -DoNetwork $doNetwork -DoDnsFlush $doDnsFlush -DoDnsProfile $doDnsProfile -DoHpDebloat $doHpDebloat -DoRestore $doRestore -DoRegistryTraces $doRegistryTraces -DnsProfile $dnsProfile -ManifestInfo $manifest
        if (-not $confirmed) {
            Write-KobraUiLog -Message 'Execution canceled by user.' -BlankLine
            Set-KobraProgress 0
            return
        }

        $currentStep = 0
        $executionResults = @()
        $registryBackup = $script:LastRegistryBackup
        Write-KobraUiLog -Message 'Executing selected tasks...' -BlankLine
        if ($null -ne $manifest) {
            Write-KobraUiLog -Message ("Delete manifest: {0}" -f $manifest.ManifestPath)
        }
        Update-KobraOperationView -Status 'Executing selected tasks' -Detail 'Running cleanup and optimization tasks now.' -Value 8

        if ($doRestore -and ($doTls -or $doNetwork -or $doHpDebloat)) {
            Update-KobraOperationView -Status 'Executing selected tasks' -Detail 'Creating a restore point before deeper changes.'
            Write-KobraUiLog -Message 'Creating restore point before system changes...'
            Invoke-KobraGuard -Log ${function:Write-KobraUiLog}
            $currentStep++
            Set-KobraProgress ([math]::Round(($currentStep / $stepCount) * 100))
        }

        if (Test-KobraRegistryBackupSelected) {
            Update-KobraOperationView -Status 'Preparing registry backup' -Detail 'Exporting registry backups before cleanup.'
            Write-KobraUiLog -Message 'Creating registry backup bundle before registry cleanup...'
            $registryBackup = New-KobraRegistryTraceBackup -Log ${function:Write-KobraUiLog}
            $script:LastRegistryBackup = $registryBackup
            Update-KobraRegistryBackupStatusUi
            $currentStep++
            Set-KobraProgress ([math]::Round(($currentStep / $stepCount) * 100))
        }

        if ((Get-KobraSafeCount $cleanupTargets) -gt 0) {
            Update-KobraOperationView -Status 'Cleaning selected items' -Detail 'Removing selected system cleanup targets.'
            $executionResults += @(Invoke-KobraShed -Targets $cleanupTargets -Log ${function:Write-KobraUiLog})
            $currentStep++
            Set-KobraProgress ([math]::Round(($currentStep / $stepCount) * 100))
        }

        if ((Get-KobraSafeCount $browserTargets) -gt 0) {
            Update-KobraOperationView -Status 'Cleaning selected items' -Detail 'Removing selected browser cache data.'
            $executionResults += @(Invoke-KobraBrowserCleanup -Browsers $browserTargets -Components $browserComponents -Log ${function:Write-KobraUiLog})
            $currentStep++
            Set-KobraProgress ([math]::Round(($currentStep / $stepCount) * 100))
        }

        if ($doRegistryTraces) {
            Update-KobraOperationView -Status 'Cleaning selected items' -Detail 'Cleaning the selected registry trace groups.'
            $executionResults += @(Invoke-KobraRegistryTraceCleanup -Log ${function:Write-KobraUiLog} -BackupInfo $registryBackup)
            $currentStep++
            Set-KobraProgress ([math]::Round(($currentStep / $stepCount) * 100))
        }

        if ($doTls) {
            Invoke-KobraTlsHardening -Log ${function:Write-KobraUiLog}
            $currentStep++
            Set-KobraProgress ([math]::Round(($currentStep / $stepCount) * 100))
        }

        if ($doNetwork) {
            Invoke-KobraTcpStrike -Log ${function:Write-KobraUiLog}
            $currentStep++
            Set-KobraProgress ([math]::Round(($currentStep / $stepCount) * 100))
        }

        if ($doDnsFlush) {
            Invoke-KobraDnsFlush -Log ${function:Write-KobraUiLog}
            $currentStep++
            Set-KobraProgress ([math]::Round(($currentStep / $stepCount) * 100))
        }

        if ($doDnsProfile) {
            Invoke-KobraDnsProfile -ProfileName $dnsProfile -Log ${function:Write-KobraUiLog}
            $currentStep++
            Set-KobraProgress ([math]::Round(($currentStep / $stepCount) * 100))
        }

        if ($doHpDebloat) {
            Invoke-KobraHPDebloat -Log ${function:Write-KobraUiLog}
            $currentStep++
            Set-KobraProgress ([math]::Round(($currentStep / $stepCount) * 100))
        }

        Write-KobraUiLog -Message 'Selected tasks complete.'
        Show-KobraCleanupSummary -ExecutionResults $executionResults -ScopeLabel 'Selected tasks'
        Complete-KobraOperationView -Status 'Selected tasks complete' -Detail 'Kobra finished your requested actions. Review the cleanup summary.' -NextView 'Results'
    }
    catch {
        Write-KobraDebug -Message ('Execution exception: ' + $_.Exception.Message)
        try { Write-KobraDebug -Message ('Execution scriptstack: ' + $_.ScriptStackTrace) } catch {}
        Write-KobraUiLog -Message ("Execution failed: {0}" -f $_.Exception.Message)
        Update-KobraOperationView -Status 'Execution failed' -Detail $_.Exception.Message -Value 0
    }
    finally {
        Set-KobraButtonsEnabled -Enabled $true
    }
})

$script:BtnStartupRefresh.Add_Click({
    try {
        Refresh-KobraStartupList
    }
    catch {
        Write-KobraUiLog -Message ("Startup refresh failed: {0}" -f $_.Exception.Message)
    }
})

$script:BtnStartupDisable.Add_Click({
    try {
        $selected = @($script:StartupList.SelectedItems | Where-Object { $_.IsEnabled })
        if ((Get-KobraSafeCount $selected) -eq 0) {
            Write-KobraUiLog -Message 'Select one or more enabled startup entries first.'
            return
        }

        $confirm = Show-KobraYesNo -Title 'KobraOptimizer - Disable Startup Entries' -Message ("Disable {0} selected startup entr{1}?" -f (Get-KobraSafeCount $selected), $(if ((Get-KobraSafeCount $selected) -eq 1) { 'y' } else { 'ies' }))
        if (-not $confirm) {
            Write-KobraUiLog -Message 'Startup disable canceled by user.'
            return
        }

        foreach ($entry in $selected) {
            Disable-KobraStartupEntry -Entry $entry -Log ${function:Write-KobraUiLog}
        }

        Refresh-KobraStartupList
    }
    catch {
        Write-KobraUiLog -Message ("Startup disable failed: {0}" -f $_.Exception.Message)
    }
})

$script:BtnStartupEnable.Add_Click({
    try {
        $selected = @($script:StartupList.SelectedItems | Where-Object { -not $_.IsEnabled })
        if ((Get-KobraSafeCount $selected) -eq 0) {
            Write-KobraUiLog -Message 'Select one or more disabled startup entries first.'
            return
        }

        $confirm = Show-KobraYesNo -Title 'KobraOptimizer - Re-enable Startup Entries' -Message ("Re-enable {0} selected startup entr{1}?" -f (Get-KobraSafeCount $selected), $(if ((Get-KobraSafeCount $selected) -eq 1) { 'y' } else { 'ies' }))
        if (-not $confirm) {
            Write-KobraUiLog -Message 'Startup re-enable canceled by user.'
            return
        }

        foreach ($entry in $selected) {
            Enable-KobraStartupEntry -Entry $entry -Log ${function:Write-KobraUiLog}
        }

        Refresh-KobraStartupList
    }
    catch {
        Write-KobraUiLog -Message ("Startup re-enable failed: {0}" -f $_.Exception.Message)
    }
})

$script:BtnWindowsUpdate.Add_Click({ Invoke-KobraOpenWindowsUpdate })
$script:BtnWindowsUpdateSettings.Add_Click({ Invoke-KobraOpenWindowsUpdateSettings })
$script:BtnWindowsStorage.Add_Click({ Invoke-KobraOpenWindowsStorage })
$script:BtnWindowsApps.Add_Click({ Invoke-KobraOpenWindowsApps })
$script:BtnWindowsStartupSettings.Add_Click({ Invoke-KobraOpenWindowsStartupSettings })
$script:BtnWindowsGameMode.Add_Click({ Invoke-KobraOpenWindowsGameMode })
$script:BtnWindowsGraphics.Add_Click({ Invoke-KobraOpenWindowsGraphics })
$script:BtnWindowsPower.Add_Click({ Invoke-KobraOpenWindowsPower })

$clickEvent = [System.Windows.Controls.Button]::ClickEvent
function Invoke-KobraButtonClick {
    param([Parameter(Mandatory)][System.Windows.Controls.Button]$Button)
    $args = New-Object System.Windows.RoutedEventArgs($clickEvent, $Button)
    $Button.RaiseEvent($args)
}



$script:Window.Add_PreviewKeyDown({
    param($sender, $e)

    $mods = [System.Windows.Input.Keyboard]::Modifiers
    $ctrlPressed = (($mods -band [System.Windows.Input.ModifierKeys]::Control) -eq [System.Windows.Input.ModifierKeys]::Control)
    if (-not $ctrlPressed) {
        return
    }

    switch ([string]$e.Key) {
        'D' { Switch-KobraView -ViewName 'Dashboard'; $e.Handled = $true }
        'Q' { Switch-KobraView -ViewName 'Analyze'; $e.Handled = $true }
        'R' { Switch-KobraView -ViewName 'Results'; $e.Handled = $true }
        'S' { Switch-KobraView -ViewName 'Startup'; $e.Handled = $true }
        'T' { Switch-KobraView -ViewName 'Tools'; $e.Handled = $true }
        'A' { Invoke-KobraButtonClick -Button $script:BtnAnalyze; $e.Handled = $true }
        'E' { Invoke-KobraButtonClick -Button $script:BtnRunSelected; $e.Handled = $true }
        'B' { Invoke-KobraButtonClick -Button $script:BtnCreateBackup; $e.Handled = $true }
        'X' { Invoke-KobraButtonClick -Button $script:BtnExit; $e.Handled = $true }
    }
})

try {
    Refresh-KobraStartupList
}
catch {
    Write-KobraUiLog -Message ("Startup list initialization failed: {0}" -f $_.Exception.Message)
}

Set-KobraQuickScanPreset
Update-KobraSelectionSummary
Update-KobraDashboard
Switch-KobraView -ViewName 'Dashboard'

$null = $script:Window.ShowDialog()
