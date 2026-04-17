#requires -Version 5.1
# ==============================================================================
# KobraOptimizer v1.6.0 - Main Launcher
# Neon UI pass + classic menu bar + toggleable log panel + laptop-friendly fit
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

$script:AppVersion   = '1.7.0'
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

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName PresentationCore

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
    if (-not (Test-KobraPathExists -Path $modulePath)) {
        throw "Required module missing: $modulePath"
    }

    Import-Module $modulePath -Force
}

try {
    [xml]$appearance = Get-Content -Path $script:XamlPath -Raw
    $reader = New-Object System.Xml.XmlNodeReader $appearance
    $script:Window = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    Write-Host "Failed to load XAML: $($_.Exception.Message)" -ForegroundColor Red
    Read-Host 'Press Enter to exit'
    exit
}

$controlNames = @(
    'BtnAnalyze','BtnRunSelected','BtnQuickShed','BtnCreateBackup','BtnOpenLogs','BtnOpenManifests',
    'StatusTextBox','ProgBar','KobraLogo','ResultsList','TxtResultsHeadline','TxtResultsSubHeadline',
    'ChkRestorePoint','ChkRegistry','ChkNetwork','ChkDnsFlush','ChkDnsProfile','ChkHPDebloat',
    'ChkUserTemp','ChkSystemTemp','ChkWinUpdate','ChkThumbCache','ChkShaderCache','ChkRecycleBin',
    'ChkChrome','ChkEdge','ChkFirefox','ChkBrowserCookies','CmbDnsProvider',
    'StartupList','BtnStartupRefresh','BtnStartupDisable','BtnStartupEnable','ChkStartupShowMicrosoft',
    'BtnWindowsUpdate','BtnWindowsUpdateSettings','BtnWindowsStorage','BtnWindowsApps',
    'BtnWindowsStartupSettings','BtnWindowsGameMode','BtnWindowsGraphics','BtnWindowsPower',
    'BtnExit','BtnDonate','BtnDisclaimer','BtnAboutMe','BtnToggleLog',
    'MenuFileAnalyze','MenuFileExecute','MenuFileQuickShed','MenuFileBackup','MenuFileExit',
    'MenuToolsLogs','MenuToolsManifests',
    'MenuHelpSupport','MenuHelpDisclaimer','MenuHelpAbout',
    'BtnNavDashboard','BtnNavAnalyze','BtnNavResults','BtnNavTools',
    'ViewDashboard','ViewAnalyze','ViewResults','ViewTools',
    'BtnDashboardAnalyze','BtnDashboardQuickShed','BtnDashboardBackup','BtnDashboardViewResults',
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
$script:CurrentView = 'Dashboard'

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
    Update-KobraRecentActivity
    Invoke-KobraUiRefresh
}

function Set-KobraProgress {
    param(
        [int]$Value = 0,
        [switch]$Indeterminate
    )

    if ($Indeterminate) {
        $script:ProgBar.IsIndeterminate = $true
    }
    else {
        $script:ProgBar.IsIndeterminate = $false
        $script:ProgBar.Value = [Math]::Max(0, [Math]::Min(100, $Value))
    }

    Invoke-KobraUiRefresh
}

function Set-KobraButtonsEnabled {
    param([bool]$Enabled)

    $buttonNames = @(
        'BtnAnalyze','BtnRunSelected','BtnQuickShed','BtnCreateBackup','BtnOpenLogs','BtnOpenManifests',
        'BtnStartupRefresh','BtnStartupDisable','BtnStartupEnable',
        'BtnExit','BtnDonate','BtnDisclaimer','BtnAboutMe','BtnToggleLog',
        'BtnNavDashboard','BtnNavAnalyze','BtnNavResults','BtnNavTools',
        'BtnDashboardAnalyze','BtnDashboardQuickShed','BtnDashboardBackup','BtnDashboardViewResults'
    )

    foreach ($name in $buttonNames) {
        $control = Get-Variable -Scope Script -Name $name -ValueOnly -ErrorAction SilentlyContinue
        if ($null -ne $control) { $control.IsEnabled = $Enabled }
    }

    if ($null -ne $script:ChkStartupShowMicrosoft) {
        $script:ChkStartupShowMicrosoft.IsEnabled = $Enabled
    }

    Invoke-KobraUiRefresh
}

function Set-KobraPreferredFont {
    $preferred = 'JetBrains Mono'
    $fallbacks = @('Consolas','Segoe UI')
    $fontName = $null

    try {
        $installedFonts = @([System.Windows.Media.Fonts]::SystemFontFamilies | ForEach-Object { $_.Source })
        if ($installedFonts -contains $preferred) {
            $fontName = $preferred
        }
        else {
            foreach ($candidate in $fallbacks) {
                if ($installedFonts -contains $candidate) {
                    $fontName = $candidate
                    break
                }
            }
        }
    }
    catch {
    }

    if ([string]::IsNullOrWhiteSpace($fontName)) {
        $fontName = 'Segoe UI'
    }

    $font = New-Object System.Windows.Media.FontFamily($fontName)
    $script:Window.FontFamily = $font
    $script:StatusTextBox.FontFamily = $font
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
        [string[]]$BrowserComponents = @('Cache')
    )

    $cleanupCandidates = @()
    $browserCandidates = @()

    if ((Get-KobraSafeCount $CleanupTargets) -gt 0) {
        $cleanupCandidates = @(Get-KobraCleanupCandidates -Targets $CleanupTargets)
    }
    if ((Get-KobraSafeCount $BrowserTargets) -gt 0) {
        $browserCandidates = @(Get-KobraBrowserCandidates -Browsers $BrowserTargets -Components $BrowserComponents)
    }

    $allCandidates = @($cleanupCandidates) + @($browserCandidates)
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
            Category  = $group.Name
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
    if ($DoRestore -and ($DoTls -or $DoNetwork -or $DoHpDebloat)) { $actions.Add('Create restore point first') }

    $messageLines = New-Object System.Collections.Generic.List[string]
    $messageLines.Add('You are about to execute the following Kobra actions:')
    $messageLines.Add('')
    foreach ($action in $actions) { $messageLines.Add(('• {0}' -f $action)) }
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
    return $browsers
}

function Get-KobraBrowserComponentsFromUi {
    $components = @('Cache')
    if ($script:ChkBrowserCookies.IsChecked) {
        $components += 'Cookies'
    }
    return $components
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
        default { return '' }
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

    if ($TotalRecords -gt 0) {
        $script:TxtResultsHeadline.Text = ('{0:N0} records can be removed' -f $TotalRecords)
        $script:TxtResultsSubHeadline.Text = ('Estimated reclaimable space: {0:N2} MB across {1} categories.' -f ($TotalBytes / 1MB), $bindingList.Count)
    }
    else {
        $script:TxtResultsHeadline.Text = 'No removable records detected'
        $script:TxtResultsSubHeadline.Text = 'Analyze scanned the current selections but did not find removable records.'
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
        '*Chrome Cookies' { return ([bool]$script:ChkChrome.IsChecked -and [bool]$script:ChkBrowserCookies.IsChecked) }
        '*Edge Cookies' { return ([bool]$script:ChkEdge.IsChecked -and [bool]$script:ChkBrowserCookies.IsChecked) }
        '*Firefox Cookies' { return ([bool]$script:ChkFirefox.IsChecked -and [bool]$script:ChkBrowserCookies.IsChecked) }
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
    }

    $script:TxtSelectedCategoryCount.Text = ('{0} categories selected' -f $selectedCategories)
    $script:TxtSelectedRecordCount.Text = ('{0:N0} records selected' -f $selectedRecords)
    $script:TxtSelectedBytes.Text = ('{0:N2} MB selected' -f ($selectedBytes / 1MB))
    $script:TxtSelectedWarnings.Text = if ($warnings.Count -gt 0) { ($warnings -join '  •  ') } else { 'Review the results, then run Clean Selected when you are ready.' }
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

    $normalBackground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString('#1A1A1A'))
    $normalBorder = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString('#1F1F1F'))
    $activeBackground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString('#00FF41'))
    $activeBorder = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString('#00FF41'))
    $activeForeground = [System.Windows.Media.Brushes]::Black
    $normalForeground = [System.Windows.Media.Brushes]::White

    $navMap = @{
        Dashboard = $script:BtnNavDashboard
        Analyze   = $script:BtnNavAnalyze
        Results   = $script:BtnNavResults
        Tools     = $script:BtnNavTools
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
        Dashboard = $script:ViewDashboard
        Analyze   = $script:ViewAnalyze
        Results   = $script:ViewResults
        Tools     = $script:ViewTools
    }

    foreach ($view in $views.Values) {
        if ($null -ne $view) { $view.Visibility = 'Collapsed' }
    }

    if ($views.ContainsKey($ViewName) -and $null -ne $views[$ViewName]) {
        $views[$ViewName].Visibility = 'Visible'
        $script:CurrentView = $ViewName
    }

    Set-KobraNavState -ActiveView $ViewName
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
        $script:BtnToggleLog.Content = 'Hide log'
        Write-KobraUiLog -Message 'Log panel expanded.'
    }
    else {
        $script:LogRow.Height = New-Object System.Windows.GridLength(0)
        $script:BtnToggleLog.Content = 'Show log'
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

$script:CmbDnsProvider.SelectedIndex = 0
Update-KobraDnsControls
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
Set-KobraAnalyzeStatus -Title 'Ready to analyze' -SubTitle 'Select your cleanup categories, then run Analyze to build a reviewable results screen.'

$script:ChkDnsProfile.Add_Click({ Update-KobraDnsControls })
$script:ChkStartupShowMicrosoft.Add_Click({ Refresh-KobraStartupList })
$script:BtnToggleLog.Add_Click({ Toggle-KobraLogPanel })
$script:BtnNavDashboard.Add_Click({ Switch-KobraView -ViewName 'Dashboard' })
$script:BtnNavAnalyze.Add_Click({ Switch-KobraView -ViewName 'Analyze' })
$script:BtnNavResults.Add_Click({ Switch-KobraView -ViewName 'Results' })
$script:BtnNavTools.Add_Click({ Switch-KobraView -ViewName 'Tools' })
$script:BtnDashboardAnalyze.Add_Click({ Switch-KobraView -ViewName 'Analyze'; Invoke-KobraButtonClick -Button $script:BtnAnalyze })
$script:BtnDashboardQuickShed.Add_Click({ Invoke-KobraButtonClick -Button $script:BtnQuickShed })
$script:BtnDashboardBackup.Add_Click({ Invoke-KobraButtonClick -Button $script:BtnCreateBackup })
$script:BtnDashboardViewResults.Add_Click({ Switch-KobraView -ViewName 'Results' })

$selectionControls = @(
    $script:ChkUserTemp,$script:ChkSystemTemp,$script:ChkWinUpdate,$script:ChkThumbCache,$script:ChkShaderCache,$script:ChkRecycleBin,
    $script:ChkChrome,$script:ChkEdge,$script:ChkFirefox,$script:ChkBrowserCookies
)
foreach ($control in $selectionControls) {
    if ($null -ne $control) {
        $control.Add_Click({ Update-KobraSelectionSummary; Update-KobraDashboard })
    }
}

$script:BtnAnalyze.Add_Click({
    Set-KobraButtonsEnabled -Enabled $false
    Set-KobraProgress 10

    try {
        Switch-KobraView -ViewName 'Analyze'
        Set-KobraAnalyzeStatus -Title 'Analyzing your system' -SubTitle 'Reviewing selected cleanup categories and browser data.'
        Write-KobraUiLog -Message 'Starting analysis scan...' -BlankLine

        $cleanupTargets = @(Get-KobraCleanupTargetsFromUi)
        if ((Get-KobraSafeCount $cleanupTargets) -eq 0) {
            $cleanupTargets = @('UserTemp','SystemTemp','WindowsUpdate','ThumbnailCache','RecycleBin')
        }

        $cleanupPreview = @(Get-KobraCleanupPreview -Targets $cleanupTargets)
        $browserTargets = @(Get-KobraBrowserTargetsFromUi)
        $browserComponents = @(Get-KobraBrowserComponentsFromUi)
        $browserPreview = @()
        if ((Get-KobraSafeCount $browserTargets) -gt 0) {
            $browserPreview = @(Get-KobraBrowserPreview -Browsers $browserTargets -Components $browserComponents)
        }

        $totalBytes = [int64]0

        Set-KobraAnalyzeStatus -Title 'Scanning cleanup targets' -SubTitle 'Inspecting temp files, caches, and recycle locations.'
        Write-KobraUiLog -Message 'Cleanup targets:' -NoTimestamp
        foreach ($item in $cleanupPreview) {
            if ($null -eq $item) { continue }
            $totalBytes += Get-KobraSafeInt64 $item.SizeBytes
            Write-KobraUiLog -Message ("  {0}: {1:N2} MB ({2} items)" -f $item.Name, [double]$item.SizeMB, (Get-KobraSafeInt64 $item.Items)) -NoTimestamp
        }

        if ((Get-KobraSafeCount $browserPreview) -gt 0) {
            Set-KobraAnalyzeStatus -Title 'Scanning browser data' -SubTitle 'Counting browser cache records and optional cookie targets.'
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

        $manifest = New-KobraDeleteManifest -CleanupTargets $cleanupTargets -BrowserTargets $browserTargets -BrowserComponents $browserComponents
        $script:LastAnalyzeManifest = $manifest
        $script:LastAnalyzeTime = Get-Date
        Update-KobraResultsPanel -CategorySummary $manifest.CategorySummary -TotalRecords $manifest.CandidateCount -TotalBytes $manifest.TotalBytes
        Write-KobraUiLog -Message ("Delete manifest written: {0}" -f $manifest.ManifestPath) -NoTimestamp
        Write-KobraUiLog -Message ("Estimated reclaimable space: {0:N2} MB" -f ($totalBytes / 1MB))
        Write-KobraUiLog -Message ("Estimated removable records: {0:N0}" -f $manifest.CandidateCount)
        Set-KobraAnalyzeStatus -Title 'Analysis complete' -SubTitle 'Your scan results are ready to review.'
        Set-KobraProgress 100
        Switch-KobraView -ViewName 'Results'
    }
    catch {
        Set-KobraAnalyzeStatus -Title 'Analysis failed' -SubTitle $_.Exception.Message
        Write-KobraUiLog -Message ("Analysis failed: {0}" -f $_.Exception.Message)
        Set-KobraProgress 0
    }
    finally {
        Set-KobraButtonsEnabled -Enabled $true
    }
})

$script:BtnQuickShed.Add_Click({
    Set-KobraButtonsEnabled -Enabled $false
    Set-KobraProgress -Indeterminate

    try {
        $defaultTargets = @('UserTemp','SystemTemp','ThumbnailCache','ShaderCache','RecycleBin')
        $manifest = New-KobraDeleteManifest -CleanupTargets $defaultTargets -BrowserTargets @() -BrowserComponents @('Cache')

        $confirmed = Show-KobraExecutionConfirmation -CleanupTargets $defaultTargets -BrowserTargets @() -BrowserComponents @('Cache') -DoTls $false -DoNetwork $false -DoDnsFlush $false -DoDnsProfile $false -DoHpDebloat $false -DoRestore $false -DnsProfile '' -ManifestInfo $manifest
        if (-not $confirmed) {
            Write-KobraUiLog -Message 'Quick Shed canceled by user.' -BlankLine
            Set-KobraProgress 0
            return
        }

        Write-KobraUiLog -Message 'Quick Shed started...' -BlankLine
        Write-KobraUiLog -Message ("Delete manifest: {0}" -f $manifest.ManifestPath)
        Invoke-KobraShed -Targets $defaultTargets -Log ${function:Write-KobraUiLog}
        Write-KobraUiLog -Message 'Quick Shed complete.'
        Set-KobraProgress 100
    }
    catch {
        Write-KobraUiLog -Message ("Quick Shed failed: {0}" -f $_.Exception.Message)
        Set-KobraProgress 0
    }
    finally {
        Set-KobraButtonsEnabled -Enabled $true
    }
})

$script:BtnCreateBackup.Add_Click({
    Set-KobraButtonsEnabled -Enabled $false
    Set-KobraProgress -Indeterminate

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
        Set-KobraProgress 100
    }
    catch {
        Write-KobraUiLog -Message ("Backup failed: {0}" -f $_.Exception.Message)
        Set-KobraProgress 0
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

$script:BtnDisclaimer.Add_Click({
    Show-KobraDisclaimer
})

$script:BtnAboutMe.Add_Click({
    Show-KobraAboutMe
})

$script:BtnRunSelected.Add_Click({
    Set-KobraButtonsEnabled -Enabled $false

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
        $dnsProfile    = Get-KobraSelectedDnsProfile

        $stepCount = 0
        if ((Get-KobraSafeCount $cleanupTargets) -gt 0) { $stepCount++ }
        if ((Get-KobraSafeCount $browserTargets) -gt 0) { $stepCount++ }
        if ($doTls)        { $stepCount++ }
        if ($doNetwork)    { $stepCount++ }
        if ($doDnsFlush)   { $stepCount++ }
        if ($doDnsProfile) { $stepCount++ }
        if ($doHpDebloat)  { $stepCount++ }
        if ($doRestore -and ($doTls -or $doNetwork -or $doHpDebloat)) { $stepCount++ }

        if ($stepCount -eq 0) {
            Write-KobraUiLog -Message 'Nothing selected. Check at least one option first.' -BlankLine
            Set-KobraProgress 0
            return
        }

        $manifest = New-KobraDeleteManifest -CleanupTargets $cleanupTargets -BrowserTargets $browserTargets -BrowserComponents $browserComponents
        $confirmed = Show-KobraExecutionConfirmation -CleanupTargets $cleanupTargets -BrowserTargets $browserTargets -BrowserComponents $browserComponents -DoTls $doTls -DoNetwork $doNetwork -DoDnsFlush $doDnsFlush -DoDnsProfile $doDnsProfile -DoHpDebloat $doHpDebloat -DoRestore $doRestore -DnsProfile $dnsProfile -ManifestInfo $manifest
        if (-not $confirmed) {
            Write-KobraUiLog -Message 'Execution canceled by user.' -BlankLine
            Set-KobraProgress 0
            return
        }

        $currentStep = 0
        Write-KobraUiLog -Message 'Executing selected tasks...' -BlankLine
        Write-KobraUiLog -Message ("Delete manifest: {0}" -f $manifest.ManifestPath)
        Set-KobraProgress 5

        if ($doRestore -and ($doTls -or $doNetwork -or $doHpDebloat)) {
            Write-KobraUiLog -Message 'Creating restore point before system changes...'
            Invoke-KobraGuard -Log ${function:Write-KobraUiLog}
            $currentStep++
            Set-KobraProgress ([math]::Round(($currentStep / $stepCount) * 100))
        }

        if ((Get-KobraSafeCount $cleanupTargets) -gt 0) {
            Invoke-KobraShed -Targets $cleanupTargets -Log ${function:Write-KobraUiLog}
            $currentStep++
            Set-KobraProgress ([math]::Round(($currentStep / $stepCount) * 100))
        }

        if ((Get-KobraSafeCount $browserTargets) -gt 0) {
            Invoke-KobraBrowserCleanup -Browsers $browserTargets -Components $browserComponents -Log ${function:Write-KobraUiLog}
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
        Set-KobraProgress 100
    }
    catch {
        Write-KobraUiLog -Message ("Execution failed: {0}" -f $_.Exception.Message)
        Set-KobraProgress 0
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

$script:MenuFileAnalyze.Add_Click({ Invoke-KobraButtonClick -Button $script:BtnAnalyze })
$script:MenuFileExecute.Add_Click({ Invoke-KobraButtonClick -Button $script:BtnRunSelected })
$script:MenuFileQuickShed.Add_Click({ Invoke-KobraButtonClick -Button $script:BtnQuickShed })
$script:MenuFileBackup.Add_Click({ Invoke-KobraButtonClick -Button $script:BtnCreateBackup })
$script:MenuFileExit.Add_Click({ Invoke-KobraButtonClick -Button $script:BtnExit })
$script:MenuToolsLogs.Add_Click({ Invoke-KobraButtonClick -Button $script:BtnOpenLogs })
$script:MenuToolsManifests.Add_Click({ Invoke-KobraButtonClick -Button $script:BtnOpenManifests })
$script:MenuHelpSupport.Add_Click({ Invoke-KobraDonate })
$script:MenuHelpDisclaimer.Add_Click({ Show-KobraDisclaimer })
$script:MenuHelpAbout.Add_Click({ Show-KobraAboutMe })

$script:Window.Add_PreviewKeyDown({
    param($sender, $e)

    $mods = [System.Windows.Input.Keyboard]::Modifiers
    $ctrlPressed = (($mods -band [System.Windows.Input.ModifierKeys]::Control) -eq [System.Windows.Input.ModifierKeys]::Control)
    if (-not $ctrlPressed) {
        return
    }

    switch ([string]$e.Key) {
        'A' { Invoke-KobraButtonClick -Button $script:BtnAnalyze; $e.Handled = $true }
        'E' { Invoke-KobraButtonClick -Button $script:BtnRunSelected; $e.Handled = $true }
        'Q' { Invoke-KobraButtonClick -Button $script:BtnQuickShed; $e.Handled = $true }
        'B' { Invoke-KobraButtonClick -Button $script:BtnCreateBackup; $e.Handled = $true }
        'L' { Invoke-KobraButtonClick -Button $script:BtnOpenLogs; $e.Handled = $true }
        'M' { Invoke-KobraButtonClick -Button $script:BtnOpenManifests; $e.Handled = $true }
        'X' { Invoke-KobraButtonClick -Button $script:BtnExit; $e.Handled = $true }
    }
})

try {
    Refresh-KobraStartupList
}
catch {
    Write-KobraUiLog -Message ("Startup list initialization failed: {0}" -f $_.Exception.Message)
}

Update-KobraSelectionSummary
Update-KobraDashboard
Switch-KobraView -ViewName 'Dashboard'

$null = $script:Window.ShowDialog()
