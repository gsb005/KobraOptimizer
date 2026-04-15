
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-KobraStartupLog {
    param([scriptblock]$Log,[string]$Message)
    if ($Log) { & $Log $Message } else { Write-Host $Message }
}

function Get-KobraSafeCount {
    param($Value)
    if ($null -eq $Value) { return 0 }
    return @($Value).Count
}

function Get-KobraStartupConfig {
    [pscustomobject]@{
        UserRunPath        = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
        MachineRunPath     = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run'
        UserBackupRunPath  = 'HKCU:\Software\KobraOptimizer\StartupBackup\Run'
        MachineBackupRunPath = 'HKLM:\Software\KobraOptimizer\StartupBackup\Run'
        UserStartupFolder  = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup'
        SystemStartupFolder= Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs\Startup'
        BackupRoot         = Join-Path $env:ProgramData 'KobraOptimizer\StartupBackup'
    }
}

function Get-KobraStartupBackupFolder {
    param([string]$Location)

    $config = Get-KobraStartupConfig
    $folder = switch ($Location) {
        'User Startup'   { Join-Path $config.BackupRoot 'UserStartup' }
        'System Startup' { Join-Path $config.BackupRoot 'SystemStartup' }
        default          { Join-Path $config.BackupRoot 'Other' }
    }

    $null = New-Item -Path $folder -ItemType Directory -Force -ErrorAction SilentlyContinue
    return $folder
}

function Test-KobraStartupMicrosoftEntry {
    param([string]$Command,[string]$Name)

    $text = ('{0} {1}' -f $Name, $Command)
    return ($text -match '(?i)\\Microsoft\\|\\Windows\\|Microsoft Corporation|Windows Security|OneDrive')
}

function Get-KobraStartupRegistryEntries {
    param(
        [string]$LivePath,
        [string]$BackupPath,
        [string]$Location,
        [bool]$Enabled
    )

    $sourcePath = if ($Enabled) { $LivePath } else { $BackupPath }
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        return @()
    }

    $item = Get-ItemProperty -LiteralPath $sourcePath -ErrorAction SilentlyContinue
    if ($null -eq $item) {
        return @()
    }

    $results = @()
    foreach ($prop in $item.PSObject.Properties) {
        if ($prop.Name -in 'PSPath','PSParentPath','PSChildName','PSDrive','PSProvider') { continue }
        if ([string]::IsNullOrWhiteSpace([string]$prop.Value)) { continue }

        $results += [pscustomobject]@{
            Name         = $prop.Name
            Status       = if ($Enabled) { 'Enabled' } else { 'Disabled' }
            Location     = $Location
            Command      = [string]$prop.Value
            Kind         = 'Registry'
            IsEnabled    = $Enabled
            IsMicrosoft  = (Test-KobraStartupMicrosoftEntry -Command ([string]$prop.Value) -Name $prop.Name)
            LivePath     = $LivePath
            BackupPath   = $BackupPath
            LiveFilePath = $null
            BackupFilePath = $null
        }
    }

    return $results
}

function Get-KobraStartupFolderEntries {
    param(
        [string]$LiveFolder,
        [string]$BackupFolder,
        [string]$Location,
        [bool]$Enabled
    )

    $folder = if ($Enabled) { $LiveFolder } else { $BackupFolder }
    if (-not (Test-Path -LiteralPath $folder)) {
        return @()
    }

    $results = @()
    foreach ($file in (Get-ChildItem -LiteralPath $folder -File -Force -ErrorAction SilentlyContinue)) {
        $liveFile = Join-Path $LiveFolder $file.Name
        $backupFile = Join-Path $BackupFolder $file.Name

        $results += [pscustomobject]@{
            Name         = $file.BaseName
            Status       = if ($Enabled) { 'Enabled' } else { 'Disabled' }
            Location     = $Location
            Command      = if ($Enabled) { $file.FullName } else { $liveFile }
            Kind         = 'StartupFolder'
            IsEnabled    = $Enabled
            IsMicrosoft  = (Test-KobraStartupMicrosoftEntry -Command $file.FullName -Name $file.BaseName)
            LivePath     = $null
            BackupPath   = $null
            LiveFilePath = $liveFile
            BackupFilePath = $backupFile
        }
    }

    return $results
}

function Get-KobraStartupEntries {
    [CmdletBinding()]
    param([switch]$IncludeMicrosoft)

    $config = Get-KobraStartupConfig
    $null = New-Item -Path $config.BackupRoot -ItemType Directory -Force -ErrorAction SilentlyContinue

    $entries = @()
    $entries += Get-KobraStartupRegistryEntries -LivePath $config.UserRunPath -BackupPath $config.UserBackupRunPath -Location 'HKCU Run' -Enabled $true
    $entries += Get-KobraStartupRegistryEntries -LivePath $config.MachineRunPath -BackupPath $config.MachineBackupRunPath -Location 'HKLM Run' -Enabled $true
    $entries += Get-KobraStartupRegistryEntries -LivePath $config.UserRunPath -BackupPath $config.UserBackupRunPath -Location 'HKCU Run' -Enabled $false
    $entries += Get-KobraStartupRegistryEntries -LivePath $config.MachineRunPath -BackupPath $config.MachineBackupRunPath -Location 'HKLM Run' -Enabled $false

    $userBackup = Get-KobraStartupBackupFolder -Location 'User Startup'
    $systemBackup = Get-KobraStartupBackupFolder -Location 'System Startup'
    $entries += Get-KobraStartupFolderEntries -LiveFolder $config.UserStartupFolder -BackupFolder $userBackup -Location 'User Startup' -Enabled $true
    $entries += Get-KobraStartupFolderEntries -LiveFolder $config.SystemStartupFolder -BackupFolder $systemBackup -Location 'System Startup' -Enabled $true
    $entries += Get-KobraStartupFolderEntries -LiveFolder $config.UserStartupFolder -BackupFolder $userBackup -Location 'User Startup' -Enabled $false
    $entries += Get-KobraStartupFolderEntries -LiveFolder $config.SystemStartupFolder -BackupFolder $systemBackup -Location 'System Startup' -Enabled $false

    if (-not $IncludeMicrosoft) {
        $entries = @($entries | Where-Object { -not $_.IsMicrosoft })
    }

    return @($entries | Sort-Object @{Expression='Status';Descending=$true}, Location, Name)
}

function Disable-KobraStartupEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][psobject]$Entry,
        [scriptblock]$Log
    )

    if (-not $Entry.IsEnabled) {
        Write-KobraStartupLog -Log $Log -Message ("Startup entry already disabled: {0}" -f $Entry.Name)
        return
    }

    switch ($Entry.Kind) {
        'Registry' {
            $null = New-Item -Path $Entry.BackupPath -Force -ErrorAction SilentlyContinue
            try {
                $current = (Get-ItemProperty -LiteralPath $Entry.LivePath -Name $Entry.Name -ErrorAction Stop).$($Entry.Name)
                New-ItemProperty -Path $Entry.BackupPath -Name $Entry.Name -Value $current -PropertyType String -Force | Out-Null
                Remove-ItemProperty -Path $Entry.LivePath -Name $Entry.Name -ErrorAction Stop
                Write-KobraStartupLog -Log $Log -Message ("Disabled startup item: {0} ({1})" -f $Entry.Name, $Entry.Location)
            }
            catch {
                throw "Could not disable startup item '$($Entry.Name)': $($_.Exception.Message)"
            }
        }
        'StartupFolder' {
            $backupFolder = Split-Path -Parent $Entry.BackupFilePath
            $null = New-Item -Path $backupFolder -ItemType Directory -Force -ErrorAction SilentlyContinue
            try {
                Move-Item -LiteralPath $Entry.LiveFilePath -Destination $Entry.BackupFilePath -Force -ErrorAction Stop
                Write-KobraStartupLog -Log $Log -Message ("Disabled startup item: {0} ({1})" -f $Entry.Name, $Entry.Location)
            }
            catch {
                throw "Could not disable startup file '$($Entry.Name)': $($_.Exception.Message)"
            }
        }
    }
}

function Enable-KobraStartupEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][psobject]$Entry,
        [scriptblock]$Log
    )

    if ($Entry.IsEnabled) {
        Write-KobraStartupLog -Log $Log -Message ("Startup entry already enabled: {0}" -f $Entry.Name)
        return
    }

    switch ($Entry.Kind) {
        'Registry' {
            $null = New-Item -Path $Entry.LivePath -Force -ErrorAction SilentlyContinue
            try {
                $value = (Get-ItemProperty -LiteralPath $Entry.BackupPath -Name $Entry.Name -ErrorAction Stop).$($Entry.Name)
                New-ItemProperty -Path $Entry.LivePath -Name $Entry.Name -Value $value -PropertyType String -Force | Out-Null
                Remove-ItemProperty -Path $Entry.BackupPath -Name $Entry.Name -ErrorAction Stop
                Write-KobraStartupLog -Log $Log -Message ("Re-enabled startup item: {0} ({1})" -f $Entry.Name, $Entry.Location)
            }
            catch {
                throw "Could not re-enable startup item '$($Entry.Name)': $($_.Exception.Message)"
            }
        }
        'StartupFolder' {
            $liveFolder = Split-Path -Parent $Entry.LiveFilePath
            $null = New-Item -Path $liveFolder -ItemType Directory -Force -ErrorAction SilentlyContinue
            try {
                Move-Item -LiteralPath $Entry.BackupFilePath -Destination $Entry.LiveFilePath -Force -ErrorAction Stop
                Write-KobraStartupLog -Log $Log -Message ("Re-enabled startup item: {0} ({1})" -f $Entry.Name, $Entry.Location)
            }
            catch {
                throw "Could not re-enable startup file '$($Entry.Name)': $($_.Exception.Message)"
            }
        }
    }
}

Export-ModuleMember -Function Get-KobraStartupEntries, Disable-KobraStartupEntry, Enable-KobraStartupEntry
