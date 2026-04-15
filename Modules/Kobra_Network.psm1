Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-KobraModuleLog {
    param(
        [scriptblock]$Log,
        [string]$Message
    )

    if ($Log) {
        & $Log $Message
    }
    else {
        Write-Host $Message
    }
}

function Get-KobraSafeCount {
    param($Value)

    if ($null -eq $Value) {
        return 0
    }

    return @($Value).Count
}

function Invoke-KobraNative {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string[]]$ArgumentList,
        [scriptblock]$Log
    )

    $output = & $FilePath @ArgumentList 2>&1
    if ($LASTEXITCODE -ne 0) {
        $joinedArgs = $ArgumentList -join ' '
        $detail = ($output | Out-String).Trim()
        throw "Command failed: $FilePath $joinedArgs`n$detail"
    }

    if ($output) {
        foreach ($line in ($output | Out-String).Trim().Split([Environment]::NewLine)) {
            if ($line.Trim()) {
                Write-KobraModuleLog -Log $Log -Message ("  {0}" -f $line.Trim())
            }
        }
    }
}

function Set-KobraRegDword {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][uint32]$Value
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        $null = New-Item -Path $Path -Force
    }

    New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType DWord -Force | Out-Null
}

function Convert-KobraRegistryPathToNative {
    param([Parameter(Mandatory)][string]$RegistryPath)

    switch -Regex ($RegistryPath) {
        '^HKLM:\\' { return ($RegistryPath -replace '^HKLM:\\', 'HKEY_LOCAL_MACHINE\') }
        '^HKCU:\\' { return ($RegistryPath -replace '^HKCU:\\', 'HKEY_CURRENT_USER\') }
        '^HKCR:\\' { return ($RegistryPath -replace '^HKCR:\\', 'HKEY_CLASSES_ROOT\') }
        '^HKU:\\'  { return ($RegistryPath -replace '^HKU:\\',  'HKEY_USERS\') }
        default     { return $RegistryPath }
    }
}

function Export-KobraRegistryBranch {
    param(
        [Parameter(Mandatory)][string]$RegistryPath,
        [Parameter(Mandatory)][string]$DestinationPath,
        [scriptblock]$Log
    )

    if (-not (Test-Path -LiteralPath $RegistryPath)) {
        Write-KobraModuleLog -Log $Log -Message ("  Registry path not found, skipped: {0}" -f $RegistryPath)
        return
    }

    $nativePath = Convert-KobraRegistryPathToNative -RegistryPath $RegistryPath
    $output = & reg.exe export $nativePath $DestinationPath /y 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-KobraModuleLog -Log $Log -Message ("  Registry export saved: {0}" -f (Split-Path -Leaf $DestinationPath))
    }
    else {
        $detail = ($output | Out-String).Trim()
        Write-KobraModuleLog -Log $Log -Message ("  Registry export failed for {0}: {1}" -f $RegistryPath, $detail)
    }
}

function Get-KobraDnsBackupSnapshot {
    $results = @()
    $entries = @(Get-DnsClientServerAddress -ErrorAction SilentlyContinue | Where-Object { $_.InterfaceAlias })

    foreach ($entry in $entries) {
        $servers = @()
        if ($null -ne $entry.ServerAddresses) {
            $servers = @($entry.ServerAddresses)
        }

        $results += [pscustomobject]@{
            InterfaceAlias  = $entry.InterfaceAlias
            InterfaceIndex  = $entry.InterfaceIndex
            AddressFamily   = [string]$entry.AddressFamily
            ServerAddresses = $servers
            Automatic       = (Get-KobraSafeCount $servers) -eq 0
        }
    }

    return $results
}

function Invoke-KobraSettingsBackup {
    [CmdletBinding()]
    param(
        [string]$BackupRoot,
        [scriptblock]$Log
    )

    if ([string]::IsNullOrWhiteSpace($BackupRoot)) {
        $BackupRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'Backups'
    }

    $null = New-Item -Path $BackupRoot -ItemType Directory -Force -ErrorAction SilentlyContinue
    $backupDir = Join-Path $BackupRoot ("KobraBackup_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $null = New-Item -Path $backupDir -ItemType Directory -Force -ErrorAction Stop

    Write-KobraModuleLog -Log $Log -Message ("Creating backup bundle: {0}" -f $backupDir)

    $dnsSnapshotPath = Join-Path $backupDir 'dns_snapshot.json'
    $dnsRecords = @(Get-KobraDnsBackupSnapshot)
    $dnsRecords | ConvertTo-Json -Depth 5 | Set-Content -Path $dnsSnapshotPath -Encoding UTF8
    Write-KobraModuleLog -Log $Log -Message '  DNS snapshot exported.'

    $restoreDnsPath = Join-Path $backupDir 'restore_dns.ps1'
    $restoreDnsScript = @'
#requires -Version 5.1
param(
    [string]$SnapshotPath = (Join-Path $PSScriptRoot 'dns_snapshot.json')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $SnapshotPath)) {
    throw "Snapshot not found: $SnapshotPath"
}

$records = Get-Content -LiteralPath $SnapshotPath -Raw | ConvertFrom-Json
if ($null -eq $records) {
    throw 'Snapshot file is empty.'
}

$records = @($records)
foreach ($record in $records) {
    $servers = @()
    if ($null -ne $record.ServerAddresses) {
        $servers = @($record.ServerAddresses)
    }

    if (@($servers).Count -eq 0) {
        Set-DnsClientServerAddress -InterfaceAlias $record.InterfaceAlias -ResetServerAddresses -ErrorAction Continue
    }
    else {
        Set-DnsClientServerAddress -InterfaceAlias $record.InterfaceAlias -ServerAddresses $servers -AddressFamily $record.AddressFamily -ErrorAction Continue
    }
}

Write-Host 'DNS restore attempt complete.'
'@
    Set-Content -Path $restoreDnsPath -Value $restoreDnsScript -Encoding UTF8
    Write-KobraModuleLog -Log $Log -Message '  DNS restore helper created.'

    try {
        (& netsh.exe int tcp show global 2>&1 | Out-String).Trim() | Set-Content -Path (Join-Path $backupDir 'netsh_tcp_global.txt') -Encoding UTF8
        (& netsh.exe int tcp show supplemental 2>&1 | Out-String).Trim() | Set-Content -Path (Join-Path $backupDir 'netsh_tcp_supplemental.txt') -Encoding UTF8
        Write-KobraModuleLog -Log $Log -Message '  TCP snapshots exported.'
    }
    catch {
        Write-KobraModuleLog -Log $Log -Message ("  TCP snapshot export reported: {0}" -f $_.Exception.Message)
    }

    try {
        (& ipconfig.exe /all 2>&1 | Out-String).Trim() | Set-Content -Path (Join-Path $backupDir 'ipconfig_all.txt') -Encoding UTF8
        Write-KobraModuleLog -Log $Log -Message '  IP configuration snapshot exported.'
    }
    catch {
        Write-KobraModuleLog -Log $Log -Message ("  IP configuration snapshot reported: {0}" -f $_.Exception.Message)
    }

    $registryExports = @(
        @{ Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols'; File = 'schannel_protocols.reg' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319'; File = 'dotnet_v4_x64.reg' },
        @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319'; File = 'dotnet_v4_wow64.reg' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile'; File = 'multimedia_systemprofile.reg' }
    )

    foreach ($export in $registryExports) {
        Export-KobraRegistryBranch -RegistryPath $export.Path -DestinationPath (Join-Path $backupDir $export.File) -Log $Log
    }

    $manifestPath = Join-Path $backupDir 'manifest.json'
    [pscustomobject]@{
        Product              = 'KobraOptimizer'
        Version              = '1.2.2'
        CreatedAt            = (Get-Date).ToString('s')
        BackupPath           = $backupDir
        DnsRecordCount       = (Get-KobraSafeCount $dnsRecords)
        RegistryExportFiles  = @($registryExports.File)
        Notes                = @(
            'Use restore_dns.ps1 to attempt DNS restoration.',
            'Registry .reg files can be imported manually if needed.'
        )
    } | ConvertTo-Json -Depth 4 | Set-Content -Path $manifestPath -Encoding UTF8

    Write-KobraModuleLog -Log $Log -Message 'Backup bundle complete.'

    return [pscustomobject]@{
        BackupPath = $backupDir
    }
}

function Invoke-KobraGuard {
    [CmdletBinding()]
    param([scriptblock]$Log)

    Write-KobraModuleLog -Log $Log -Message 'Creating system restore point...'

    try {
        Enable-ComputerRestore -Drive "$($env:SystemDrive)\" -ErrorAction Stop
    }
    catch {
        Write-KobraModuleLog -Log $Log -Message '  System protection may already be enabled or could not be changed.'
    }

    try {
        Checkpoint-Computer -Description 'KobraOptimizer Safety Guard' -RestorePointType 'MODIFY_SETTINGS' -ErrorAction Stop
        Write-KobraModuleLog -Log $Log -Message '  Restore point created.'
    }
    catch {
        Write-KobraModuleLog -Log $Log -Message ("  Restore point creation reported: {0}" -f $_.Exception.Message)
    }
}

function Invoke-KobraTlsHardening {
    [CmdletBinding()]
    param([scriptblock]$Log)

    Write-KobraModuleLog -Log $Log -Message 'Applying TLS hardening profile...'

    $base = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols'
    $protocols = @(
        @{ Name = 'SSL 2.0'; Enabled = 0; DisabledByDefault = 1 },
        @{ Name = 'SSL 3.0'; Enabled = 0; DisabledByDefault = 1 },
        @{ Name = 'TLS 1.0'; Enabled = 0; DisabledByDefault = 1 },
        @{ Name = 'TLS 1.1'; Enabled = 0; DisabledByDefault = 1 },
        @{ Name = 'TLS 1.2'; Enabled = 1; DisabledByDefault = 0 },
        @{ Name = 'TLS 1.3'; Enabled = 1; DisabledByDefault = 0 }
    )

    foreach ($protocol in $protocols) {
        foreach ($role in @('Client','Server')) {
            $path = Join-Path $base ("{0}\{1}" -f $protocol.Name, $role)
            Set-KobraRegDword -Path $path -Name 'Enabled' -Value ([uint32]$protocol.Enabled)
            Set-KobraRegDword -Path $path -Name 'DisabledByDefault' -Value ([uint32]$protocol.DisabledByDefault)
        }
        Write-KobraModuleLog -Log $Log -Message ("  Protocol set: {0}" -f $protocol.Name)
    }

    $netFrameworkKeys = @(
        'HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319'
    )

    foreach ($path in $netFrameworkKeys) {
        Set-KobraRegDword -Path $path -Name 'SchUseStrongCrypto' -Value 1
        Set-KobraRegDword -Path $path -Name 'SystemDefaultTlsVersions' -Value 1
    }

    Write-KobraModuleLog -Log $Log -Message '.NET strong crypto flags enabled.'
    Write-KobraModuleLog -Log $Log -Message 'TLS hardening complete.'
}

function Invoke-KobraTcpStrike {
    [CmdletBinding()]
    param([scriptblock]$Log)

    Write-KobraModuleLog -Log $Log -Message 'Applying network TCP strike...'

    Invoke-KobraNative -FilePath 'netsh.exe' -ArgumentList @('int','tcp','set','global','ecncapability=enabled') -Log $Log
    Invoke-KobraNative -FilePath 'netsh.exe' -ArgumentList @('int','tcp','set','global','autotuninglevel=normal') -Log $Log
    Invoke-KobraNative -FilePath 'netsh.exe' -ArgumentList @('int','tcp','set','supplemental','template=internet','congestionprovider=cubic') -Log $Log

    $multimediaPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile'
    Set-KobraRegDword -Path $multimediaPath -Name 'NetworkThrottlingIndex' -Value 0xffffffff
    Set-KobraRegDword -Path $multimediaPath -Name 'SystemResponsiveness' -Value 0

    Write-KobraModuleLog -Log $Log -Message 'Registry network profile tuned.'
    Write-KobraModuleLog -Log $Log -Message 'Network TCP strike complete.'
}

function Invoke-KobraDnsFlush {
    [CmdletBinding()]
    param([scriptblock]$Log)

    Write-KobraModuleLog -Log $Log -Message 'Flushing DNS cache...'

    try {
        Clear-DnsClientCache -ErrorAction Stop
        Write-KobraModuleLog -Log $Log -Message '  DNS client cache cleared.'
    }
    catch {
        Write-KobraModuleLog -Log $Log -Message '  Clear-DnsClientCache not available or failed, trying ipconfig.'
    }

    $output = & ipconfig.exe /flushdns 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-KobraModuleLog -Log $Log -Message '  ipconfig /flushdns completed.'
    }
    else {
        Write-KobraModuleLog -Log $Log -Message ("  ipconfig /flushdns reported: {0}" -f (($output | Out-String).Trim()))
    }

    Write-KobraModuleLog -Log $Log -Message 'DNS flush complete.'
}

function Get-KobraDnsProfiles {
    [ordered]@{
        Cloudflare = [pscustomobject]@{
            Key         = 'Cloudflare'
            DisplayName = 'Cloudflare'
            IPv4        = @('1.1.1.1','1.0.0.1')
            IPv6        = @('2606:4700:4700::1111','2606:4700:4700::1001')
        }
        Google = [pscustomobject]@{
            Key         = 'Google'
            DisplayName = 'Google'
            IPv4        = @('8.8.8.8','8.8.4.4')
            IPv6        = @('2001:4860:4860::8888','2001:4860:4860::8844')
        }
        Automatic = [pscustomobject]@{
            Key         = 'Automatic'
            DisplayName = 'Automatic / DHCP'
            IPv4        = @()
            IPv6        = @()
        }
    }
}

function Get-KobraDnsPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Cloudflare','Google','Automatic')]
        [string]$ProfileName
    )

    $profiles = Get-KobraDnsProfiles
    return $profiles[$ProfileName]
}

function Get-KobraTargetAdapters {
    $configs = @(Get-NetIPConfiguration -ErrorAction SilentlyContinue | Where-Object {
        $_.NetAdapter.Status -eq 'Up' -and ($_.IPv4DefaultGateway -or $_.IPv6DefaultGateway -or $_.NetAdapter.HardwareInterface)
    })

    if ((Get-KobraSafeCount $configs) -eq 0) {
        $configs = @(Get-NetIPConfiguration -ErrorAction SilentlyContinue | Where-Object {
            $_.NetAdapter.Status -eq 'Up'
        })
    }

    return @($configs | Sort-Object InterfaceMetric, InterfaceAlias -Unique)
}

function Invoke-KobraDnsProfile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Cloudflare','Google','Automatic')]
        [string]$ProfileName,
        [scriptblock]$Log
    )

    $profile = Get-KobraDnsPlan -ProfileName $ProfileName
    $adapters = @(Get-KobraTargetAdapters)

    if ((Get-KobraSafeCount $adapters) -eq 0) {
        throw 'No active network adapters were found.'
    }

    Write-KobraModuleLog -Log $Log -Message ("Applying DNS profile: {0}" -f $profile.DisplayName)

    foreach ($adapter in $adapters) {
        $alias = $adapter.InterfaceAlias
        Write-KobraModuleLog -Log $Log -Message ("  Adapter: {0}" -f $alias)

        if ($ProfileName -eq 'Automatic') {
            try {
                Set-DnsClientServerAddress -InterfaceAlias $alias -ResetServerAddresses -ErrorAction Stop
                Write-KobraModuleLog -Log $Log -Message '    DNS reset to automatic (DHCP).'
            }
            catch {
                Write-KobraModuleLog -Log $Log -Message ("    Could not reset DNS: {0}" -f $_.Exception.Message)
            }
            continue
        }

        try {
            Set-DnsClientServerAddress -InterfaceAlias $alias -ServerAddresses $profile.IPv4 -AddressFamily IPv4 -ErrorAction Stop
            Write-KobraModuleLog -Log $Log -Message ("    IPv4 DNS set: {0}" -f ($profile.IPv4 -join ', '))
        }
        catch {
            Write-KobraModuleLog -Log $Log -Message ("    IPv4 DNS update failed: {0}" -f $_.Exception.Message)
        }

        try {
            Set-DnsClientServerAddress -InterfaceAlias $alias -ServerAddresses $profile.IPv6 -AddressFamily IPv6 -ErrorAction Stop
            Write-KobraModuleLog -Log $Log -Message ("    IPv6 DNS set: {0}" -f ($profile.IPv6 -join ', '))
        }
        catch {
            Write-KobraModuleLog -Log $Log -Message '    IPv6 DNS update skipped or not supported on this adapter.'
        }
    }

    Write-KobraModuleLog -Log $Log -Message 'DNS profile operation complete.'
}

Export-ModuleMember -Function Invoke-KobraSettingsBackup, Invoke-KobraGuard, Invoke-KobraTlsHardening, Invoke-KobraTcpStrike, Invoke-KobraDnsFlush, Get-KobraDnsPlan, Invoke-KobraDnsProfile
