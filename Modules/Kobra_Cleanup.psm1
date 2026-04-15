
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

function Get-KobraCleanupCatalog {
    return [ordered]@{
        UserTemp = @{
            Name    = 'User Temp'
            Paths   = @("$env:LOCALAPPDATA\Temp")
            Include = @()
            Special = $null
        }
        SystemTemp = @{
            Name    = 'System Temp'
            Paths   = @("$env:windir\Temp")
            Include = @()
            Special = $null
        }
        WindowsUpdate = @{
            Name    = 'Windows Update Cache'
            Paths   = @("$env:windir\SoftwareDistribution\Download")
            Include = @()
            Special = $null
        }
        ThumbnailCache = @{
            Name    = 'Thumbnail / Icon Cache'
            Paths   = @("$env:LOCALAPPDATA\Microsoft\Windows\Explorer")
            Include = @('thumbcache*.db','iconcache*')
            Special = $null
        }
        ShaderCache = @{
            Name    = 'DirectX Shader Cache'
            Paths   = @("$env:LOCALAPPDATA\D3DSCache")
            Include = @()
            Special = $null
        }
        RecycleBin = @{
            Name    = 'Recycle Bin'
            Paths   = @()
            Include = @()
            Special = 'RecycleBin'
        }
    }
}

function Get-KobraFolderStats {
    param(
        [string[]]$Paths,
        [string[]]$Include
    )

    $totalBytes = [int64]0
    $totalItems = 0

    foreach ($path in $Paths) {
        if (-not (Test-Path -LiteralPath $path)) {
            continue
        }

        if ($Include -and (Get-KobraSafeCount $Include) -gt 0) {
            foreach ($pattern in $Include) {
                $items = Get-ChildItem -LiteralPath $path -Filter $pattern -Force -ErrorAction SilentlyContinue
                foreach ($item in $items) {
                    $totalItems++
                    if (-not $item.PSIsContainer) {
                        $totalBytes += [int64]$item.Length
                    }
                }
            }
        }
        else {
            $items = Get-ChildItem -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue
            foreach ($item in $items) {
                $totalItems++
                if (-not $item.PSIsContainer) {
                    $totalBytes += [int64]$item.Length
                }
            }
        }
    }

    [pscustomobject]@{
        Items     = $totalItems
        SizeBytes = $totalBytes
        SizeMB    = [Math]::Round(($totalBytes / 1MB), 2)
    }
}

function Get-KobraRecycleBinStats {
    $totalBytes = [int64]0
    $totalItems = 0

    foreach ($drive in (Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue)) {
        $recyclePath = Join-Path $drive.Root '$Recycle.Bin'
        if (-not (Test-Path -LiteralPath $recyclePath)) {
            continue
        }

        $items = Get-ChildItem -LiteralPath $recyclePath -Recurse -Force -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            $totalItems++
            if (-not $item.PSIsContainer) {
                $totalBytes += [int64]$item.Length
            }
        }
    }

    [pscustomobject]@{
        Items     = $totalItems
        SizeBytes = $totalBytes
        SizeMB    = [Math]::Round(($totalBytes / 1MB), 2)
    }
}

function Get-KobraCleanupPreview {
    [CmdletBinding()]
    param(
        [string[]]$Targets
    )

    $catalog = Get-KobraCleanupCatalog
    $results = @()

    foreach ($target in $Targets) {
        if (-not $catalog.Contains($target)) {
            continue
        }

        $entry = $catalog[$target]
        $stats = if ($entry['Special'] -eq 'RecycleBin') {
            Get-KobraRecycleBinStats
        }
        else {
            Get-KobraFolderStats -Paths $entry.Paths -Include $entry.Include
        }

        $results += [pscustomobject]@{
            Key       = $target
            Name      = $entry.Name
            Items     = $stats.Items
            SizeBytes = $stats.SizeBytes
            SizeMB    = $stats.SizeMB
        }
    }

    return $results
}

function Clear-KobraDirectoryContents {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [string[]]$Include,
        [scriptblock]$Log,
        [string]$Label
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-KobraModuleLog -Log $Log -Message ("  {0}: path not found, skipping." -f $Label)
        return
    }

    $removed = 0

    if ($Include -and (Get-KobraSafeCount $Include) -gt 0) {
        foreach ($pattern in $Include) {
            $items = Get-ChildItem -LiteralPath $Path -Filter $pattern -Force -ErrorAction SilentlyContinue
            foreach ($item in $items) {
                try {
                    Remove-Item -LiteralPath $item.FullName -Force -Recurse -ErrorAction Stop
                    $removed++
                }
                catch {
                    Write-KobraModuleLog -Log $Log -Message ("  {0}: skipped locked item {1}" -f $Label, $item.FullName)
                }
            }
        }
    }
    else {
        $items = Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            try {
                Remove-Item -LiteralPath $item.FullName -Force -Recurse -ErrorAction Stop
                $removed++
            }
            catch {
                Write-KobraModuleLog -Log $Log -Message ("  {0}: skipped locked item {1}" -f $Label, $item.FullName)
            }
        }
    }

    Write-KobraModuleLog -Log $Log -Message ("  {0}: removed {1} entries." -f $Label, $removed)
}

function Clear-KobraWindowsUpdateCache {
    param([scriptblock]$Log)

    $services = @('wuauserv','bits','dosvc')
    $stopped  = @()

    foreach ($serviceName in $services) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($null -ne $service -and $service.Status -ne 'Stopped') {
            try {
                Stop-Service -Name $serviceName -Force -ErrorAction Stop
                $stopped += $serviceName
                Write-KobraModuleLog -Log $Log -Message ("  Stopped service: {0}" -f $serviceName)
            }
            catch {
                Write-KobraModuleLog -Log $Log -Message ("  Could not stop service: {0}" -f $serviceName)
            }
        }
    }

    Clear-KobraDirectoryContents -Path "$env:windir\SoftwareDistribution\Download" -Log $Log -Label 'Windows Update Cache'

    foreach ($serviceName in $stopped) {
        try {
            Start-Service -Name $serviceName -ErrorAction Stop
            Write-KobraModuleLog -Log $Log -Message ("  Restarted service: {0}" -f $serviceName)
        }
        catch {
            Write-KobraModuleLog -Log $Log -Message ("  Could not restart service: {0}" -f $serviceName)
        }
    }
}

function Invoke-KobraShed {
    [CmdletBinding()]
    param(
        [string[]]$Targets,
        [scriptblock]$Log
    )

    if (-not $Targets -or (Get-KobraSafeCount $Targets) -eq 0) {
        Write-KobraModuleLog -Log $Log -Message 'No cleanup targets were supplied.'
        return
    }

    $catalog = Get-KobraCleanupCatalog
    Write-KobraModuleLog -Log $Log -Message 'Starting system cleanup pass...'

    foreach ($target in $Targets) {
        if (-not $catalog.Contains($target)) {
            continue
        }

        $entry = $catalog[$target]
        Write-KobraModuleLog -Log $Log -Message ("Cleaning: {0}" -f $entry.Name)

        switch ($target) {
            'UserTemp' {
                Clear-KobraDirectoryContents -Path "$env:LOCALAPPDATA\Temp" -Log $Log -Label $entry.Name
            }
            'SystemTemp' {
                Clear-KobraDirectoryContents -Path "$env:windir\Temp" -Log $Log -Label $entry.Name
            }
            'WindowsUpdate' {
                Clear-KobraWindowsUpdateCache -Log $Log
            }
            'ThumbnailCache' {
                Clear-KobraDirectoryContents -Path "$env:LOCALAPPDATA\Microsoft\Windows\Explorer" -Include @('thumbcache*.db','iconcache*') -Log $Log -Label $entry.Name
            }
            'ShaderCache' {
                Clear-KobraDirectoryContents -Path "$env:LOCALAPPDATA\D3DSCache" -Log $Log -Label $entry.Name
            }
            'RecycleBin' {
                try {
                    Clear-RecycleBin -Force -Confirm:$false -ErrorAction Stop
                    Write-KobraModuleLog -Log $Log -Message '  Recycle Bin emptied.'
                }
                catch {
                    Write-KobraModuleLog -Log $Log -Message '  Recycle Bin could not be fully emptied.'
                }
            }
        }
    }

    Write-KobraModuleLog -Log $Log -Message 'System cleanup pass complete.'
}

function Get-KobraFileCandidatesFromPath {
    param(
        [string]$Path,
        [string[]]$Include,
        [string]$Category
    )

    $results = @()
    if (-not (Test-Path -LiteralPath $Path)) {
        return $results
    }

    if ($Include -and (Get-KobraSafeCount $Include) -gt 0) {
        foreach ($pattern in $Include) {
            $items = Get-ChildItem -LiteralPath $Path -Filter $pattern -File -Force -ErrorAction SilentlyContinue
            foreach ($item in $items) {
                $results += [pscustomobject]@{
                    Category  = $Category
                    Path      = $item.FullName
                    SizeBytes = [int64]$item.Length
                }
            }
        }
    }
    else {
        $items = Get-ChildItem -LiteralPath $Path -Recurse -File -Force -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            $results += [pscustomobject]@{
                Category  = $Category
                Path      = $item.FullName
                SizeBytes = [int64]$item.Length
            }
        }
    }

    return $results
}

function Get-KobraRecycleBinCandidates {
    $results = @()

    foreach ($drive in (Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue)) {
        $recyclePath = Join-Path $drive.Root '$Recycle.Bin'
        if (-not (Test-Path -LiteralPath $recyclePath)) {
            continue
        }

        $items = Get-ChildItem -LiteralPath $recyclePath -Recurse -File -Force -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            $results += [pscustomobject]@{
                Category  = 'Recycle Bin'
                Path      = $item.FullName
                SizeBytes = [int64]$item.Length
            }
        }
    }

    return $results
}

function Get-KobraCleanupCandidates {
    [CmdletBinding()]
    param([string[]]$Targets)

    $catalog = Get-KobraCleanupCatalog
    $results = @()

    foreach ($target in $Targets) {
        if (-not $catalog.Contains($target)) {
            continue
        }

        $entry = $catalog[$target]
        if ($entry['Special'] -eq 'RecycleBin') {
            $results += @(Get-KobraRecycleBinCandidates)
        }
        else {
            foreach ($path in $entry.Paths) {
                $results += @(Get-KobraFileCandidatesFromPath -Path $path -Include $entry.Include -Category $entry.Name)
            }
        }
    }

    return $results
}

Export-ModuleMember -Function Get-KobraCleanupPreview, Get-KobraCleanupCandidates, Invoke-KobraShed
