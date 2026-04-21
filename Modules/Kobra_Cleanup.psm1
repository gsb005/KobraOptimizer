
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

function Get-KobraItemBytes {
    param($Item)

    if ($null -eq $Item) {
        return [int64]0
    }

    try {
        if (-not $Item.PSIsContainer) {
            return [int64]$Item.Length
        }

        $sum = (Get-ChildItem -LiteralPath $Item.FullName -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        if ($null -eq $sum) {
            return [int64]0
        }
        return [int64]$sum
    }
    catch {
        return [int64]0
    }
}

function Remove-KobraEmptyDirectories {
    param([string]$RootPath)

    if (-not (Test-Path -LiteralPath $RootPath)) {
        return
    }

    $dirs = @(Get-ChildItem -LiteralPath $RootPath -Recurse -Directory -Force -ErrorAction SilentlyContinue | Sort-Object FullName -Descending)
    foreach ($dir in $dirs) {
        try {
            if ((Get-ChildItem -LiteralPath $dir.FullName -Force -ErrorAction SilentlyContinue | Measure-Object).Count -eq 0) {
                Remove-Item -LiteralPath $dir.FullName -Force -ErrorAction Stop
            }
        }
        catch {
        }
    }
}

function Get-KobraFailureKind {
    param([string]$Message)

    if ([string]::IsNullOrWhiteSpace($Message)) {
        return 'failed'
    }

    $text = $Message.ToLowerInvariant()
    if ($text.Contains('used by another process') -or $text.Contains('being used by another process') -or $text.Contains('cannot access the file') -or $text.Contains('because it is being used')) {
        return 'locked'
    }

    if ($text.Contains('access to the path') -or $text.Contains('access is denied') -or $text.Contains('permission')) {
        return 'permission'
    }

    return 'failed'
}

function New-KobraCleanupResult {
    param(
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

    return [pscustomobject]@{
        Section        = 'System'
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
        return (New-KobraCleanupResult -Category $Label -Reason 'path not found')
    }

    $found = 0
    $attempted = 0
    $removed = 0
    $skipped = 0
    $failed = 0
    $locked = 0
    [int64]$foundBytes = 0
    [int64]$removedBytes = 0

    $files = @{}

    if ($Include -and (Get-KobraSafeCount $Include) -gt 0) {
        foreach ($pattern in $Include) {
            $items = @(Get-ChildItem -LiteralPath $Path -Filter $pattern -Force -ErrorAction SilentlyContinue)
            foreach ($item in $items) {
                if ($item.PSIsContainer) {
                    foreach ($child in @(Get-ChildItem -LiteralPath $item.FullName -Recurse -File -Force -ErrorAction SilentlyContinue)) {
                        $files[$child.FullName] = $child
                    }
                }
                else {
                    $files[$item.FullName] = $item
                }
            }
        }
    }
    else {
        foreach ($item in @(Get-ChildItem -LiteralPath $Path -Recurse -File -Force -ErrorAction SilentlyContinue)) {
            $files[$item.FullName] = $item
        }
    }

    foreach ($itemPath in ($files.Keys | Sort-Object)) {
        $item = $files[$itemPath]
        $found++
        $attempted++
        $itemBytes = Get-KobraItemBytes -Item $item
        $foundBytes += $itemBytes

        try {
            Remove-Item -LiteralPath $item.FullName -Force -ErrorAction Stop
            $removed++
            $removedBytes += $itemBytes
        }
        catch {
            $kind = Get-KobraFailureKind -Message $_.Exception.Message
            switch ($kind) {
                'locked' {
                    $locked++
                    $skipped++
                    Write-KobraModuleLog -Log $Log -Message ("  {0}: skipped locked item {1}" -f $Label, $item.FullName)
                }
                'permission' {
                    $failed++
                    Write-KobraModuleLog -Log $Log -Message ("  {0}: failed item {1} (permission issue)" -f $Label, $item.FullName)
                }
                default {
                    $failed++
                    Write-KobraModuleLog -Log $Log -Message ("  {0}: failed item {1}" -f $Label, $item.FullName)
                }
            }
        }
    }

    Remove-KobraEmptyDirectories -RootPath $Path

    Write-KobraModuleLog -Log $Log -Message ("  {0}: found={1}; attempted={2}; removed={3}; skipped={4}; failed={5}; locked={6}" -f $Label, $found, $attempted, $removed, $skipped, $failed, $locked)

    return (New-KobraCleanupResult -Category $Label -FoundCount $found -AttemptedCount $attempted -RemovedCount $removed -SkippedCount $skipped -FailedCount $failed -LockedCount $locked -FoundBytes $foundBytes -RemovedBytes $removedBytes)
}

function Clear-KobraWindowsUpdateCache {
    param([scriptblock]$Log)

    $services = @('wuauserv','bits','dosvc')
    $stopped  = @()

    foreach ($serviceName in $services) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($null -ne $service -and $service.Status -ne 'Stopped') {
            try {
                Stop-Service -Name $serviceName -Force -WarningAction SilentlyContinue -ErrorAction Stop
                $stopped += $serviceName
                Write-KobraModuleLog -Log $Log -Message ("  Stopped service: {0}" -f $serviceName)
            }
            catch {
                Write-KobraModuleLog -Log $Log -Message ("  Could not stop service: {0}" -f $serviceName)
            }
        }
    }

    $result = Clear-KobraDirectoryContents -Path "$env:windir\SoftwareDistribution\Download" -Log $Log -Label 'Windows Update Cache'

    foreach ($serviceName in $stopped) {
        try {
            Start-Service -Name $serviceName -WarningAction SilentlyContinue -ErrorAction Stop
            Write-KobraModuleLog -Log $Log -Message ("  Restarted service: {0}" -f $serviceName)
        }
        catch {
            Write-KobraModuleLog -Log $Log -Message ("  Could not restart service: {0}" -f $serviceName)
        }
    }

    return $result
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

    $results = @()

    foreach ($target in $Targets) {
        if (-not $catalog.Contains($target)) {
            continue
        }

        $entry = $catalog[$target]
        Write-KobraModuleLog -Log $Log -Message ("Cleaning: {0}" -f $entry.Name)

        switch ($target) {
            'UserTemp' {
                $results += @(Clear-KobraDirectoryContents -Path "$env:LOCALAPPDATA\Temp" -Log $Log -Label $entry.Name)
            }
            'SystemTemp' {
                $results += @(Clear-KobraDirectoryContents -Path "$env:windir\Temp" -Log $Log -Label $entry.Name)
            }
            'WindowsUpdate' {
                $results += @(Clear-KobraWindowsUpdateCache -Log $Log)
            }
            'ThumbnailCache' {
                $results += @(Clear-KobraDirectoryContents -Path "$env:LOCALAPPDATA\Microsoft\Windows\Explorer" -Include @('thumbcache*.db','iconcache*') -Log $Log -Label $entry.Name)
            }
            'ShaderCache' {
                $results += @(Clear-KobraDirectoryContents -Path "$env:LOCALAPPDATA\D3DSCache" -Log $Log -Label $entry.Name)
            }
            'RecycleBin' {
                $before = Get-KobraRecycleBinStats
                try {
                    Clear-RecycleBin -Force -Confirm:$false -ErrorAction Stop
                    Write-KobraModuleLog -Log $Log -Message '  Recycle Bin emptied.'
                    $after = Get-KobraRecycleBinStats
                    $removedCount = [Math]::Max(0, $before.Items - $after.Items)
                    $removedBytes = [Math]::Max([int64]0, $before.SizeBytes - $after.SizeBytes)
                    $results += @(New-KobraCleanupResult -Category $entry.Name -FoundCount $before.Items -AttemptedCount $before.Items -RemovedCount $removedCount -FoundBytes $before.SizeBytes -RemovedBytes $removedBytes)
                }
                catch {
                    Write-KobraModuleLog -Log $Log -Message '  Recycle Bin could not be fully emptied.'
                    $results += @(New-KobraCleanupResult -Category $entry.Name -FoundCount $before.Items -AttemptedCount $before.Items -SkippedCount $before.Items -FailedCount 1 -FoundBytes $before.SizeBytes -Reason 'recycle bin could not be fully emptied')
                }
            }
        }
    }

    Write-KobraModuleLog -Log $Log -Message 'System cleanup pass complete.'
    return $results
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
