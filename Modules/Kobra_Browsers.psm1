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

function Get-KobraBrowserCatalog {
    [ordered]@{
        Chrome = @{
            Name        = 'Google Chrome'
            ProcessName = 'chrome'
            Root        = "$env:LOCALAPPDATA\Google\Chrome\User Data"
            ProfileDirs = @('Default','Profile *')
            CacheDirs   = @('Cache','Code Cache','GPUCache','DawnCache','GrShaderCache','Service Worker\CacheStorage')
            RootDirs    = @('Crashpad')
        }
        Edge = @{
            Name        = 'Microsoft Edge'
            ProcessName = 'msedge'
            Root        = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
            ProfileDirs = @('Default','Profile *')
            CacheDirs   = @('Cache','Code Cache','GPUCache','DawnCache','GrShaderCache','Service Worker\CacheStorage')
            RootDirs    = @('Crashpad')
        }
        Firefox = @{
            Name        = 'Firefox'
            ProcessName = 'firefox'
            Root        = "$env:APPDATA\Mozilla\Firefox\Profiles"
            ProfileDirs = @('*')
            CacheDirs   = @('cache2','startupCache','thumbnails','shader-cache')
            RootDirs    = @()
        }
    }
}

function Get-KobraBrowserCachePaths {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Chrome','Edge','Firefox')]
        [string]$Browser
    )

    $catalog = Get-KobraBrowserCatalog
    $entry   = $catalog[$Browser]
    $paths   = New-Object System.Collections.Generic.List[string]

    if (-not (Test-Path -LiteralPath $entry.Root)) {
        return @()
    }

    foreach ($profilePattern in $entry.ProfileDirs) {
        $profiles = Get-ChildItem -LiteralPath $entry.Root -Directory -Force -ErrorAction SilentlyContinue | Where-Object {
            $_.Name -like $profilePattern
        }

        foreach ($profile in $profiles) {
            foreach ($cacheDir in $entry.CacheDirs) {
                $fullPath = Join-Path $profile.FullName $cacheDir
                if (Test-Path -LiteralPath $fullPath) {
                    $paths.Add($fullPath)
                }
            }
        }
    }

    foreach ($rootDir in $entry.RootDirs) {
        $fullPath = Join-Path $entry.Root $rootDir
        if (Test-Path -LiteralPath $fullPath) {
            $paths.Add($fullPath)
        }
    }

    return $paths.ToArray()
}

function Get-KobraSizeForPaths {
    param([string[]]$Paths)

    $bytes = [int64]0
    foreach ($path in $Paths) {
        $items = Get-ChildItem -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            if (-not $item.PSIsContainer) {
                $bytes += [int64]$item.Length
            }
        }
    }
    return $bytes
}

function Get-KobraBrowserPreview {
    [CmdletBinding()]
    param([string[]]$Browsers)

    $catalog = Get-KobraBrowserCatalog
    $results = @()

    foreach ($browser in $Browsers) {
        if (-not $catalog.Contains($browser)) {
            continue
        }

        $paths = Get-KobraBrowserCachePaths -Browser $browser
        $bytes = Get-KobraSizeForPaths -Paths $paths

        $results += [pscustomobject]@{
            Key       = $browser
            Name      = $catalog[$browser].Name
            SizeBytes = $bytes
            SizeMB    = [Math]::Round(($bytes / 1MB), 2)
        }
    }

    return $results
}

function Clear-KobraPathContents {
    param(
        [string]$Path,
        [scriptblock]$Log
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return 0
    }

    $removed = 0
    $items = Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    foreach ($item in $items) {
        try {
            Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
            $removed++
        }
        catch {
        }
    }
    return $removed
}

function Test-KobraBrowserRunning {
    param([string]$ProcessName)

    return [bool](Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
}

function Invoke-KobraBrowserCleanup {
    [CmdletBinding()]
    param(
        [string[]]$Browsers,
        [scriptblock]$Log
    )

    if (-not $Browsers -or (Get-KobraSafeCount $Browsers) -eq 0) {
        Write-KobraModuleLog -Log $Log -Message 'No browsers selected for cleanup.'
        return
    }

    $catalog = Get-KobraBrowserCatalog
    Write-KobraModuleLog -Log $Log -Message 'Starting browser cleanup...'

    foreach ($browser in $Browsers) {
        if (-not $catalog.Contains($browser)) {
            continue
        }

        $entry = $catalog[$browser]
        Write-KobraModuleLog -Log $Log -Message ("Cleaning browser cache: {0}" -f $entry.Name)

        if (Test-KobraBrowserRunning -ProcessName $entry.ProcessName) {
            Write-KobraModuleLog -Log $Log -Message '  Browser appears to be open. Locked files may be skipped.'
        }

        $paths = Get-KobraBrowserCachePaths -Browser $browser
        if (-not $paths -or (Get-KobraSafeCount $paths) -eq 0) {
            Write-KobraModuleLog -Log $Log -Message '  No cache paths found.'
            continue
        }

        $removed = 0
        foreach ($path in $paths) {
            $removed += Clear-KobraPathContents -Path $path -Log $Log
        }

        Write-KobraModuleLog -Log $Log -Message ("  Cleaned {0} entries." -f $removed)
    }

    Write-KobraModuleLog -Log $Log -Message 'Browser cleanup complete.'
}


function Get-KobraBrowserCandidates {
    [CmdletBinding()]
    param([string[]]$Browsers)

    $catalog = Get-KobraBrowserCatalog
    $results = @()

    foreach ($browser in $Browsers) {
        if (-not $catalog.Contains($browser)) {
            continue
        }

        $paths = Get-KobraBrowserCachePaths -Browser $browser
        foreach ($path in $paths) {
            $items = Get-ChildItem -LiteralPath $path -Recurse -File -Force -ErrorAction SilentlyContinue
            foreach ($item in $items) {
                $results += [pscustomobject]@{
                    Category  = $catalog[$browser].Name
                    Path      = $item.FullName
                    SizeBytes = [int64]$item.Length
                }
            }
        }
    }

    return $results
}

Export-ModuleMember -Function Get-KobraBrowserPreview, Get-KobraBrowserCandidates, Invoke-KobraBrowserCleanup
