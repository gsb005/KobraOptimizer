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
            Name           = 'Google Chrome'
            ProcessName    = 'chrome'
            Root           = "$env:LOCALAPPDATA\Google\Chrome\User Data"
            ProfileDirs    = @('Default','Profile *')
            CacheDirs      = @('Cache','Code Cache','GPUCache','DawnCache','GrShaderCache','Service Worker\CacheStorage')
            RootDirs       = @('Crashpad')
            CookieFiles    = @('Network\Cookies','Network\Cookies-journal','Network\Cookies-wal','Network\Cookies-shm','Cookies','Cookies-journal','Cookies-wal','Cookies-shm')
        }
        Edge = @{
            Name           = 'Microsoft Edge'
            ProcessName    = 'msedge'
            Root           = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
            ProfileDirs    = @('Default','Profile *')
            CacheDirs      = @('Cache','Code Cache','GPUCache','DawnCache','GrShaderCache','Service Worker\CacheStorage')
            RootDirs       = @('Crashpad')
            CookieFiles    = @('Network\Cookies','Network\Cookies-journal','Network\Cookies-wal','Network\Cookies-shm','Cookies','Cookies-journal','Cookies-wal','Cookies-shm')
        }
        Firefox = @{
            Name           = 'Firefox'
            ProcessName    = 'firefox'
            Root           = "$env:APPDATA\Mozilla\Firefox\Profiles"
            ProfileDirs    = @('*')
            CacheDirs      = @('cache2','startupCache','thumbnails','shader-cache')
            RootDirs       = @()
            CookieFiles    = @('cookies.sqlite','cookies.sqlite-wal','cookies.sqlite-shm')
        }
    }
}

function Get-KobraBrowserComponentInfo {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Cache','Cookies')]
        [string]$Component
    )

    switch ($Component) {
        'Cache' {
            return [pscustomobject]@{
                Suffix = 'Cache'
                Note   = 'Temporary browser files'
            }
        }
        'Cookies' {
            return [pscustomobject]@{
                Suffix = 'Cookies'
                Note   = 'May sign you out of websites'
            }
        }
    }
}

function Get-KobraBrowserComponentItems {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Chrome','Edge','Firefox')]
        [string]$Browser,

        [Parameter(Mandatory)]
        [ValidateSet('Cache','Cookies')]
        [string]$Component
    )

    $catalog = Get-KobraBrowserCatalog
    $entry   = $catalog[$Browser]
    $results = @()

    if (-not (Test-Path -LiteralPath $entry.Root)) {
        return @()
    }

    $profiles = @()
    foreach ($profilePattern in $entry.ProfileDirs) {
        $profiles += Get-ChildItem -LiteralPath $entry.Root -Directory -Force -ErrorAction SilentlyContinue | Where-Object {
            $_.Name -like $profilePattern
        }
    }

    $profiles = @($profiles | Sort-Object FullName -Unique)

    switch ($Component) {
        'Cache' {
            foreach ($profile in $profiles) {
                foreach ($cacheDir in $entry.CacheDirs) {
                    $fullPath = Join-Path $profile.FullName $cacheDir
                    if (Test-Path -LiteralPath $fullPath) {
                        $results += [pscustomobject]@{
                            Path      = $fullPath
                            Kind      = 'Directory'
                            Category  = ('{0} Cache' -f $entry.Name)
                            Note      = 'Temporary browser files'
                        }
                    }
                }
            }

            foreach ($rootDir in $entry.RootDirs) {
                $fullPath = Join-Path $entry.Root $rootDir
                if (Test-Path -LiteralPath $fullPath) {
                    $results += [pscustomobject]@{
                        Path      = $fullPath
                        Kind      = 'Directory'
                        Category  = ('{0} Cache' -f $entry.Name)
                        Note      = 'Temporary browser files'
                    }
                }
            }
        }
        'Cookies' {
            foreach ($profile in $profiles) {
                foreach ($cookieFile in $entry.CookieFiles) {
                    $fullPath = Join-Path $profile.FullName $cookieFile
                    if (Test-Path -LiteralPath $fullPath) {
                        $results += [pscustomobject]@{
                            Path      = $fullPath
                            Kind      = 'File'
                            Category  = ('{0} Cookies' -f $entry.Name)
                            Note      = 'May sign you out of websites'
                        }
                    }
                }
            }
        }
    }

    return @($results | Sort-Object Path -Unique)
}

function Get-KobraFilesForBrowserItem {
    param([psobject]$Item)

    if ($null -eq $Item) {
        return @()
    }

    if ($Item.Kind -eq 'Directory') {
        return @(Get-ChildItem -LiteralPath $Item.Path -Recurse -File -Force -ErrorAction SilentlyContinue)
    }

    if (($Item.Kind -eq 'File') -and (Test-Path -LiteralPath $Item.Path)) {
        return @(Get-Item -LiteralPath $Item.Path -Force -ErrorAction SilentlyContinue)
    }

    return @()
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

function Get-KobraBrowserPreview {
    [CmdletBinding()]
    param(
        [string[]]$Browsers,
        [string[]]$Components = @('Cache')
    )

    $catalog = Get-KobraBrowserCatalog
    $results = @()

    foreach ($browser in $Browsers) {
        if (-not $catalog.Contains($browser)) {
            continue
        }

        foreach ($component in $Components) {
            $componentItems = @(Get-KobraBrowserComponentItems -Browser $browser -Component $component)
            $files = @()
            foreach ($item in $componentItems) {
                $files += @(Get-KobraFilesForBrowserItem -Item $item)
            }

            $bytes = [int64]0
            foreach ($file in $files) {
                if ($null -ne $file) {
                    $bytes += [int64]$file.Length
                }
            }

            $componentInfo = Get-KobraBrowserComponentInfo -Component $component
            $results += [pscustomobject]@{
                Key       = ('{0}_{1}' -f $browser, $component)
                Name      = ('{0} {1}' -f $catalog[$browser].Name, $componentInfo.Suffix)
                Items     = @($files).Count
                SizeBytes = $bytes
                SizeMB    = [Math]::Round(($bytes / 1MB), 2)
                Note      = $componentInfo.Note
            }
        }
    }

    return $results
}

function Invoke-KobraBrowserCleanup {
    [CmdletBinding()]
    param(
        [string[]]$Browsers,
        [string[]]$Components = @('Cache'),
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
        $isRunning = Test-KobraBrowserRunning -ProcessName $entry.ProcessName
        if ($isRunning) {
            Write-KobraModuleLog -Log $Log -Message ("{0} appears to be open. Locked files may be skipped." -f $entry.Name)
        }

        foreach ($component in $Components) {
            $componentInfo = Get-KobraBrowserComponentInfo -Component $component
            Write-KobraModuleLog -Log $Log -Message ("Cleaning browser {0}: {1}" -f $component.ToLower(), $entry.Name)
            $componentItems = @(Get-KobraBrowserComponentItems -Browser $browser -Component $component)

            if ((Get-KobraSafeCount $componentItems) -eq 0) {
                Write-KobraModuleLog -Log $Log -Message '  No matching browser data found.'
                continue
            }

            $removed = 0
            foreach ($item in $componentItems) {
                try {
                    if ($item.Kind -eq 'Directory') {
                        $removed += Clear-KobraPathContents -Path $item.Path -Log $Log
                    }
                    elseif ((Test-Path -LiteralPath $item.Path)) {
                        Remove-Item -LiteralPath $item.Path -Force -ErrorAction Stop
                        $removed++
                    }
                }
                catch {
                    Write-KobraModuleLog -Log $Log -Message ("  Skipped locked browser item: {0}" -f $item.Path)
                }
            }

            Write-KobraModuleLog -Log $Log -Message ("  {0}: removed {1} records. {2}" -f $componentInfo.Suffix, $removed, $componentInfo.Note)
        }
    }

    Write-KobraModuleLog -Log $Log -Message 'Browser cleanup complete.'
}

function Get-KobraBrowserCandidates {
    [CmdletBinding()]
    param(
        [string[]]$Browsers,
        [string[]]$Components = @('Cache')
    )

    $catalog = Get-KobraBrowserCatalog
    $results = @()

    foreach ($browser in $Browsers) {
        if (-not $catalog.Contains($browser)) {
            continue
        }

        foreach ($component in $Components) {
            $componentItems = @(Get-KobraBrowserComponentItems -Browser $browser -Component $component)
            foreach ($item in $componentItems) {
                $files = @(Get-KobraFilesForBrowserItem -Item $item)
                foreach ($file in $files) {
                    $results += [pscustomobject]@{
                        Category  = $item.Category
                        Path      = $file.FullName
                        SizeBytes = [int64]$file.Length
                        Note      = $item.Note
                    }
                }
            }
        }
    }

    return $results
}

Export-ModuleMember -Function Get-KobraBrowserPreview, Get-KobraBrowserCandidates, Invoke-KobraBrowserCleanup
