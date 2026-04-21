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

function New-KobraBrowserResult {
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
        Section        = 'Browser'
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

function Get-KobraChromiumBrowser {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string[]]$ProcessNames,
        [Parameter(Mandatory)][string[]]$ProfileRoots
    )

    return @{
        Name         = $Name
        Engine       = 'Chromium'
        ProcessNames = $ProcessNames
        ProfileRoots = $ProfileRoots
        ProfileDirs  = @('Default','Profile *')
        RootDirs     = @()
        CacheDirs    = @(
            'Cache',
            'Media Cache',
            'GPUCache',
            'ShaderCache',
            'Code Cache',
            'Service Worker\CacheStorage',
            'Application Cache',
            'File System',
            'Storage\ext\*\def\Cache',
            'Storage\ext\*\def\Media Cache',
            'Storage\ext\*\def\GPUCache',
            'Storage\ext\*\def\ShaderCache',
            'Storage\ext\*\def\Code Cache',
            'Storage\ext\*\def\Platform Notifications',
            'Platform Notifications',
            'component_crx_cache',
            'GraphiteDawnCache',
            'GrShaderCache',
            'DawnCache'
        )
        CacheFiles   = @('TopSites.json')
        CookieFiles  = @(
            'Network\Cookies*',
            'Cookies*',
            'Origin Bound Certs',
            'QuotaManager',
            'Extension Cookies'
        )
        CookieDirs   = @(
            'Local Storage',
            'WebStorage',
            'IndexedDB',
            'databases'
        )
        HistoryFiles = @(
            'History',
            'History Index*.*',
            'Archived History',
            'Visited Links',
            'Current Tabs',
            'Last Tabs',
            'Top Sites',
            'History Provider Cache',
            'Network Action Predictor',
            'Shortcuts',
            'DownloadMetadata'
        )
        HistoryDirs  = @()
    }
}

function Get-KobraBrowserCatalog {
    [ordered]@{
        Chrome  = Get-KobraChromiumBrowser -Name 'Google Chrome' -ProcessNames @('chrome') -ProfileRoots @(
            "$env:LOCALAPPDATA\Google\Chrome\User Data",
            "$env:LOCALAPPDATA\Google\Chrome Beta\User Data",
            "$env:LOCALAPPDATA\Google\Chrome SxS\User Data",
            "$env:LOCALAPPDATA\Chromium\User Data"
        )
        Edge    = Get-KobraChromiumBrowser -Name 'Microsoft Edge' -ProcessNames @('msedge') -ProfileRoots @(
            "$env:LOCALAPPDATA\Microsoft\Edge\User Data",
            "$env:LOCALAPPDATA\Microsoft\Edge Beta\User Data",
            "$env:LOCALAPPDATA\Microsoft\Edge Dev\User Data",
            "$env:LOCALAPPDATA\Microsoft\Edge SxS\User Data"
        )
        Opera   = Get-KobraChromiumBrowser -Name 'Opera' -ProcessNames @('opera') -ProfileRoots @(
            "$env:APPDATA\Opera Software\Opera Stable",
            "$env:LOCALAPPDATA\Opera Software\Opera Stable",
            "$env:APPDATA\Opera Software\Opera Next",
            "$env:LOCALAPPDATA\Opera Software\Opera Next",
            "$env:APPDATA\Opera Software\Opera Developer",
            "$env:LOCALAPPDATA\Opera Software\Opera Developer",
            "$env:APPDATA\Opera Software\Opera GX Stable",
            "$env:LOCALAPPDATA\Opera Software\Opera GX Stable"
        )
        Brave   = Get-KobraChromiumBrowser -Name 'Brave Browser' -ProcessNames @('brave') -ProfileRoots @(
            "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data",
            "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser-Beta\User Data",
            "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser-Nightly\User Data"
        )
        Vivaldi = Get-KobraChromiumBrowser -Name 'Vivaldi' -ProcessNames @('vivaldi') -ProfileRoots @(
            "$env:LOCALAPPDATA\Vivaldi\User Data"
        )
        Firefox = @{
            Name         = 'Firefox'
            Engine       = 'Mozilla'
            ProcessNames = @('firefox')
            ProfileRoots = @(
                "$env:APPDATA\Mozilla\Firefox\Profiles",
                "$env:LOCALAPPDATA\Mozilla\Firefox\Profiles",
                "$env:LOCALAPPDATA\Packages\Mozilla.Firefox_n80bbvh6b1yt2\LocalCache\Roaming\Mozilla\Firefox\Profiles",
                "$env:LOCALAPPDATA\Packages\Mozilla.Firefox_n80bbvh6b1yt2\LocalCache\Local\Mozilla\Firefox\Profiles"
            )
            ProfileDirs  = @('*')
            RootDirs     = @(
                "$env:COMMONAPPDATA\Mozilla-*"
            )
            CacheDirs    = @(
                'cache',
                'cache2',
                'cache.trash*',
                'jumpListCache',
                'startupCache',
                'OfflineCache',
                'safebrowsing'
            )
            CacheFiles   = @()
            CookieFiles  = @(
                'cookies.sqlite*',
                'permissions.sqlite*',
                'webappsstore.sqlite*'
            )
            CookieDirs   = @(
                'indexedDB',
                'storage\persistent',
                'storage\permanent',
                'storage\default',
                'storage\temporary'
            )
            HistoryFiles = @(
                'history.dat',
                'urlbarhistory.sqlite*',
                'downloads.rdf',
                'downloads.sqlite*',
                'downloads.json'
            )
            HistoryDirs  = @('thumbnails')
        }
    }
}

function Get-KobraBrowserComponentInfo {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Cache','Cookies','History')]
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
        'History' {
            return [pscustomobject]@{
                Suffix = 'History'
                Note   = 'Browsing and download history where safe to remove'
            }
        }
    }
}

function Resolve-KobraBrowserMatches {
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][string]$Pattern,
        [Parameter(Mandatory)][ValidateSet('File','Directory')][string]$ExpectedType
    )

    if (-not (Test-Path -LiteralPath $BasePath)) {
        return @()
    }

    $searchPath = Join-Path $BasePath $Pattern
    $hasWildcard = ($Pattern.IndexOf('*') -ge 0) -or ($Pattern.IndexOf('?') -ge 0) -or ($Pattern.IndexOf('[') -ge 0)

    if (-not $hasWildcard) {
        if (-not (Test-Path -LiteralPath $searchPath)) {
            return @()
        }

        $item = Get-Item -LiteralPath $searchPath -Force -ErrorAction SilentlyContinue
        if ($null -eq $item) {
            return @()
        }

        if ($ExpectedType -eq 'File' -and -not $item.PSIsContainer) {
            return @($item)
        }

        if ($ExpectedType -eq 'Directory' -and $item.PSIsContainer) {
            return @($item)
        }

        return @()
    }

    $params = @{
        Path        = $searchPath
        Force       = $true
        ErrorAction = 'SilentlyContinue'
    }

    if ($ExpectedType -eq 'File') {
        $params.File = $true
    }
    else {
        $params.Directory = $true
    }

    return @(Get-ChildItem @params)
}

function Get-KobraBrowserProfilePaths {
    param([hashtable]$Entry)

    $profiles = @()
    foreach ($root in @($Entry.ProfileRoots)) {
        if (-not (Test-Path -LiteralPath $root)) {
            continue
        }

        foreach ($profilePattern in @($Entry.ProfileDirs)) {
            $profiles += @(Resolve-KobraBrowserMatches -BasePath $root -Pattern $profilePattern -ExpectedType 'Directory')
        }
    }

    return @($profiles | Sort-Object FullName -Unique)
}

function Get-KobraBrowserRootPaths {
    param([hashtable]$Entry)

    $roots = @()
    foreach ($root in @($Entry.ProfileRoots) + @($Entry.RootDirs)) {
        if ([string]::IsNullOrWhiteSpace($root)) {
            continue
        }

        if ($root.IndexOf('*') -ge 0 -or $root.IndexOf('?') -ge 0 -or $root.IndexOf('[') -ge 0) {
            $roots += @(Get-ChildItem -Path $root -Force -Directory -ErrorAction SilentlyContinue)
        }
        elseif (Test-Path -LiteralPath $root) {
            $roots += @(Get-Item -LiteralPath $root -Force -ErrorAction SilentlyContinue)
        }
    }

    return @($roots | Sort-Object FullName -Unique)
}

function New-KobraBrowserMatch {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][ValidateSet('File','Directory')][string]$Kind,
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][string]$Note
    )

    return [pscustomobject]@{
        Path     = $Path
        Kind     = $Kind
        Category = $Category
        Note     = $Note
    }
}

function Get-KobraBrowserComponentItems {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Chrome','Edge','Firefox','Opera','Brave','Vivaldi')]
        [string]$Browser,

        [Parameter(Mandatory)]
        [ValidateSet('Cache','Cookies','History')]
        [string]$Component
    )

    $catalog = Get-KobraBrowserCatalog
    $entry = $catalog[$Browser]
    $info = Get-KobraBrowserComponentInfo -Component $Component
    $category = ('{0} {1}' -f $entry.Name, $info.Suffix)
    $results = @()

    $profiles = @(Get-KobraBrowserProfilePaths -Entry $entry)
    $roots = @(Get-KobraBrowserRootPaths -Entry $entry)

    switch ($Component) {
        'Cache' {
            foreach ($profile in $profiles) {
                foreach ($pattern in @($entry.CacheDirs)) {
                    foreach ($match in @(Resolve-KobraBrowserMatches -BasePath $profile.FullName -Pattern $pattern -ExpectedType 'Directory')) {
                        $results += @(New-KobraBrowserMatch -Path $match.FullName -Kind 'Directory' -Category $category -Note $info.Note)
                    }
                }
                foreach ($pattern in @($entry.CacheFiles)) {
                    foreach ($match in @(Resolve-KobraBrowserMatches -BasePath $profile.FullName -Pattern $pattern -ExpectedType 'File')) {
                        $results += @(New-KobraBrowserMatch -Path $match.FullName -Kind 'File' -Category $category -Note $info.Note)
                    }
                }
            }

            foreach ($root in $roots) {
                foreach ($pattern in @('component_crx_cache','GraphiteDawnCache','GrShaderCache','ShaderCache')) {
                    foreach ($match in @(Resolve-KobraBrowserMatches -BasePath $root.FullName -Pattern $pattern -ExpectedType 'Directory')) {
                        $results += @(New-KobraBrowserMatch -Path $match.FullName -Kind 'Directory' -Category $category -Note $info.Note)
                    }
                }
            }
        }
        'Cookies' {
            foreach ($profile in $profiles) {
                foreach ($pattern in @($entry.CookieFiles)) {
                    foreach ($match in @(Resolve-KobraBrowserMatches -BasePath $profile.FullName -Pattern $pattern -ExpectedType 'File')) {
                        $results += @(New-KobraBrowserMatch -Path $match.FullName -Kind 'File' -Category $category -Note $info.Note)
                    }
                }
                foreach ($pattern in @($entry.CookieDirs)) {
                    foreach ($match in @(Resolve-KobraBrowserMatches -BasePath $profile.FullName -Pattern $pattern -ExpectedType 'Directory')) {
                        $results += @(New-KobraBrowserMatch -Path $match.FullName -Kind 'Directory' -Category $category -Note $info.Note)
                    }
                }
            }
        }
        'History' {
            foreach ($profile in $profiles) {
                foreach ($pattern in @($entry.HistoryFiles)) {
                    foreach ($match in @(Resolve-KobraBrowserMatches -BasePath $profile.FullName -Pattern $pattern -ExpectedType 'File')) {
                        $results += @(New-KobraBrowserMatch -Path $match.FullName -Kind 'File' -Category $category -Note $info.Note)
                    }
                }
                foreach ($pattern in @($entry.HistoryDirs)) {
                    foreach ($match in @(Resolve-KobraBrowserMatches -BasePath $profile.FullName -Pattern $pattern -ExpectedType 'Directory')) {
                        $results += @(New-KobraBrowserMatch -Path $match.FullName -Kind 'Directory' -Category $category -Note $info.Note)
                    }
                }
            }
        }
    }

    return @($results | Sort-Object Path, Kind -Unique)
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
        $file = Get-Item -LiteralPath $Item.Path -Force -ErrorAction SilentlyContinue
        if ($null -ne $file) {
            return @($file)
        }
    }

    return @()
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

function Test-KobraBrowserRunning {
    param([string[]]$ProcessNames)

    foreach ($processName in @($ProcessNames)) {
        if (Get-Process -Name $processName -ErrorAction SilentlyContinue) {
            return $true
        }
    }

    return $false
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
            $uniqueFiles = @{}
            [int64]$bytes = 0

            foreach ($item in $componentItems) {
                foreach ($file in @(Get-KobraFilesForBrowserItem -Item $item)) {
                    if ($null -eq $file -or $uniqueFiles.ContainsKey($file.FullName)) {
                        continue
                    }

                    $uniqueFiles[$file.FullName] = $true
                    $bytes += [int64]$file.Length
                }
            }

            $componentInfo = Get-KobraBrowserComponentInfo -Component $component
            $results += [pscustomobject]@{
                Key       = ('{0}_{1}' -f $browser, $component)
                Name      = ('{0} {1}' -f $catalog[$browser].Name, $componentInfo.Suffix)
                Items     = $uniqueFiles.Count
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
    $results = @()
    Write-KobraModuleLog -Log $Log -Message 'Starting browser cleanup...'

    foreach ($browser in $Browsers) {
        if (-not $catalog.Contains($browser)) {
            continue
        }

        $entry = $catalog[$browser]
        $isRunning = Test-KobraBrowserRunning -ProcessNames $entry.ProcessNames
        if ($isRunning) {
            Write-KobraModuleLog -Log $Log -Message ("{0} appears to be open. Locked files may be skipped." -f $entry.Name)
        }

        foreach ($component in $Components) {
            $componentInfo = Get-KobraBrowserComponentInfo -Component $component
            $category = ('{0} {1}' -f $entry.Name, $componentInfo.Suffix)
            Write-KobraModuleLog -Log $Log -Message ("Cleaning browser {0}: {1}" -f $component.ToLower(), $entry.Name)

            $componentItems = @(Get-KobraBrowserComponentItems -Browser $browser -Component $component)
            if ((Get-KobraSafeCount $componentItems) -eq 0) {
                Write-KobraModuleLog -Log $Log -Message '  No matching browser data found.'
                $results += @(New-KobraBrowserResult -Category $category -Reason $componentInfo.Note)
                continue
            }

            $filesByPath = @{}
            $directoryRoots = @{}
            foreach ($item in $componentItems) {
                if ($item.Kind -eq 'Directory') {
                    $directoryRoots[$item.Path] = $true
                }

                foreach ($file in @(Get-KobraFilesForBrowserItem -Item $item)) {
                    if ($null -ne $file) {
                        $filesByPath[$file.FullName] = $file
                    }
                }
            }

            if ($filesByPath.Count -eq 0) {
                Write-KobraModuleLog -Log $Log -Message '  No matching browser data found.'
                $results += @(New-KobraBrowserResult -Category $category -Reason $componentInfo.Note)
                continue
            }

            $foundCount = 0
            $attemptedCount = 0
            $removedCount = 0
            $skippedCount = 0
            $failedCount = 0
            $lockedCount = 0
            [int64]$foundBytes = 0
            [int64]$removedBytes = 0

            foreach ($path in ($filesByPath.Keys | Sort-Object)) {
                $file = $filesByPath[$path]
                $foundCount++
                $attemptedCount++
                $fileBytes = [int64]$file.Length
                $foundBytes += $fileBytes

                try {
                    Remove-Item -LiteralPath $path -Force -ErrorAction Stop
                    $removedCount++
                    $removedBytes += $fileBytes
                }
                catch {
                    $kind = Get-KobraFailureKind -Message $_.Exception.Message
                    switch ($kind) {
                        'locked' {
                            $lockedCount++
                            $skippedCount++
                            Write-KobraModuleLog -Log $Log -Message ("  Skipped locked browser item: {0}" -f $path)
                        }
                        'permission' {
                            $failedCount++
                            Write-KobraModuleLog -Log $Log -Message ("  Failed browser item (permission issue): {0}" -f $path)
                        }
                        default {
                            $failedCount++
                            Write-KobraModuleLog -Log $Log -Message ("  Failed browser item: {0}" -f $path)
                        }
                    }
                }
            }

            foreach ($root in $directoryRoots.Keys) {
                Remove-KobraEmptyDirectories -RootPath $root
            }

            Write-KobraModuleLog -Log $Log -Message ("  {0}: found={1}; attempted={2}; removed={3}; skipped={4}; failed={5}; locked={6}. {7}" -f $componentInfo.Suffix, $foundCount, $attemptedCount, $removedCount, $skippedCount, $failedCount, $lockedCount, $componentInfo.Note)
            $results += @(New-KobraBrowserResult -Category $category -FoundCount $foundCount -AttemptedCount $attemptedCount -RemovedCount $removedCount -SkippedCount $skippedCount -FailedCount $failedCount -LockedCount $lockedCount -FoundBytes $foundBytes -RemovedBytes $removedBytes -Reason $componentInfo.Note)
        }
    }

    Write-KobraModuleLog -Log $Log -Message 'Browser cleanup complete.'
    return $results
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
            $seen = @{}
            foreach ($item in $componentItems) {
                foreach ($file in @(Get-KobraFilesForBrowserItem -Item $item)) {
                    if ($null -eq $file -or $seen.ContainsKey($file.FullName)) {
                        continue
                    }

                    $seen[$file.FullName] = $true
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
