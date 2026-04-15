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

function Invoke-KobraHPDebloat {
    [CmdletBinding()]
    param([scriptblock]$Log)

    Write-KobraModuleLog -Log $Log -Message 'Starting HP telemetry trim...'
    Write-KobraModuleLog -Log $Log -Message 'Only analytics / support-assistant style services are targeted.'

    $serviceNames = @(
        'HPAppHelperCap',
        'HPDiagsCap',
        'HPNetworkCap',
        'HPSysInfoCap',
        'HP TechPulse Core',
        'HP Insights Analytics Service',
        'HPAnalyticsService'
    )

    foreach ($serviceName in $serviceNames) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($null -eq $service) {
            continue
        }

        try {
            if ($service.Status -ne 'Stopped') {
                Stop-Service -Name $serviceName -Force -ErrorAction Stop
                Write-KobraModuleLog -Log $Log -Message ("  Stopped service: {0}" -f $serviceName)
            }
        }
        catch {
            Write-KobraModuleLog -Log $Log -Message ("  Could not stop service: {0}" -f $serviceName)
        }

        try {
            Set-Service -Name $serviceName -StartupType Disabled -ErrorAction Stop
            Write-KobraModuleLog -Log $Log -Message ("  Disabled service: {0}" -f $serviceName)
        }
        catch {
            Write-KobraModuleLog -Log $Log -Message ("  Could not disable service: {0}" -f $serviceName)
        }
    }

    if (Get-Command -Name Get-ScheduledTask -ErrorAction SilentlyContinue) {
        $taskPatterns = @(
            '*HP Support Assistant*',
            '*HP Analytics*',
            '*TechPulse*',
            '*HP Insights*'
        )

        $tasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
            $name = $_.TaskName
            $path = $_.TaskPath
            ($taskPatterns | Where-Object { $name -like $_ -or $path -like $_ }).Count -gt 0
        }

        foreach ($task in $tasks) {
            try {
                Disable-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction Stop | Out-Null
                Write-KobraModuleLog -Log $Log -Message ("  Disabled scheduled task: {0}{1}" -f $task.TaskPath, $task.TaskName)
            }
            catch {
                Write-KobraModuleLog -Log $Log -Message ("  Could not disable task: {0}{1}" -f $task.TaskPath, $task.TaskName)
            }
        }
    }
    else {
        Write-KobraModuleLog -Log $Log -Message 'ScheduledTask cmdlets not available on this system.'
    }

    Write-KobraModuleLog -Log $Log -Message 'HP telemetry trim complete.'
}

Export-ModuleMember -Function Invoke-KobraHPDebloat
