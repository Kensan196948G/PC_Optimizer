Set-StrictMode -Version Latest

function Invoke-CleanupMaintenance {
    [CmdletBinding()]
    param(
        [switch]$WhatIfMode,
        [string[]]$Tasks = @('temp','browser','store')
    )

    $supported = @('temp','browser','store','dns')
    $normalizedTasks = @($Tasks | ForEach-Object { "$_".ToLowerInvariant() } | Where-Object { $_ -in $supported } | Select-Object -Unique)
    $executed = @()

    function Clear-ChildrenSafe {
        param(
            [Parameter(Mandatory)][string]$RootPath,
            [string[]]$ExcludeNamePatterns = @()
        )
        if (-not (Test-Path -LiteralPath $RootPath)) { return }
        $items = @(Get-ChildItem -LiteralPath $RootPath -Force -ErrorAction SilentlyContinue)
        foreach ($item in $items) {
            $skip = $false
            foreach ($pat in $ExcludeNamePatterns) {
                if ($item.Name -like $pat) { $skip = $true; break }
            }
            if ($skip) { continue }
            Remove-Item -LiteralPath $item.FullName -Force -Recurse -ErrorAction SilentlyContinue
        }
    }

    foreach ($task in $normalizedTasks) {
        if ($WhatIfMode) {
            $executed += [PSCustomObject]@{ Task = $task; Status = 'PreviewOnly' }
            continue
        }

        switch ($task) {
            'temp' {
                # Keep PowerShell temporary module-proxy folders to avoid runtime command-loading failures.
                Clear-ChildrenSafe -RootPath $env:TEMP -ExcludeNamePatterns @('remoteIpMoProxy_*')
                $executed += [PSCustomObject]@{ Task = $task; Status = 'Completed' }
            }
            'browser' {
                Clear-ChildrenSafe -RootPath (Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data\Default\Cache')
                Clear-ChildrenSafe -RootPath (Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\User Data\Default\Cache')
                $ffProfiles = Join-Path $env:APPDATA 'Mozilla\Firefox\Profiles'
                if (Test-Path -LiteralPath $ffProfiles) {
                    @(Get-ChildItem -LiteralPath $ffProfiles -Directory -ErrorAction SilentlyContinue) | ForEach-Object {
                        $entries = Join-Path $_.FullName 'cache2\entries'
                        Clear-ChildrenSafe -RootPath $entries
                    }
                }
                $executed += [PSCustomObject]@{ Task = $task; Status = 'Completed' }
            }
            'store' {
                $storeCache = "$env:LOCALAPPDATA\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalCache"
                Clear-ChildrenSafe -RootPath $storeCache
                $executed += [PSCustomObject]@{ Task = $task; Status = 'Completed' }
            }
            'dns' {
                & ipconfig /flushdns | Out-Null
                $executed += [PSCustomObject]@{ Task = $task; Status = 'Completed' }
            }
        }
    }

    [PSCustomObject]@{
        Component = 'Cleanup'
        WhatIf    = [bool]$WhatIfMode
        Tasks     = $normalizedTasks
        Status    = 'OK'
        Results   = @($executed)
    }
}

Export-ModuleMember -Function Invoke-CleanupMaintenance
