# Update-SysAdmin.ps1
# Push SysAdmin module to remote machines safely
function Update-SysAdminTools {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$SourceModuleRoot = '\\newton\admin\SysAdmin Powershell Module\SysAdminTools',

        [Parameter(Mandatory=$false)]
        [string]$ModuleName = 'SysAdminTools',

        [Parameter(Mandatory=$false)]
        [string[]]$ComputerName,

        [Parameter(Mandatory=$false)]
        [string]$ComputerListPath = (Get-SysAdminComputerList), #'\\newton\admin\SysAdmin Powershell Module\Data\SysAdmin_Computers.txt',

        [Parameter(Mandatory=$false)]
        [switch]$AddToList,
    
        [Parameter(Mandatory=$false)]
        [switch]$Force,
    
        [Parameter(Mandatory=$false)]
        [switch]$PruneOldVersions
    )

    $ErrorActionPreference = 'Stop'

    function Get-ModuleVersionFromPsd1 {
        param([string]$Psd1Path)
        $m = Import-PowerShellDataFile -Path $Psd1Path
        return [version]$m.ModuleVersion
    }

    function Get-HighestVersionFolder {
     param([string]$Path)
     if (-not (Test-Path $Path)) { return $null }

     $dirs = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue
     if (-not $dirs) { return $null }

      $versions = $dirs.Name | ForEach-Object {
          try { [version]$_ } catch { $null }
       } | Where-Object { $_ -ne $null }

       if (-not $versions) { return $null }
        return ($versions | Sort-Object -Descending | Select-Object -First 1)
    }

    # --- Validate Source
    $sourcePsd1 = Join-Path $SourceModuleRoot "$ModuleName.psd1"
    if (-not (Test-Path $sourcePsd1)) { throw "Source Manifest not found: $sourcePsd1" }
    $sourceVersion = Get-ModuleVersionFromPsd1 -Psd1Path $sourcePsd1
    
    # Build Target List
    $targets = @()
    
    if ($ComputerName -and $ComputerName.count -gt 0) {
        $targets += $ComputerName
    
        if ($AddToList) {
            if (-not (Test-Path $ComputerListPath)) {
                New-Item -Path $ComputerListPath -ItemType File -Force | Out-Null
            }
    
            $existing = Get-content $ComputerListPath -ErrorAction SilentlyContinue | 
                ForEach-Object { $_.Trim() } | Where-Object { $_ }
    
            $merged = ($existing + $ComputerName) | 
                ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique
    
            Set-Content -PassThru $ComputerListPath -Value $merged
        }
    }
    
    else {
        if (-not (Test-Path $ComputerListPath)) {
            throw "No ComputerName provided and list file not found: $ComputerListPath"
        }
        $targets = Get-Content $ComputerListPath | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }
    
    if (-not $targets -or $targets.count -eq 0) { throw "No target computers specified." }
    
    # --- Deploy ---
    $results = foreach ($c in $targets) {
        
        $status = [ordered]@{ 
            Computer      = $c 
            Reachable     = $false
            Action        = ''
            RemoteVersion = ''
            SourceVersion = $sourceVersion.ToString()
            Message       = ''
        }
    
        try { 
            if (-not (Test-Connection -ComputerName $c -count 1 -Quiet)) {
                $status.Action = 'Failed'
                $status.Message = 'Ping Failed'
                [pscustomobject]$status
                continue
            }
            $status.Reachable = $true
    
            $remotebase = "\\$c\c$\ProgramFiles\WindowsPowerShell\Modules\$ModuleName"
            $remoteHighest = Get-HighestVersionFolder -Path $remotebase
            if ($remoteHighest) { $status.RemoteVersion = $remoteHighest.ToString() }
    
            $needsUpdate = $Force -or (-not $remoteHighest) -or ($remoteHighest -lt $sourceVersion)
    
            if (-not $needsUpdate) {
                $status.Action = 'Skipped'
                $status.Message = "Alright up-to-date (remote highest: $remoteHighest)"
                [pscustomobject]$status
                continue
            }
    
            # Ensure module base exists
            if (-not (Test-Path $remotebase)) {
                New-Item -Path $remotebase -ItemType Directory -Force | Out-Null
            }
    
            # Destination version folder
            $remoteVersionPath = Join-Path $remotebase $sourceVersion.ToString()
    
            # Remove version folder if it exists
            if (Test-Path $remoteVersionPath) {
                Remove-Item -Path $remoteVersionPath -Recurse -Force -ErrorAction Stop
                Start-Sleep -Milliseconds 200
            }
    
            # Recreate folder
            New-Item -Path $remoteVersionPath -ItemType Directory -Force | Out-Null
    
            # Copy module contents
            Copy-Item -Path (Join-Path $SourceModuleRoot '*') -Destination $remoteVersionPath -Recurse -Force -ErrorAction SilentlyContinue
    
            # prune old versions (keeps newest only)
            if ($PruneOldVersions) {
                $all = Get-ChildItem -Path $remotebase -Directory -ErrorAction SilentlyContinue
                if ($all) {
                    $versionDirs = foreach ($d in $all) {
                        try {
                            [pscustomobject]@{ Dir = $d.FullName; Ver = [version]$d.Name }
                        } catch { }
                    }

                    if ($versionDirs) {
                        $keeps = ($versionDirs | Sort-Object Ver -Descending | Select-Object -First 1).Dir
                        foreach ($vd in $versionDirs) {
                            if ($vd.Dir -ne $keep) {
                                try { Remove-Item -Path $vd.Dir -Recurse -Force -ErrorAction Stop } catch { }
                            }
                        }
                    }
                }
            }

            $status.Action = if ($remoteHighest) { 'Updated' } else { 'Installed' }
            $status.Message = "Deployed $ModuleName $sourceVersion to $remoteVersionPath"
            [pscustomobject]$Status
        }

        catch {
            $Status.Action = 'Failed'
            $status.Message = $_.Exception.Message
            [pscustomobject]$status
        }
    }

    $results | Sort-Object Computer | Format-Table -AutoSize
}