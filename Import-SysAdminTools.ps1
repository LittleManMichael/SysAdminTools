# Import-SysAdmin.ps1 
# Copies SysAdmin module from a central share to the local module path, then imports it. 

param(
    [Parameter(Mandatory=$false)]
    [string]$SourceModuleRoot = '\\newton\admin\SysAdmin Powershell Module\SysAdminTools',

    [Parameter(Mandatory=$false)]
    [string]$ModuleName = 'SysAdminTools'
)

$ErrorActionPreference = 'Stop'

$psd1Path = Join-Path $SourceModuleRoot "$ModuleName.psd1"
if (-not (Test-Path $psd1Path)) {
    throw "Module manifest not found: $psd1Path"
}

$manifest = Import-PowerShellDataFile -Path $psd1Path
$version = [string]$manifest.ModuleVersion

$targetRoot = Join-Path $env:ProgramFiles "WindowsPowerShell\Modules\$ModuleName\$version"
$targetRootModule = Join-Path $env:ProgramFiles "WindowsPowerShell\Modules\$ModuleName"

# Install (copy) if missing
if (Test-Path $targetRoot) {
    Write-Warning "[SysAdmin] $ModuleName already exists at $targetRoot"
    $choice = Read-Host "Overwrite? (Y/N)"

    if ($choice -match '^(Y|y)$') {
        Remove-Item -Path $targetRoot -Recurse -Force -ErrorAction Stop
        New-Item -Path $targetRoot -ItemType Directory -Force | Out-Null
    }
    else { 
        Write-Host "[SysAdmin] Skipped overwrite. Importing existing module..." -f Yellow
        Import-Module $ModuleName -Force -ErrorAction Stop
        return
    }
}
else {
    New-Item -Path $targetRoot -ItemType Directory -Force | Out-Null
}

# Copy fresh
Copy-Item -Path (Join-Path $SourceModuleRoot '*') -Destination $targetRoot -Recurse -Force -ErrorAction Stop

# Import
Import-Module $ModuleName -Force -ErrorAction Stop
Write-Host "[SysAdmin] Imported $ModuleName ($version) from $targetRoot" -f Green
