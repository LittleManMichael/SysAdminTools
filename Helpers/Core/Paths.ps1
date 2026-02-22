# Helpers\Core\Paths.ps1

function Get-SysAdminToolsModuleRoot {
    [CmdletBinding()]
    param()

    $mod = Get-Module -Name 'SysAdminTools'
    if (-not $mod) {
        throw "SysAdminTools is not imported. Run: Import-Module SysAdminTools"
    }
    return $mod.ModuleBase
}

function Assert-SysAdminToolsShareAvailable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$ShareRoot = (Get-SysAdminToolsShareRoot)
    )

    if (-not (Test-Path $ShareRoot)) {
        throw "SysAdminTools share root is not reachable: $ShareRoot"
    }
}

function Get-SysAdminToolsPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Logs','Reports','Errors','Scripts','Helpers')]
        [string]$Type,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory=$false)]
        [switch]$UseLocal
    )

    if (-not ($UseLocal)) {
        $root = Get-SysAdminToolsShareRoot
        Assert-SysAdminToolsShareAvailable -ShareRoot $root
    }
    else {
        $root = Get-SysAdminToolsModuleRoot
    }

    $base = Join-Path $root $Type 
    if ($Name) { 
        return Join-Path $base $name 
    }
    return $base
}

function New-SysAdminToolsOutputFolders {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [switch]$UseLocal
    )

    foreach ($t in @('Logs','Reports','Errors')) {
        $p = Get-SysAdminToolsPath -Type $t -UseLocal:$UseLocal 
        if (-not (Test-Path $p)) {
            New-Item -Path $p -ItemType Directory -Force | Out-Null
        }
    }
}
