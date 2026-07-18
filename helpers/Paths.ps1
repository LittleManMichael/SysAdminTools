<#
.UPDATED   2026-07-18
.BY        Michael Sprous (Mike)
.NOTES     Path resolution for the standard module folders, either on this
           network's authoritative share (default) or under the locally
           imported module copy.
#>

function Get-SysAdminToolsPath `
{
    <#
    .SYNOPSIS
    Returns the path to a standard SysAdminTools folder. Defaults to the
    authoritative share; use -Local for the locally imported module copy.

    .EXAMPLE
    Get-SysAdminToolsPath -Type Reports -Name 'DriveSpace'
    # \\<share>\SysAdminTools\Reports\DriveSpace

    .EXAMPLE
    Get-SysAdminToolsPath -Type Resources -Local
    # C:\...\Modules\SysAdminTools\<version>\Resources
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [ValidateSet('Scripts', 'Helpers', 'Resources', 'Reports', 'Logs', 'Archive', 'Docs')]
        [string]$Type,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [switch]$Local
    )

    # if local host option selected the root is the root of the psmodule
    if ($Local) 
    {
        $root = Get-SysAdminToolsRoot
    }
    
    # if not local host option, provide the share root (admin share) of the psmodule
    else 
    {
        $root = Get-SysAdminToolsShareRoot
    }

    # Join the path of the root and the path of the selected directory parameter
    $path = Join-Path $root $Type
    
    # if a specific name is selected then gather that instead
    if ($Name) 
    {
        $path = Join-Path $path $Name
    }

    # Output the resulting path location
    return $path
}

function New-SysAdminToolsReportFolder 
{
    <#
    .SYNOPSIS
    Ensures a report-type folder exists under Reports on the share and returns
    its path. Report folders are organized by report type, e.g. DriveSpace,
    ServerStatus, StaleAccounts\Users.

    .EXAMPLE
    $folder = New-SysAdminToolsReportFolder -ReportType 'DriveSpace'
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ReportType
    )

    # Gather the reports path from the function above. 
    $path = Get-SysAdminToolsPath -Type Reports -Name $ReportType

    # If the test path does not work, create it.
    if (-not (Test-Path $path)) 
    {
        New-Item -Path $path -ItemType Directory -Force | Out-Null
    }

    # output the resulting path location
    return $path
}
