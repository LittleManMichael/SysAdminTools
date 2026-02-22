# Helpers\Core\Config.ps1
# This is the primary config for the SysAdminTools repo root.

function Get-SysAdminToolsShareRoot {
    [CmdletBinding()]
    param()

    # Set the PARENT directory of the module root
    '\\newton\admin\SysAdmin Powershell Module\SysAdminTools'
}

function Get-SysAdminComputerList {
    [CmdletBinding()]
    param()

    # Set the filepath for the list of computers
    '\\newton\admin\SysAdmin Powershell Module\SysAdminTools\Data\SysAdmin_Computers.txt'
}