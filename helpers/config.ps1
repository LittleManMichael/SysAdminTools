<#
.UPDATED   2026-07-17
.BY        Michael Sprous (Mike)
.NOTES     Added detailed annotations regarding process of functions

#>

function Get-SysAdminToolsRoot 
{
    <#
    .SYNOPSIS
    Returns the root folder of the locally imported SysAdminTools module.
    #>
    [CmdletBinding()]
    param()

    # Output the result
    return $script:SysAdminToolsRoot
}

function Get-SysAdminToolsSettings 
{
    <#
    .SYNOPSIS
    Returns the entire settings table loaded from Resources\Settings.psd1.
    #>
    [CmdletBinding()]
    param()

    # Output the result
    return $script:SysAdminToolsSettings
}

function Get-SysAdminToolsSetting 
{
    <#
    .SYNOPSIS
    Returns one value from Resources\Settings.psd1. Supports dot paths for
    nested values.

    .EXAMPLE
    Get-SysAdminToolsSetting -Name 'ShareRoot'

    .EXAMPLE
    Get-SysAdminToolsSetting -Name 'DistributionLists.SysAdmins'

    .EXAMPLE
    Get-SysAdminToolsSetting -Name 'ServerNamePatterns.Exchange'
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    # Set the value variable as the current settings.psd1 content.
    $value = $script:SysAdminToolsSettings

    # Go though each part in the name and property (requested)
    foreach ($part in ($Name -split '\.')) 
    {
        # Check if the value AND a property were requested.
        if ($value -is [System.Collections.IDictionary] -and $value.Contains($part)) 
        {
            # if so, set the value as the property of the selected value
            $value = $value[$part]
        }
        
        # if it checked and couldn't find it, throw and error stating it can't and give recommendation to add it.
        else 
        {
            throw "Setting '$Name' was not found in Resources\Settings.psd1 on this network. Add it to the settings file (see Resources\Settings.psd1 for the expected keys)."
        }
    }

    # output the value
    return $value
}

function Get-SysAdminToolsShareRoot 
{
    <#
    .SYNOPSIS
    Returns this network's authoritative SysAdminTools share root, from settings.
    #>
    [CmdletBinding()]
    param()

    # Output the result
    return (Get-SysAdminToolsSetting -Name 'ShareRoot')
}
