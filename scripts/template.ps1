<#
.UPDATED   yyyy-MM-dd
.BY        your name
.NOTES     What changed and why. Replace this whole line each edit - one line
           is enough. Longer history lives in Archive.
#>
 
# ============================= TEMPLATE =============================
# The loader IGNORES this file (underscore in the name), so it can sit here
# safely. To create a new command:
#
#   1. Copy this file into Scripts\PS7 (or Scripts\PS51) as <Verb-Noun>.ps1.
#      Use an approved PowerShell verb (Get-Verb lists them).
#   2. Rename the function below to match the filename EXACTLY.
#   3. Write the command. Read environment values (share paths, relay, name
#      patterns) with Get-SysAdminToolsSetting - never hardcode them.
#   4. Update the header block at the top: date, your name, one-line note.
#
# While developing on the share, save work-in-progress as
# <Verb-Noun>_WIP.ps1 in the same folder - the underscore keeps the loader
# away from it until you rename it over the live file.
# ====================================================================
 
function Verb-Noun {
    <#
    .SYNOPSIS
    One line: what this command does.
 
    .DESCRIPTION
    A few sentences for Get-Help: what it touches, what it outputs, anything
    an admin should know before running it.
 
    .EXAMPLE
    Verb-Noun
    What this example does.
    #>
    [CmdletBinding()]
    param(
        # [Parameter(Mandatory = $false)]
        # [string]$ComputerName
    )
 
    try {
 
        # --- operational work goes here ---
 
        Write-SysAdminToolsAudit -Command $MyInvocation.MyCommand.Name -Result Success
    }
    catch {
        Write-SysAdminToolsAudit -Command $MyInvocation.MyCommand.Name -Result Failure -Note $_.Exception.Message
        throw
    }
}
 
# If this command replaces a legacy name, keep the old name working.
# The loader exports aliases automatically:
# Set-Alias -Name Old-LegacyName -Value Verb-Noun
 
