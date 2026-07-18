<#
.UPDATED   2026-07-18
.BY        Michael Sprous (mike)
.NOTES     Central audit logging. One CSV row per command run - who, what,
           when, where, result. That is the entire logging system by design:
           detailed troubleshooting information belongs in console output and
           reports, not in log files nobody reads.
#>

function Write-SysAdminToolsAudit {
    <#
    .SYNOPSIS
    Appends one accountability row for a command run to the central audit log
    on the share: Timestamp, Command, User, Computer, Result, Note.

    .DESCRIPTION
    The audit log lives at <ShareRoot>\Logs\Audit_yyyy-MM.csv (one file per
    month so no single file grows forever).

    There is deliberately no local fallback: if the share is unreachable, the
    admin has bigger problems than audit retention. In that case this function
    writes a warning and returns - a logging failure never blocks or kills the
    operational command.

    .EXAMPLE
    # Standard usage inside a command (see Scripts\_Template.ps1):
    try {
        # ... operational work ...
        Write-SysAdminToolsAudit -Command $MyInvocation.MyCommand.Name -Result Success
    }
    catch {
        Write-SysAdminToolsAudit -Command $MyInvocation.MyCommand.Name -Result Failure -Note $_.Exception.Message
        throw
    }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Command,

        [Parameter(Mandatory)]
        [ValidateSet('Success', 'Failure')]
        [string]$Result,

        [Parameter(Mandatory = $false)]
        [string]$Note = ''
    )

    try 
    {
        # gather the logging directory
        $logDir = Get-SysAdminToolsPath -Type Logs

        # if one doesn't exist, create one
        if (-not (Test-Path $logDir)) 
        {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }

        # Join the logging directory with the audit CSV file
        $logFile = Join-Path $logDir ("Audit_{0}.csv" -f (Get-Date -Format 'yyyy-MM'))

        # Create the audit log object
        $row = [pscustomobject] `
        @{
            Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            Command   = $Command
            User      = $env:USERNAME
            Computer  = $env:COMPUTERNAME
            Result    = $Result
            Note      = $Note
        }

        # Export the audit log to the CSV file.
        $row | Export-Csv -Path $logFile -Append -NoTypeInformation -Encoding UTF8
    }
    
    # If this command fails, catch with this warning.
    catch 
    {
        Write-Warning ("[SysAdminTools] Could not write to the central audit log: {0}" -f $_.Exception.Message)
    }
}
