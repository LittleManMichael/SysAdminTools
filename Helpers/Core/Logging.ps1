# Helpers\Core\Logging.ps1

function Write-SysAdminToolsLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ScriptName,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO',

        [Parameter(Mandatory=$false)]
        [switch]$UseLocal
    )

    $ErrorActionPreference = 'Stop'

    New-SysAdminToolsOutputFolders -UseLocal:$UseLocal

    $logDir = Get-SysAdminToolsPath -Type Logs -UseLocal:$UseLocal
    $date   = Get-Date -Format 'yyyyMMdd'
    $time   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $user   = $env:USERNAME
    $hostn  = $env:COMPUTERNAME

    # One file. Per Script. Per Day.
    $logFile = Join-Path $logDir ("{0}_{1}.log" -f $ScriptName, $date)
    $line = "{0} [{1}] ({2}@{3}) {4}" -f $time, $Level, $user, $hostn, $Message
    Add-Content -Path $logFile -Value $line -Encoding UTF8
}

function Write-SysAdminToolsErrorFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ScriptName,

        [Parameter(Mandatory)]
        [System.Exception]$Exception,

        [Parameter(Mandatory=$false)]
        [switch]$UseLocal
    )

    $ErrorActionPreference = 'Stop'

    New-SysAdminToolsOutputFolders -UseLocal:$UseLocal

    $errDir = Get-SysAdminToolsPath -Type Errors -UseLocal:$UseLocal
    $stamp  = Get-Date -Format 'yyyMMdd_HHmmss'
    $file   = Join-Path ("{0}_{1}.err.txt" -f $ScriptName, $stamp)

    $body = @(
        "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "User: $($env:USERNAME)"
        "Machine: $($env:COMPUTERNAME)"
        "Message: $($Exception.Message)"
        ""
        "StackTrace:"
        $Exception.StackTrace
    ) -join "`r`n"

    Set-Content -Path $file -Value $body -Encoding UTF8
}

function Write-SysAdminToolsReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ScriptName,

        [Parameter(Mandatory)]
        [string]$FileName,

        [Parameter(Mandatory)]
        [string]$Content,

        [switch]$UseLocal
    )
    
    $ErrorActionPreference = 'Stop'

    New-SysAdminToolsOutputFolders -UseLocal:$UseLocal

    $reportsDir = Get-SysAdminToolsPath -Type Reports -UseLocal:$UseLocal
    $fullpath = Join-Path $reportsDir ("{0}_{1}" -f $ScriptName, $FileName)

    Write-FileSafe -Path $fullpath -Content $Content

    return $fullpath
}