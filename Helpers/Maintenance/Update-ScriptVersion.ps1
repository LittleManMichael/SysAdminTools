# Script version bumper

function Update-ScriptVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ScriptName,

        [Parameter(Mandatory)]
        [ValidatePattern('^\d+(\.\d+){1,2}$')] # 1.2 or 1.2.3
        [string]$NewVersion,

        [Parameter(Mandatory=$false)]
        [string]$Note = '',

        [Parameter(Mandatory=$false)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string]$NotePath,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$SourceShareRoot = '\\newton\admin\SysAdmin Powershell Module\SysAdminTools',

        [Parameter(Mandatory=$false)]
        [switch]$Force
    )

    $ErrorActionPreference = 'Stop'

    # --- Resolve notes type (inline or file)
    $resolvedNote = ''
    if ($NotePath) { 
        $resolvedNote = Get-content -Path $NotePath -Raw
    } else {
        $resolvedNote = $Note
    }

    # If no note provided. Throw an error stating to provide a note or use -Force
    if (-not $resolvedNote -and -not $Force) {
        throw "No note provided. Provide a Note via -Note or -NotePath, or use -Force to proceed without a note."
    }

    $scriptDir = Join-Path $SourceShareRoot "Scripts\$ScriptName"
    $scriptPath = Join-Path $scriptDir "$ScriptName.ps1"
    $archiveDir = Join-Path $scriptDir "Archive"

    if (-not (Test-Path $scriptPath)) {
        throw "Script not found on share: $scriptPath"
    }

    if (-not (Test-Path $archiveDir)) {
        New-Item -Path $archiveDir -ItemType Directory -Force | Out-Null
    }

    # Dectect common "copy" files sitting next to the real one
    $copyLike = Get-ChildItem -Path $scriptDir -Filter '*.ps1' -File -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -ne "$ScriptName.ps1" -and 
            ($_.Name -match '(?i)copy' -or $_.BaseName -match '(?i)copy')
        }

    if ($copyLike -and -not $Force) {
        $names = ($copyLike | Select-Object -ExpandProperty Name) -join ', '
        throw "Found copy-like files in $($scriptDir): $names. Cleam them up or use -Force."
    }

    # Read current raw file data
    $content = Get-content -Path $scriptPath -Raw

    # ensure function name exists in the file
    if ($content -notmatch "(?im)^\s*function\s+$([regex]::Escape($ScriptName))\b") {
        if (-not $Force) {
            throw "Function '$ScriptName' not found in $ScriptName.ps1. Fix the function name or use -Force."
        }
    }

    # Dectect current script version (default 0.0 if missing)
    $currentVersion = '0.0'
    if ($content -match '(?im)^\s.VERSION\s+([0-99]+(\.[0-99]+){1,2})\s*$') {
        $currentVersion = $Matches[1]
    }

    $today = Get-Date -Format 'yyyy-MM-dd'

    # Archive current script before modifying.
    $archiveName = "{0}_v{1}_{2}.ps1" -f $today, $currentVersion, $ScriptName
    Copy-Item -Path $scriptPath -Destination (Join-Path $archiveDir $archiveName) -Force

    # Build notes block (Safe for multi-line)
    $noteBlock = ''
    if ($resolvedNote) {
        $noteLines = ($resolvedNote -split "`r?`n")
        $noteBlock = ($noteLines | ForEach-Object { ".NOTES    $_" }) -join "`r`n"
    }

    # Create the header if missing; otherwise update
    if ($content -notmatch '(?s)<#.*?#>') {
        $header = @"
<# 
.SCRIPTNAME $ScriptName
.VERSION $NewVersion
.LASTEDIT $today
$noteBlock
#>

"@ 
        $content = $header + $content 
    }
    else {
        $content = [regex]::Replace($content, '(?im)^\s*\.VERSION\s+.*$', ".VERSION    $NewVersion")
        $content = [regex]::Replace($content, '(?im)^\s*\.LASTEDIT\s+.*$', ".LASTEDIT    $today")

        if ($resolvedNote) {
            # Remove existing NOTES lines and inject fresh ones
            $content = [regex]::Replace($content, '(?im)^\s*\.NOTES\s+.*\r?\n?' , '')
            $content = [regex]::Replace($content, '(?s)(<#.*?\r?\n)', "`$1$noteBlock `r`n")
        }
    }

    # Write back to share
    Set-Content -Path $scriptPath -Value $content -Encoding UTF8

    Write-Host "[SysAdmin] Archived $ScriptName v$currentVersion -> $archiveName" -ForegroundColor Green
    Write-Host "[SysAdmin] Updated $ScriptName on share to v$NewVersion" -ForegroundColor Green
}

