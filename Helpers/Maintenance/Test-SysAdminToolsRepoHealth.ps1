
function Test-SysAdminToolsRepoHealth {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$SourceShareModuleRoot = (Get-SysAdminToolsShareRoot), #'\\newton\admin\SysAdmin Powershell Module\SysAdminTools\',

        [Parameter(Mandatory=$false)]
        [switch]$ProblemsOnly
    )

    $ErrorActionPreference = 'Stop'

    $scriptsRoot = Join-Path $SourceShareModuleRoot 'Scripts'
    if (-not (Test-Path $scriptsRoot)) {
        throw "Scripts folder not found: $scriptsRoot"
    }

    $scriptDirs = Get-ChildItem -Path $scriptsRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne '_ARCHIVE' }
    if (-not $scriptDirs) {
        Write-Warning "No script folders found under: $scriptsRoot"
        return @()
    }

    $results = foreach ($dir in $scriptDirs) {
        $scriptName   = $dir.Name 
        $scriptFolder = $dir | Select-Object -ExpandProperty FullName
        $currentPath  = Join-Path $scriptFolder ($scriptName + '.ps1')
        $archivePath  = Join-Path $scriptFolder 'Archive'
        $dataPath     = Join-Path $scriptFolder 'Data'

        $issues = New-Object System.Collections.Generic.List[String]

        # Check current script exists
        if (-not (Test-Path $currentPath)) {
            $issues.Add("MissingCurrentPS1")
        }

        # Check Archive Folder Exists
        if (-not (Test-Path $archivePath)) {
            $issues.Add("MissingArchiveFolder")
        }

        # Extra ps1 files showing in the script folder root 
        $ps1InRoot = Get-ChildItem -Path $scriptFolder -Filter '*.ps1' -File -ErrorAction SilentlyContinue
        $extraPs1 = @()
        if ($ps1InRoot) {
            $extraPs1 = $ps1InRoot | Where-Object { $_.Name -ne ($scriptName + '.ps1') }
            if ($extraPs1.Count -gt 0) {
                $issues.Add("ExtraPS1InRoot: " + (($extraPs1 | Select-Object -ExpandProperty Name) -join ', '))
            }
        }

        # Any copy-like files present
        $copyLike = @()
        if ($ps1InRoot) {
            $copyLike = $ps1InRoot | Where-Object {
                $_.Name -ne ($scriptName + '.ps1') -and ($_.Name -match '(?i)copy')
            }
            if ($copyLike.count -gt 0) {
                $issues.Add("CopyFilesPresent: " + (($copyLike | Select-Object -ExpandProperty Name) -join ', '))
            }
        }

        # Validate header, version, and function name. (only if current exists)
        $version = ''
        $hasHeader = $false
        $hasVersion = $false
        $functionMatch = $false

        if (Test-Path $currentPath) {
            $content = Get-Content -Path $currentPath -Raw -ErrorAction Stop 

            $hasHeader = ($content -match '(?s)<#.*?#>')
            if (-not $hasHeader) { $issues.Add("MissingHeaderBlock") }

            if ($content -match '(?im)^\s*\.VERSION\s+([0-99]+(\.[0-99]+){1,2})\s*$') {
                $hasVersion = $true
                $version = $Matches[1]
            } else {
                $issues.Add("MissingVersionField")
            }

            # Function Name needs to match file/folder name
            $functionMatch = ($content -match "(?im)^\s*function\s+$([regex]::Escape($scriptName))\b")
            if (-not $functionMatch) { $issues.Add("FunctionNameMismatch") }
        }

        # Archive naming sanity
        $archiveCount = 0
        $badArchiveNames = @()
        if (Test-Path $archivePath) {
            $archived = Get-ChildItem -Path $archivePath -Filter '*.ps1' -File -ErrorAction SilentlyContinue
            if ($archived) {
                $archiveCount = $archived.count
                # expected pattern: YYY-MM-DD_vX.Y_ScriptName.ps1
                $badArchiveNames = $archived | Where-Object {
                    $_.Name -notmatch '^\d{4}-\d{2}-\d{2}_v\d+(\.\d+){1,2}_.+\.ps1$'
                }
                if ($badArchiveNames.count -gt 0) {
                    $issues.Add("BadArchiveNames: " + (($badArchiveNames | Select-Object -ExpandProperty Name) -join ', '))
                }
            }
        }

        [pscustomobject]@{
            ScriptName        = $scriptName
            CurrentPS1Exists  = (Test-Path $currentPath)
            Version           = $version
            ArchiveExists     = (Test-Path $archivePath)
            ArchiveCount      = $archiveCount
            DataFolderExists  = (Test-Path $dataPath)
            IssuesCount       = $issues.count
            Issues            = ($issues -join '; ')
            CurrentPath       = $currentPath
        }
    }

    if ($ProblemsOnly) {
        $results = $results | Where-Object { $_.IssuesCount -gt 0 }
    }

    # Display-Friendly Output
    $results | 
        Sort-Object IssuesCount, ScriptName |
        Format-Table ScriptName, Version, ArchiveCount, IssuesCount, Issues -AutoSize

    return $results
}
