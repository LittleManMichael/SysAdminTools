# SysAdminTools.psm1
# Hybrid module loader for Core, Maintenance, and Script functions

Set-StrictMode -Version Latest
# Set a module-scoped variable that all functions can access.
$script:SysAdminToolsRoot = $PSScriptRoot 
$ModuleRoot = $PSScriptRoot

# --- Master list of all functions we intend to export ---
$FunctionsToExport = New-Object System.Collections.Generic.List[string]

# ----------------------------------------------------------------
# 1. Process Core Helpers (Multiple functions per file)
# ----------------------------------------------------------------
$coreHelpersPath = Join-Path $ModuleRoot 'Helpers\Core'
if (Test-Path $coreHelpersPath) {
    $coreHelperFiles = Get-ChildItem -Path $coreHelpersPath -Filter '*.ps1' -File

    foreach ($file in $coreHelperFiles) {
        # First, load the file's contents into memory
        try {
            . $file.FullName
        }
        catch {
            Write-Warning "[SysAdminTools] Failed loading $($file.FullName): $($_.Exception.Message)"
            continue # Skip to the next file if loading fails
        }

        # Now, parse the file to find the names of all functions within it
        try {
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)
            $functionsInFile = $ast.FindAll( { $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)

            if ($functionsInFile) {
                # Add each discovered function name to our master list
                foreach ($functionAst in $functionsInFile) {
                    $FunctionsToExport.Add($functionAst.Name) | Out-Null
                }
            }
        }
        catch {
             Write-Warning "[SysAdminTools] Failed to parse $($file.FullName) for functions: $($_.Exception.Message)"
        }
    }
}

# ----------------------------------------------------------------
# 2. Process Maintenance Helpers (Filename matches function name)
# ----------------------------------------------------------------
$maintenanceHelpersPath = Join-Path $ModuleRoot 'Helpers\Maintenance'
if (Test-Path $maintenanceHelpersPath) {
    $maintenanceHelperFiles = Get-ChildItem -Path $maintenanceHelpersPath -Filter '*.ps1' -File

    foreach ($file in $maintenanceHelperFiles) {
        # The function name is the same as the filename (without extension)
        $functionName = $file.BaseName
        $FunctionsToExport.Add($functionName) | Out-Null

        # Load the function's code
        try {
            . $file.FullName
        }
        catch {
            Write-Warning "[SysAdminTools] Failed loading $($file.FullName): $($_.Exception.Message)"
        }
    }
}

# ----------------------------------------------------------------
# 3. Process Main Operational Scripts
# ----------------------------------------------------------------
$scriptsRoot = Join-Path $ModuleRoot 'Scripts'
if (Test-Path $scriptsRoot) {
    $scriptDirs = Get-ChildItem -Path $scriptsRoot -Directory |
        Where-Object { $_.Name -notlike '_*' }

    foreach ($dir in $scriptDirs) {
        $scriptName = $dir.Name
        $scriptFile = Join-Path $dir.FullName ($scriptName + '.ps1')

        if (Test-Path $scriptFile) {
            $FunctionsToExport.Add($scriptName) | Out-Null
            
            # Load the function's code
            try {
                . $scriptFile
            }
            catch {
                Write-Warning "[SysAdminTools] Failed loading script $scriptName : $($_.Exception.Message)"
            }
        }
        else {
            Write-Warning "[SysAdminTools] Missing expected script file: $scriptFile"
        }
    }
}

# ----------------------------------------------------------------
# 4. Final Export
# ----------------------------------------------------------------
# Get a unique list of all functions we've collected
$FinalExportList = $FunctionsToExport | Get-Unique

# A final safety check to ensure we only try to export functions that actually exist
# This protects against typos or functions that failed to load
$VerifiedExportList = foreach ($funcName in $FinalExportList) {
    if (Get-Command $funcName -CommandType Function -ErrorAction SilentlyContinue) {
        $funcName # Output the name if the function exists
    }
}

# Using Write-Host as requested to ensure visibility
Write-Host "[SysAdminTools] Exporting the following functions: $($VerifiedExportList -join ', ')"

Export-ModuleMember -Function $VerifiedExportList
