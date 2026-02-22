# Helpers\Core\Common.ps1

function Write-FileSafe {
    <# Salfe writes a file by ensuring parent folders exist and replacing existing files cleanrly #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Content
    )

    $parent = Split-Path -Parent $Path

    if (-not (Test-Path $parent)) {
        New-Item -Path $parent -ItemType Directory -Force | Out-Null
    }

    if (Test-Path $Path) {
        Remove-Item -Path $Path -Force -ErrorAction SilentlyContinue
    }

    Set-Content -Path $Path -Value $Content -Encoding UTF8
}

