SysAdminTools Command Reference
One row per operational command. The Legacy name column is the alias kept for muscle memory — both names work at the console, and searching the share for either name should lead you to the file listed here. Add a row every time a command is ported or created.
Last updated: 2026-07-17
Operational commands
Command	Legacy name (alias)	PS7	PS51	File	What it does
Get-SysAdminToolsInfo	—	—	✔	Scripts\PS51\Get-SysAdminToolsInfo.ps1	Shows how the module loaded this session: engine mode, which network's settings are active, and which commands came from PS7, PS51, or require PowerShell 7. Run it after import to confirm the module is healthy.
Maintainer helpers
These load from Helpers\ and are available at the console, but exist mainly for scripts to use.
Function	File	What it does
Get-SysAdminToolsSetting	Helpers\Config.ps1	Returns one value from this network's Resources\Settings.psd1; dot paths reach nested values (ServerNamePatterns.Exchange).
Get-SysAdminToolsSettings	Helpers\Config.ps1	Returns the whole settings table.
Get-SysAdminToolsShareRoot	Helpers\Config.ps1	This network's authoritative SysAdminTools share root, from settings.
Get-SysAdminToolsRoot	Helpers\Config.ps1	Root folder of the locally imported module copy.
Get-SysAdminToolsPath	Helpers\Paths.ps1	Path to a standard module folder (Reports, Logs, Resources, ...) on the share, or locally with -Local.
New-SysAdminToolsReportFolder	Helpers\Paths.ps1	Ensures Reports<type> exists on the share and returns its path.
Write-SysAdminToolsAudit	Helpers\Logging.ps1	Appends the one-line-per-run accountability row to the central audit log.
