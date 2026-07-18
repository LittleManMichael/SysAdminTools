SysAdminTools
Internal SysAdmin PowerShell module — PowerShell 7 first, with automatic Windows PowerShell 5.1 compatibility. One file per command, fault-isolated loading, per-network settings, central one-line audit logging.
About this repo: this is a working/collaboration space only. The authoritative copies of SysAdminTools live on the internal networks' admin shares — everything here gets carried on-domain and built there. Nothing environment-specific belongs in this repo: no real share paths, hostnames, addresses, or domain names. The per-network Resources\Settings.psd1 is never committed (only Settings.sample.psd1 is; .gitignore enforces it).
Layout
SysAdminTools.psd1 / .psm1    Manifest and loader
Scripts/PS7, Scripts/PS51     One file per command; filename = function name
Scripts/_Template.ps1         Starting point for new commands
Helpers/                      Config, Paths, Logging
Resources/                    Settings.sample.psd1, templates, lists
Reports/  Logs/  Archive/     Operational output, audit CSVs, prior versions
Docs/                         Architecture, Command Reference, leadership brief
Start with Docs\Architecture.md — it explains the whole design, including the loader's version-selection rules, the underscore rule for _WIP.ps1 files, and the change workflow.
Standing it up on a network
Copy this tree to the network's authoritative admin share.
In Resources, copy Settings.sample.psd1 to Settings.psd1 and fill in that network's real values.
Copy the tree from the share into %ProgramFiles%\WindowsPowerShell\Modules\SysAdminTools<version>\ on an endpoint (both PowerShell 5.1 and PowerShell 7 scan that location).
Import-Module SysAdminTools in either console, then run Get-SysAdminToolsInfo and confirm Mode and Network look right.
Deployment/update tooling (push to the endpoint list, automatic archiving, health check) arrives in Phase 2 — see the roadmap at the end of the Architecture doc.
