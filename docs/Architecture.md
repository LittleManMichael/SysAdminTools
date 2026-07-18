# SysAdminTools Architecture

**Audience:** module maintainers.
**Last updated:** 2026-07-17

This document explains how the rebuilt SysAdminTools module is put together and the rules that keep it reliable. It is written to stand alone: a new maintainer joining the team years from now should be able to read this and understand the whole system.

---

## 1. Why the rebuild

The legacy module is a single file of roughly 5,000 lines. Every change means hand-editing that one file, so one problematic edit can break every command at once, there is no change history, and there is no isolation between scripts. The arrival of PowerShell 7 made it worse: PS7 rewrites had to be published under separate command names (for example a `-ps7` suffixed copy), pushing version confusion onto the operators.

SysAdminTools replaces that with one file per command, automatic PowerShell version selection, fault isolation at load time, automatic archiving, and a single one-line-per-run audit log — while operators keep the exact same commands they use today.

## 2. Design principles

In order: simplicity over complexity, readability over cleverness, maintainability over abstraction, consistency over individual preference, PowerShell-native solutions, built-in Microsoft technologies only, and operational usefulness over software-engineering purity. The tiebreaker for any decision: choose the implementation a competent Windows System Administrator can understand the fastest.

## 3. Repository layout

```
SysAdminTools/
├── SysAdminTools.psd1        Module manifest (5.1 minimum, Desktop + Core)
├── SysAdminTools.psm1        The loader (see section 4)
├── Scripts/
│   ├── _Template.ps1         Starting point for new commands (ignored by loader)
│   ├── PS7/                  Primary implementations - one file per command
│   └── PS51/                 Compatibility implementations - one file per command
├── Helpers/                  Shared functions (Config, Paths, Logging, ...)
├── Resources/                Supporting files: Settings, HTML templates, CSV lists
│   └── Settings.sample.psd1  Settings template (real Settings.psd1 is per network)
├── Reports/                  Operational report output, organized by report type
├── Logs/                     Central audit CSVs (Audit_yyyy-MM.csv)
├── Archive/                  Prior versions of scripts (PS7/ and PS51/ mirrors)
└── Docs/                     This document, TTP, Command Reference, briefs
```

On each network, this whole tree lives on that network's authoritative admin share. Local copies on endpoints are deployments of it. `Scripts\PS7` is deliberately flat — glancing at it is a complete inventory of the toolkit.

## 4. How the loader works

`SysAdminTools.psm1` runs identically under Windows PowerShell 5.1 and PowerShell 7 (so everything in that one file must stay 5.1-compatible). At import it does four things and nothing else:

1. **Reads settings.** Loads `Resources\Settings.psd1` (a local file read). If only the sample exists, it loads that and warns loudly.
2. **Loads helpers.** Dot-sources every `.ps1` under `Helpers\`.
3. **Loads commands.** Dot-sources the correct implementation of each command per the selection rules in section 5.
4. **Exports.** `Export-ModuleMember -Function * -Alias *` — everything loaded is exported, including legacy-name aliases.

Rules the loader enforces, all of which exist to protect operations:

- **Filename = function name.** `Scripts\PS7\Get-Widget.ps1` must define `function Get-Widget`. If it doesn't, the loader warns by name at import so the mistake is caught immediately instead of surfacing as a missing command later.
- **The underscore rule.** Any `.ps1` whose filename contains an underscore is ignored by the loader. This is what makes the team's existing `_WIP.ps1` and `_BROKEN.ps1` habits safe: work-in-progress can sit right next to the live file without any risk of it loading. (In the legacy monolith, a WIP copy defining the same function name would silently hijack the real command.)
- **Fault isolation.** A file that fails to load produces a warning naming that file; every other command still loads. One bad edit no longer takes down the toolkit.
- **No side effects.** Import performs no network access, no Active Directory queries, no prompts, and writes no logs.
- **No `Set-StrictMode`.** Deliberate: many operational scripts predate strict mode and would begin failing on uninitialized-variable reads if the module scope enforced it.

Import ends with a one-line summary, e.g. `[SysAdminTools] Loaded in PowerShell 7 mode: 24 commands (9 PS7, 15 PS51)`.

## 5. Version selection rules

`Scripts\PS51` is the universal baseline — nearly all 5.1 code runs fine under PS7. `Scripts\PS7` is the override set for commands that benefit from PS7 (parallelism, newer syntax).

| Command exists in... | PowerShell 7 session | PowerShell 5.1 session |
|---|---|---|
| PS7 and PS51 | PS7 version loads | PS51 version loads |
| PS51 only | PS51 version loads | PS51 version loads |
| PS7 only | PS7 version loads | A stub loads that throws: *"The command 'X' requires PowerShell 7. Open a PowerShell 7 (pwsh) console and run it there."* |

Two consequences worth internalizing:

- **PS7 files may freely use PS7-only syntax** (ternaries, `ForEach-Object -Parallel`, `??`, chain operators). A 5.1 session never parses those files, so there is no risk of parse errors on 5.1.
- **The stub beats absence.** "Command not found" reads as a broken install; the stub reads as an instruction. `Get-SysAdminToolsInfo` lists stubbed commands under `RequirePS7`.

## 6. Per-network settings

The module deploys to three networks whose domain names, file shares, and machine naming conventions all differ. The rule that makes this manageable: **scripts never hardcode environment values.** Share paths, SMTP relay, mail addresses, domain suffixes, and server-name patterns all come from `Resources\Settings.psd1` via `Get-SysAdminToolsSetting`.

- Every network's share carries its own real `Settings.psd1` — same filename, same keys, network-specific values. It is never committed to the repo; only `Settings.sample.psd1` (placeholders) is.
- Because of this, every script file is byte-identical across all three networks. A fix made on one network is a straight file copy to the others.
- `Get-SysAdminToolsSetting` supports dot paths for nested values: `Get-SysAdminToolsSetting -Name 'ServerNamePatterns.Exchange'`. A missing key throws a clear error naming the key and pointing at the settings file.
- When a script needs a new environment value: add the key to `Settings.sample.psd1` first (placeholder), then to each network's real file.

`Get-SysAdminToolsInfo` displays which network's settings are loaded (`NetworkName`), so an admin can always confirm.

## 7. Command files

Every command file follows the same shape (see `Scripts\_Template.ps1`):

- **Header block** at the top of the file — the entire change-tracking system:

  ```powershell
  <#
  .UPDATED   2026-07-17
  .BY        mike
  .NOTES     Widened the FP server filter
  #>
  ```

  Whoever edits the file updates these three lines. No version numbers, no bump command, no hashes. Longer history lives in `Archive\`.

- **Comment-based help** (`.SYNOPSIS`, `.DESCRIPTION`, `.EXAMPLE`) inside the function so `Get-Help` works for every command. This replaces the legacy module's hand-maintained help-text variables.

- **Audit calls** wrapping the operational work (section 8).

- **Approved verbs, legacy aliases.** New and renamed commands use approved PowerShell verbs (`Get-Verb`). When a command replaces a legacy name, the old name is kept working with a `Set-Alias` line at the bottom of the file, and the rename is recorded in `Docs\Command-Reference.md` so anyone searching by the old name finds the file. Renames are decided per command at port time, not in a big-bang pass.

- **The alias line goes in every implementation file of that command.** If `Get-DriveSpaceReport` exists in both PS7 and PS51, both files carry the same `Set-Alias -Name Check-DriveSpace -Value Get-DriveSpaceReport` line. Reason: in a PS7 session the PS51 file is skipped entirely, so an alias living only there would silently disappear for PS7 users.

## 8. Audit logging

One helper, one rule: `Write-SysAdminToolsAudit` appends **one CSV row per command run** to `<ShareRoot>\Logs\Audit_yyyy-MM.csv` — Timestamp, Command, User, Computer, Result (Success/Failure), Note. Monthly files keep any single file from growing forever, and `Export-Csv -Append` handles headers and quoting automatically.

That is the entire logging system, by design. Detailed troubleshooting information belongs in console output and reports, not in log files nobody reads.

There is deliberately **no local fallback**: if the share is unreachable, the admin has bigger problems than audit retention. In that case the helper writes a warning and returns — a logging failure never blocks or fails the operational command.

## 9. Change workflow and archives

**Routine edit:** open the file on the share, make the change, update `.UPDATED` / `.BY` / `.NOTES` at the top. Done.

**Larger or risky work:** save as `<Command>_WIP.ps1` in the same folder (the loader ignores it), test, then rename over the live file. If a change goes bad, the `_BROKEN.ps1` convention and `Archive\` restore the last known-good version.

**Archives:** `Archive\PS7\<Command>\` and `Archive\PS51\<Command>\` hold dated prior copies, named `yyyy-MM-dd_<Command>.ps1`. In Phase 2 the deployment tool takes this snapshot automatically — before pushing, it archives any script whose content changed since its last archived copy, so history is captured at the natural checkpoint without anyone thinking about it. Until that tooling lands, copy the current file into Archive by hand before significant edits.

**Repository health check (Phase 2):** a maintenance command will flag (a) files whose `.UPDATED` date is older than the file's last-write time (edited without updating the header), (b) underscore files that have lingered, and (c) for awareness, commands that exist only in PS7 (meaning 5.1 sessions get the stub).

## 10. Adding a new command

1. Copy `Scripts\_Template.ps1` into `Scripts\PS51` (or `Scripts\PS7` if it needs PS7 features) as `<Verb-Noun>.ps1`, using an approved verb.
2. Rename the function inside to match the filename exactly.
3. Write the command. Pull all environment values through `Get-SysAdminToolsSetting`.
4. Fill in the header block and the comment-based help.
5. Re-import (`Import-Module SysAdminTools -Force`) and confirm it appears in the load summary and in `Get-SysAdminToolsInfo`.

Write the PS51 version only when 5.1 compatibility is actually required for that command; otherwise a single implementation in the appropriate folder is enough.

## 11. Porting a command from the legacy module

Checklist per command:

1. Decide the canonical name (approved verb). Note the legacy name for the alias and the Command Reference.
2. Extract every hardcoded environment value (paths, relay, addresses, name patterns, domain suffixes) into settings keys.
3. Wrap the work in try/catch with `Write-SysAdminToolsAudit` Success/Failure calls.
4. Add the header block and comment-based help; add the `Set-Alias` line (to every implementation file).
5. Start from the working 5.1 logic as the PS51 file. Add a PS7 file only where PS7 genuinely improves it (typically: parallel sweeps across many servers).
6. Test on the smallest safe scope, in both consoles where both implementations exist.
7. Add the row to `Docs\Command-Reference.md`.

## 12. Verifying a deployment

- Import in **both** consoles (`powershell` and `pwsh`) and check the one-line load summary in each.
- Run `Get-SysAdminToolsInfo`: confirm `Mode`, `Network` (right settings file), and that the command lists look right.
- To see the stub behavior work, note any command under `RequirePS7` in a 5.1 session and run it — it should throw the "requires PowerShell 7" instruction, not "command not found".

## 13. Deliberately not in this design

No source control on the internal networks (the Archive model covers history), no PowerShell Gallery or third-party modules, no build pipeline, no per-script version numbers or content hashes, no enterprise logging framework. Each of these was considered and rejected as complexity the team would have to carry without proportional operational benefit.

## 14. Roadmap

| Phase | Content | Status |
|---|---|---|
| 1 | Architecture, loader, manifest, settings model, core helpers, documentation baseline | Complete (pending on-domain validation) |
| 2 | Pilot commands (server availability, drive space), deployment/update tooling with automatic archiving, repository health check, session profile auto-import | Next |
| 3 | Port remaining legacy commands in priority order; PS7 parallel upgrades where beneficial | Planned |
| 4 | Cutover: retire the legacy module, update import bootstraps and profiles | Planned |
