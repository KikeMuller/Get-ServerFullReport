<!-- synced-with: README.md @ 67455a7 (2026-09-05) -->

[Español](README.md) · **English**

# Get-ServerFullReport

**As-built inventory and health check for Windows servers, in a single `.ps1`.**

[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?logo=powershell&logoColor=white)](https://learn.microsoft.com/powershell/)
[![Windows Server 2016+](https://img.shields.io/badge/Windows-Server%202016%2B%20%7C%2010%20%7C%2011-0078D6?logo=windows&logoColor=white)](#requirements)
[![MIT License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Read-only](https://img.shields.io/badge/Read--only-does%20not%20modify%20the%20host-brightgreen.svg)](#security)
[![No dependencies](https://img.shields.io/badge/Dependencies-none-lightgrey.svg)](#requirements)

Collects **104 configuration sections** from a Windows server or workstation, analyses them with **60 rules** mapped to CIS and ISO 27001 controls, and produces a navigable HTML report with an executive summary of findings at the top. It works both as documentation (a formal as-built, an audit annex) and as a review (what is wrong with this server today).

One file. No external modules, no installation, no internet access. Copy and run.

> **Note on language:** the generated report, the script's console messages and the rest of the documentation are currently in **Spanish**. An English report is on the roadmap. Section IDs, parameter names and column names are language-neutral, so the report is usable with this README at hand — but if you need English output, that is not available yet.

<p align="center">
  <img src="docs/img/reporte-light.png" alt="Report in light theme: health score, KPI cards and executive summary of findings" width="100%">
</p>

<details>
<summary>See it in dark theme</summary>
<p align="center">
  <img src="docs/img/reporte-dark.png" alt="The same report in dark theme" width="100%">
</p>
</details>

> **[See a full example report](docs/reporte-ejemplo.html)** — download it and open it in a browser (GitHub does not render embedded HTML).

---

## Contents

- [Why](#why)
- [What it does](#what-it-does)
- [Installation](#installation)
- [Usage](#usage)
- [Parameters](#parameters)
- [What it collects](#what-it-collects)
- [The findings engine](#the-findings-engine)
- [Baseline comparison (drift)](#baseline-comparison-drift)
- [Fleet mode](#fleet-mode)
- [Output formats](#output-formats)
- [Requirements](#requirements)
- [Security](#security)
- [Known issues](#known-issues)
- [Contributing](#contributing)
- [License](#license)

---

## Why

`msinfo32` exports hardware and drivers, and nothing else: no installed software, no services, no roles, no IIS or DHCP configuration, no privileged accounts, no security posture. Commercial inventory tools cost money and have to be deployed. And when an audit shows up, or somebody asks "what did this server look like before the change?", the answer is usually an out-of-date spreadsheet.

This script fills that gap: one file you copy to the server, run, and end up with a complete document plus a list of what needs fixing.

## What it does

- **104 sections** with automatic role detection: if the machine has no IIS, the section says so explicitly instead of sitting there empty.
- **Executive summary at the top**: health score, KPI cards and a findings table by severity — each finding with its evidence, the recommended action, the applicable CIS/ISO control and a link to the section it came from.
- **An OK is only emitted if the rule could actually evaluate the data.** If the section does not exist, the role is absent or the query failed, the topic stays silent. Claiming something is fine without having verified it is worse than saying nothing.
- **Self-contained HTML**: auto-generated sidebar TOC, global search, sortable and filterable tables, light/dark theme and print CSS for the signed annex. No CDNs, no external resources.
- **JSON, CSV and Markdown output** in addition to HTML, with a SHA256 hash so it can stand as evidence.
- **Comparison against a previous run**: which software, services, ports, firewall rules and accounts changed.
- **Fleet mode**: runs against dozens of servers in parallel and produces a consolidated index.
- **Read-only**: it does not modify anything at all on the audited machine.

## Installation

There is no installation. Download the file and run it:

```powershell
# Option 1: clone
git clone https://github.com/KikeMuller/Get-ServerFullReport.git
cd Get-ServerFullReport

# Option 2: grab just the script
Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/KikeMuller/Get-ServerFullReport/main/Get-ServerFullReport.ps1' -OutFile 'Get-ServerFullReport.ps1'
```

If you downloaded it through a browser, Windows marks it as a file from the internet and refuses to run it:

```powershell
Unblock-File .\Get-ServerFullReport.ps1
```

Where the execution policy is restricted, the least invasive option is to relax it for the current session only:

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
```

For recurring production use, the right answer is to **sign the script** with your organisation's code-signing certificate and distribute it from an internal repository.

## Usage

```powershell
# Local machine. Produces HTML + JSON in the script's folder.
.\Get-ServerFullReport.ps1

# A remote server, every output format
.\Get-ServerFullReport.ps1 -ComputerName SRV-APP-01 -OutputPath D:\AsBuilt -Format HTML,JSON,CSV,Markdown

# Quick security and accounts sweep across several machines
.\Get-ServerFullReport.ps1 -ComputerName SRV01,SRV02,SRV03 -Sections 1.11,1.14 -ThrottleLimit 8

# Compare against last month's snapshot
.\Get-ServerFullReport.ps1 -BaselinePath D:\AsBuilt\AsBuilt_SRV-APP-01_20260804_090000.json

# Shareable version: masks IPs, accounts, paths, serials and thumbprints
.\Get-ServerFullReport.ps1 -Redact

# Everything, including the slow parts
.\Get-ServerFullReport.ps1 -IncludeMissingUpdates -PerfSampleSeconds 60 -ScanGitRepos -EventLogDays 14
```

### As a scheduled task

```powershell
$action  = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument '-NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Get-ServerFullReport.ps1" -OutputPath "\\fs01\AsBuilt" -Quiet'
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 3am
$conf    = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest

Register-ScheduledTask -TaskName 'Weekly AsBuilt' -Action $action -Trigger $trigger -Principal $conf
```

The script returns an exit code you can chain on: `0` all good, `1` collection errors, `2` the machine was unreachable.

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-ComputerName` | `string[]` | local machine | One or more machines. More than one switches on fleet mode. |
| `-OutputPath` | `string` | script folder | Where to write the reports. |
| `-Format` | `string[]` | `HTML,JSON` | `HTML`, `JSON`, `CSV`, `Markdown`. Accepts a comma-separated list. |
| `-Sections` | `string[]` | all | Prefixes to include, e.g. `1.5,1.11`. |
| `-SkipSections` | `string[]` | none | Prefixes to skip. Useful for dropping the slow ones. |
| `-BaselinePath` | `string` | — | JSON from a previous run. Adds the change section. |
| `-Redact` | `switch` | off | Masks IPs, accounts, paths, serials and thumbprints. |
| `-ThrottleLimit` | `int` | `8` | Machines in parallel in fleet mode. |
| `-Credential` | `pscredential` | — | Credentials for the remote machines. |
| `-PamBrokerEndpoint` | `string` | — | `host:port` of the PAM broker. Without it, connectivity is not tested. |
| `-IncludeMissingUpdates` | `switch` | off | Queries missing updates. Slow: can take minutes. |
| `-EventLogDays` | `int` | `7` | Window for the event log analysis. |
| `-PerfSampleSeconds` | `int` | `0` | Seconds of performance sampling. `0` skips it. |
| `-ScanGitRepos` | `switch` | off | Looks for Git repositories. Slow on large disks. |
| `-GitScanPaths` | `string[]` | typical paths | Where to look for venvs and repos. |
| `-Quiet` | `switch` | off | No console output. For scheduled tasks. |

`Get-Help .\Get-ServerFullReport.ps1 -Full` has the complete help with examples (in Spanish).

## What it collects

104 sections across seventeen groups, with automatic role detection — a role that is not installed still gets its section, stating explicitly that it is absent, so the as-built documents that fact too.

| Group | Coverage |
|---|---|
| **1.1 – 1.5 · Machine baseline** | Hardware and serial number · OS configuration, time zone and domain role · installed hotfixes and missing updates · drivers · roles and features · installed applications · services with binary path and unquoted service path detection · firewall profiles and rules with ports and addresses · adapters, IPs, DNS, MTU, listening ports with owning process, persistent routes, hosts file and proxy · disks, volumes, BitLocker, deduplication and FSRM quotas |
| **1.6 – 1.10 · Server roles** | IIS application pools, sites and bindings · File Server shares, SMB permissions and **effective NTFS ACLs** flagging risky rights · DHCP scopes, reservations, options and statistics · DNS zones, records, forwarders and scavenging · Terminal Services / RD Licensing with 90 days of RDP access history |
| **1.11 · Security and compliance** | Certificates with CN, expiry and private key · SChannel protocols, ciphers and hashes · PAM agent · NTP · scheduled tasks with privilege level · antivirus and EDR · pending reboot · backups · OS hardening (SMBv1, SMB signing, UAC, NLA, LLMNR, NetBIOS) · TPM, Secure Boot, Device Guard and Credential Guard · LAPS · activation and licensing · suspicious root CAs |
| **1.12 – 1.14 · Directory and access** | WSUS · applied GPOs (RSOP) · privileged service accounts · local Administrators members resolved by SID · local users with password age · Remote Desktop Users group · password and lockout policy · audit policy (`auditpol`) · sensitive user rights (`secedit`) |
| **1.15 – 1.17 · Platform, dev and health** | Python interpreters, pip packages, venv, conda and Git repositories · domain and AD site, replication if it is a DC · virtualisation layer and VMware Tools version · VMs if it is a Hyper-V host · SQL Server instances · DFS · shared printers · cluster · NIC teaming · iSCSI/MPIO · WinRM configuration · uptime and page file · unexpected shutdowns · critical event log errors grouped by ID · performance baseline · top processes · SMART disk health |

The full section inventory — every section ID with its exact column names — is in [`docs/INVENTARIO_SECCIONES.md`](docs/INVENTARIO_SECCIONES.md). That file is a reference table of identifiers: the section titles and notes are in Spanish, but the IDs and column names are the same in any language.

## The findings engine

60 rules grouped into 34 compliance topics. Every finding carries a severity, concrete evidence, a recommended action, a CIS / ISO 27001 control and a link to the section it came from.

| Severity | Meaning |
|---|---|
| **CRIT** | Needs immediate action: expired certificate, disk about to fill up, SMBv1 enabled, no antivirus. |
| **WARN** | A deviation to plan for: missing patches, accounts with non-expiring passwords, unencrypted volume. |
| **INFO** | An observation worth documenting, not a defect. |
| **OK** | Verified and correct. |

The health score starts at 100 and subtracts 12 per CRIT and 4 per WARN. INFO findings do not subtract.

**The principle that governs the whole engine:** no rule emits a finding — positive or negative — about data it could not read. If the value comes back as `N/D` or "could not be determined", the topic stays silent. A false OK is worse than silence, because the sysadmin will trust it.

## Baseline comparison (drift)

```powershell
.\Get-ServerFullReport.ps1 -BaselinePath D:\AsBuilt\AsBuilt_SRV-APP-01_20260801.json
```

Adds a section with what changed since that JSON: software, services, listening ports, firewall rules, administrative accounts, certificates and scheduled tasks. It automatically ignores what varies by nature (uptime, free space, counters).

**A new account in the local Administrators group is reported as CRIT.**

This is what turns the script into a change-management tool: keep the JSON from every run, and the next one tells you what happened in between.

## Fleet mode

```powershell
$servers = (Get-ADComputer -Filter { OperatingSystem -like '*Server*' }).Name
.\Get-ServerFullReport.ps1 -ComputerName $servers -OutputPath \\fs01\AsBuilt -ThrottleLimit 12
```

Runs in parallel (runspaces on PS 5.1, `ForEach-Object -Parallel` if it detects PS 7) and produces:

- `_flota_index.html` — a matrix of the whole estate, sortable by health score
- `_flota_hallazgos.csv` — every finding from every machine, ready for a pivot table
- `_flota_inalcanzables.csv` — machines that were powered off, without WinRM or off the domain (which is a finding in itself)

Fleet mode re-launches the script as a child process per machine, so the `.ps1` has to be saved on disk — it will not work if the code was pasted straight into a console.

## Output formats

| Format | What for |
|---|---|
| **HTML** | The navigable as-built. Self-contained, printable as an audit annex. |
| **JSON** | Baselines, fleet consolidation, loading into a CMDB. UTF-8 without BOM, ISO 8601 dates, with a SHA256 hash. |
| **CSV** | One file per section, for Excel and pivot tables. |
| **Markdown** | To paste into Confluence, SharePoint or a wiki. |

## Requirements

- **PowerShell 5.1 or later.** Nothing to install: it ships with Windows Server 2016+ and Windows 10/11. If PS 7 is present it uses it for parallelism.
- **Run as Administrator.** Without it, sections 1.11, 1.14 and part of 1.5 come back incomplete; the script warns and carries on.
- **For remote machines**: WinRM enabled and port 5985/5986 reachable, with local administrator rights on the target.

Tested on Windows Server 2016, 2019, 2022 and Windows 10/11.

## Security

**The script does not modify anything** on the audited machine: it only queries. It does not use `Set-`, `New-`, `Remove-`, `Start-` or `Stop-` against the system.

**The report contains sensitive information**: service accounts, internal IP addresses, share names, certificate thumbprints, listening ports. Treat it with the same care as any architecture document:

- Store it on a share with restricted ACLs; do not email it around.
- Use `-Redact` for versions that leave the infrastructure team.
- The `-Credential` password is never written to the JSON; the account is recorded masked if you run with `-Redact`.

Every report includes audit metadata (who generated it, from which machine, script version) and the file's SHA256 hash.

## Known issues

- The BitLocker heuristic for detecting laptops relies on `TipoSistema`, which rarely reports it correctly. On laptops, check section 1.5.3 by hand.
- The performance baseline (1.17.4) depends on localised counter names. There is an index-based fallback, but on less common OS languages it may return a message instead of data.
- The repeated print header uses `position:fixed`, which repeats per page on Chromium-based browsers (Edge, Chrome) but is not guaranteed everywhere.
- The health score saturates at 0 when there are many criticals, so it does not discriminate well between "bad" and "very bad".
- Report output and script messages are Spanish-only for now.

## Contributing

Bug reports and pull requests are welcome. See [CONTRIBUTING.md](CONTRIBUTING.md) (in Spanish) — it documents the hard rules that keep the script working in production: PowerShell 5.1 as the floor, the unary comma operator, no non-ASCII characters in the `.ps1`, read-only behaviour, one try/catch per collector, and never asserting anything about data that could not be read.

**If you are reporting a bug, attach the JSON from the run** (with `-Redact` if it holds sensitive data). It contains everything collected plus the embedded execution log, which is what makes the problem reproducible without access to your machine.

## License

[MIT](LICENSE) — © 2026 Enrique Müller.

Provided as is, without warranty. It is a read-only tool, but try it on a lab machine before running it broadly in production.
