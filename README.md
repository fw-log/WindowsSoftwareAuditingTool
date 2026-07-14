Company Endpoint Audit Utility
=====================================================================

This is a Windows-first device auditing tool designed for room, lab, department, or site audits with minimal user choices.

Default workflow
----------------
Double-click `Run-Software-Audit.bat`.

The app asks for one decision:

1. Start New Room Audit
   - Prompts for a room name.
   - Creates a new Excel workbook named after the room under `Audits`.
   - Uses the first audited computer as the baseline.
   - Writes that device into the same asset-tag column on every audit page.

2. Use Existing Audit File
   - Opens File Explorer.
   - User selects an existing audit workbook.
   - The current device is added as the next asset-tag column.
   - Asset columns stay aligned across every audit page.

Workbook layout
---------------
Each workbook owns a fixed audit layout:

- `Software`
- `Device`
- `Security`
- `Network`
- `Services`
- `Compliance`

Column A contains the audit item. Column B contains the editable baseline or expected value and is 550 pixels wide. Columns C and later are 80-pixel device asset-tag columns. All cell contents, including numeric-only values, are left aligned. The same asset tag always uses the same column on every page.

Baseline behavior
-----------------
The first computer audited in a new room workbook becomes the baseline.

For each device:

- Audited values remain visible in green cells.
- Missing or failed checks are marked red.
- Differences from baseline, new software, or review items are marked yellow.
- Comments contain supporting details where available.

Changing an editable value in column B immediately recalculates the device-cell colors through Excel conditional formatting. True/False and Yes/No baselines use dropdown validation so only supported values can be entered.

`% Space Available` is maintained on the Device sheet; the corresponding Compliance row links to it so both sheets always show the same baseline and audited values.

Software inventory sources
--------------------------
Software is collected from multiple local Windows sources:

- Add/Remove Programs registry uninstall entries under HKLM and HKCU
- WOW6432Node uninstall entries
- Main AppX/MSIX packages
- Program Files executable metadata
- Start Menu shortcut target metadata

The tool still filters common Windows inbox apps, framework/runtime packages, updates, driver packages, and other noise where practical.

Audit categories
----------------
Software:

- Baseline software list from the first audited computer
- Version comparison for later devices
- Light-grey `Additional Software` section for software not in the baseline
- Missing baseline software
- Publisher/source metadata in comments

Device:

- Computer name and domain/workgroup
- Manufacturer, model, serial number, BIOS asset tag, chassis type
- OS caption, version, build, architecture, install date, last boot, uptime, time zone
- CPU, RAM, fixed-disk percentage available, battery, BIOS, graphics adapter
- Physical disks, monitors, and printers
- Pending reboot status

Security:

- TPM
- Secure Boot
- BitLocker system drive protection
- Microsoft Defender status and signature dates
- Windows Firewall profiles
- UAC
- Remote Desktop
- SMBv1
- Security Center antivirus products
- Machine certificates expiring within 45 days

Network:

- DNS host entry
- Active IP addresses
- MAC addresses
- Default gateways
- DNS servers
- DHCP-enabled adapters
- Network profiles
- Wi-Fi SSID, when available
- Mapped network drives

Services:

- Microsoft Defender Antivirus
- Windows Firewall
- Windows Update
- BITS
- Intune Management Extension
- Configuration Manager client
- Microsoft Defender for Endpoint sensor

Company-specific required services can be added near the top of `SoftwareAudit.ps1` in `$RequiredCompanyServices`.

Compliance:

- Rollup of important device, security, and service checks
- Dynamic overall compliance summary
- Disk-space percentage (linked to the Device sheet), pending reboot, TPM, Secure Boot, BitLocker, Defender, firewall, UAC, SMBv1, antivirus, and required-service status
- Enabled local users, disabled local users, and Local Administrators members

Generated output
----------------
- New room workbooks are saved under `Audits`.
- Existing workbooks are backed up under `Backups` before being updated.
- Each app launch creates a troubleshooting log under `Logs`.

Troubleshooting logs
--------------------
Logs are named `SoftwareAudit_<computer>_<timestamp>_PID<process>.log`. They record:

- App startup and PowerShell/OS details
- Software inventory counts by source
- Progress through each audit collector
- Excel startup, workbook open, backup, worksheet update, save, and cleanup operations
- Detailed exception type, location, PowerShell stack, and inner exceptions

The GUI displays the current log path in its activity pane. Console mode prints the path at startup. Logging is best-effort and does not stop an audit if the log cannot be written.

Requirements
------------
- Windows PC
- Microsoft Excel installed
- PowerShell 5.1 or later

Running from PowerShell
-----------------------
GUI:

```powershell
powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -File .\SoftwareAudit.ps1
```

Console fallback:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\SoftwareAudit.ps1 -Console
```
