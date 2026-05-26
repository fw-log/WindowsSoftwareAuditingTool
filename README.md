Portable Software Audit Utility - Installed Apps Only
=====================================================================

It keeps the reference workbook color behavior as the default formatting, with the updated green audit fill requested by the user:
- Blue = software/version columns only
- Green = audited asset cell when installed and same version, RGB(131,226,142)
- Red = audited asset cell when missing software
- Yellow = audited asset cell when installed but different version, or version unavailable

Every audited software row in an asset column is explicitly marked green, red, or yellow.

Inventory behavior
------------------
The script reads:
- Installed Apps / Add-Remove Programs registry entries
- Main AppX/MSIX packages that appear in Installed Apps

The script does not scan:
- Program Files
- arbitrary executable files
- Start Menu shortcuts
- AppsFolder entries

Asset tag behavior
------------------
The script automatically tries to read the BIOS asset tag / serial number. If it finds one, it still asks the user to confirm that the detected value is right. Manual entry is only requested if no usable value is found or the user rejects the detected value.

Version capture
---------------
The script checks DisplayVersion, BundleVersion, ProductVersion, VersionMajor/VersionMinor, packed MSI Version values, and AppX/MSIX package versions.

Noise reduction
---------------
Common Windows inbox apps and framework/runtime/dependency packages are filtered out, including Calculator, Photos, Xbox apps, Microsoft Store framework packages, .NET runtimes, Visual C++ redistributables, WebView2 Runtime, update helpers, and driver packages.

The filter also removes app names or package IDs that start with MicrosoftWindows, Microsoft.Sec, Microsoft.Start, Microsoft.Windows, or Microsoft.Win32.

Main options
------------
1. Audit existing spreadsheet
   - Adds or updates the asset column.
   - Green blank cells mean installed and same version.
   - Red blank cells mean missing.
   - Yellow cells show the detected installed version when different.
   - Backups are stored under Backups.

2. Export filtered Installed Apps inventory to CSV
   - Exports only when selected.
   - Output goes under Exports.

Requirements
------------
- Windows PC
- Microsoft Excel installed
- PowerShell 5.1 or later
- No administrator rights required for normal use
