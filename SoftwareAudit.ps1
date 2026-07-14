<#
Company Endpoint Audit Utility
- GUI-first Windows endpoint audit tool.
- Combines Installed Apps, AppX/MSIX, Program Files metadata, and Start Menu target metadata.
- Filters duplicate entries and common Windows/framework noise.
- Requires Microsoft Excel on the audited PC.
#>

#requires -version 5.1
[CmdletBinding()]
param(
    [switch]$Console
)

$ErrorActionPreference = "Stop"

# Original workbook colors. These are Excel/VBA RGB() integer values.
# Blue and green are the same light-tinted Office theme colors used in the reference workbook.
$OriginalBlueColor   = 15653023  # RGB(159,216,238) - software/version columns
$OriginalGreenColor  = 9364099   # RGB(131,226,142) - installed/same-version asset cells
$OriginalYellowColor = 65535     # RGB(255,255,0)   - version mismatch cells
$OriginalRedColor    = 255       # RGB(255,0,0)     - missing software cells
$SectionGrayColor    = 14277081  # RGB(217,217,217) - section separator rows
$GridBorderColor     = 14277081  # RGB(217,217,217) - light borders between populated cells
$AuditColumnPixels   = 80
$BaselineColumnPixels = 550
$AuditRowHeightPoints = 15

$FilterWindowsDefaultsAndFrameworks = $true

# Add company-specific services here to make them hard compliance requirements.
# Example: @{ Name = 'YourAgentServiceName'; Display = 'Your Agent'; Required = $true }
$RequiredCompanyServices = @()

$script:TroubleshootingLogPath = ''

function Initialize-TroubleshootingLog {
    try {
        $scriptRoot = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.ScriptName }
        $logDirectory = Join-Path $scriptRoot 'Logs'
        if (-not (Test-Path $logDirectory)) {
            New-Item -ItemType Directory -Path $logDirectory -Force -ErrorAction Stop | Out-Null
        }

        $computer = if ([string]::IsNullOrWhiteSpace($env:COMPUTERNAME)) { 'UnknownComputer' } else { $env:COMPUTERNAME }
        $safeComputer = $computer -replace '[^a-zA-Z0-9_-]', '_'
        $fileName = 'SoftwareAudit_{0}_{1}_PID{2}.log' -f $safeComputer, (Get-Date -Format 'yyyyMMdd_HHmmss'), $PID
        $script:TroubleshootingLogPath = Join-Path $logDirectory $fileName
        New-Item -ItemType File -Path $script:TroubleshootingLogPath -Force -ErrorAction Stop | Out-Null

        Write-TroubleshootingLog -Category 'Startup' -Message 'Troubleshooting log initialized.'
        Write-TroubleshootingLog -Category 'Startup' -Message ("Script: {0}" -f $MyInvocation.ScriptName)
        Write-TroubleshootingLog -Category 'Startup' -Message ("PowerShell: {0}; Process architecture: {1}; OS: {2}" -f $PSVersionTable.PSVersion, $(if ([Environment]::Is64BitProcess) { '64-bit' } else { '32-bit' }), [Environment]::OSVersion.VersionString)
    } catch {
        $script:TroubleshootingLogPath = ''
    }
}

function Write-TroubleshootingLog {
    param(
        [string]$Message,
        [ValidateSet('DEBUG','INFO','WARN','ERROR')]
        [string]$Level = 'INFO',
        [string]$Category = 'General'
    )
    if ([string]::IsNullOrWhiteSpace($script:TroubleshootingLogPath)) { return }
    try {
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
        $text = if ($null -eq $Message) { '' } else { $Message.ToString() }
        foreach ($line in @($text -split "`r?`n")) {
            Add-Content -Path $script:TroubleshootingLogPath -Value ("[{0}] [{1}] [{2}] {3}" -f $timestamp, $Level, $Category, $line) -Encoding UTF8 -ErrorAction Stop
        }
    } catch {
        # Logging is diagnostic only and must never stop an audit.
    }
}

function Write-TroubleshootingException {
    param(
        $ErrorRecord,
        [string]$Context = 'Unhandled error'
    )
    if ($null -eq $ErrorRecord) {
        Write-TroubleshootingLog -Level 'ERROR' -Category 'Exception' -Message $Context
        return
    }

    Write-TroubleshootingLog -Level 'ERROR' -Category 'Exception' -Message ("{0}: {1}" -f $Context, $ErrorRecord.Exception.Message)
    Write-TroubleshootingLog -Level 'ERROR' -Category 'Exception' -Message ("Exception type: {0}" -f $ErrorRecord.Exception.GetType().FullName)
    if (-not [string]::IsNullOrWhiteSpace($ErrorRecord.FullyQualifiedErrorId)) {
        Write-TroubleshootingLog -Level 'ERROR' -Category 'Exception' -Message ("Error ID: {0}" -f $ErrorRecord.FullyQualifiedErrorId)
    }
    if ($ErrorRecord.InvocationInfo -and -not [string]::IsNullOrWhiteSpace($ErrorRecord.InvocationInfo.PositionMessage)) {
        Write-TroubleshootingLog -Level 'ERROR' -Category 'Exception' -Message ($ErrorRecord.InvocationInfo.PositionMessage.Trim())
    }
    if (-not [string]::IsNullOrWhiteSpace($ErrorRecord.ScriptStackTrace)) {
        Write-TroubleshootingLog -Level 'ERROR' -Category 'Exception' -Message ("PowerShell stack:`n{0}" -f $ErrorRecord.ScriptStackTrace)
    }

    $inner = $ErrorRecord.Exception.InnerException
    while ($inner) {
        Write-TroubleshootingLog -Level 'ERROR' -Category 'Exception' -Message ("Inner exception ({0}): {1}" -f $inner.GetType().FullName, $inner.Message)
        $inner = $inner.InnerException
    }
}

function Normalize-Name {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return "" }
    $n = $Name.ToLowerInvariant()
    $n = $n -replace '&', 'and'
    $n = $n -replace '\+', 'plus'
    $n = $n -replace '[^a-z0-9]+', ''
    return $n
}

function Normalize-VersionText {
    param($Version)
    if ($null -eq $Version) { return "" }
    return ($Version.ToString().Trim() -replace '\s+', ' ')
}

function Test-NonEmptyVersion {
    param($Version)
    $v = Normalize-VersionText $Version
    if ([string]::IsNullOrWhiteSpace($v)) { return $false }
    if ($v -match '^(installed|unknown|n/a|null|0|0\.0)$') { return $false }
    return $true
}

function Remove-TrailingZeroVersionSegments {
    param([string]$Version)
    $v = Normalize-VersionText $Version
    if ($v -notmatch '^\d+(\.\d+)+$') { return $v.ToLowerInvariant() }
    while ($v -match '^(.+)\.0$') { $v = $matches[1] }
    return $v.ToLowerInvariant()
}

function Test-VersionMatch {
    param([string]$Expected, [string]$Installed)
    $e = Normalize-VersionText $Expected
    $i = Normalize-VersionText $Installed
    if ([string]::IsNullOrWhiteSpace($e) -or [string]::IsNullOrWhiteSpace($i)) { return $false }
    if ($i -eq 'Installed') { return $false }
    if ($e.Equals($i, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    if ((Remove-TrailingZeroVersionSegments $e) -eq (Remove-TrailingZeroVersionSegments $i)) { return $true }
    return $false
}

function Convert-PackedMsiVersion {
    param($PackedVersion)
    if ($null -eq $PackedVersion) { return "" }
    try {
        [UInt32]$num = [UInt32]$PackedVersion
        if ($num -eq 0) { return "" }
        $major = ($num -shr 24) -band 0xFF
        $minor = ($num -shr 16) -band 0xFF
        $build = $num -band 0xFFFF
        if ($major -eq 0 -and $minor -eq 0 -and $build -eq 0) { return "" }
        return "$major.$minor.$build"
    } catch { return "" }
}

function Get-RegistryDisplayVersion {
    param($Item)
    foreach ($p in @('DisplayVersion','BundleVersion','ProductVersion')) {
        try { if (Test-NonEmptyVersion $Item.$p) { return (Normalize-VersionText $Item.$p) } } catch {}
    }
    try {
        if ($null -ne $Item.VersionMajor -and $null -ne $Item.VersionMinor) { return ("{0}.{1}" -f $Item.VersionMajor, $Item.VersionMinor) }
    } catch {}
    try {
        $packed = Convert-PackedMsiVersion $Item.Version
        if (Test-NonEmptyVersion $packed) { return $packed }
    } catch {}
    try {
        if ($Item.QuietDisplayName -match '(\d+(\.\d+){1,5}([\+\-][A-Za-z0-9\.]+)?)') { return $matches[1] }
    } catch {}
    return ""
}

function New-AppEntry {
    param(
        [string]$Name,
        [string]$Version = "",
        [string]$Publisher = "",
        [string]$Source = "Installed Apps",
        [string]$Id = ""
    )
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    $cleanName = $Name.Trim()
    [PSCustomObject]@{
        Name       = $cleanName
        Version    = Normalize-VersionText $Version
        Publisher  = if ($Publisher) { $Publisher.ToString().Trim() } else { "" }
        Source     = $Source
        Id         = $Id
        Normalized = Normalize-Name $cleanName
    }
}

$ExplicitKeepNamePatterns = @(
    '^microsoft\s+edge$',
    '^microsoft\s+onedrive$',
    '^microsoft\s+teams$',
    '^microsoft\s+visual\s+studio\s+code$',
    '^visual\s+studio\s+code$',
    '^microsoft\s+visual\s+studio\s+tools\s+for\s+applications',
    '^power\s+automate$',
    '^r\s+for\s+windows',
    '^rstudio$',
    '^nvidia\s+physx',
    '^uxp\s+webview\s+support$'
)

$DefaultWindowsAppNamePatterns = @(
    '^alarms?\s*&?\s*clock$', '^calculator$', '^camera$', '^clipchamp$', '^clock$', '^copilot$', '^cortana$',
    '^dev\s+home$', '^family$', '^feedback\s+hub$', '^get\s+help$', '^get\s+started$', '^mail\s+and\s+calendar$',
    '^maps$', '^media\s+player$', '^microsoft\s+365\s*\(office\)$', '^microsoft\s+bing', '^microsoft\s+clipchamp$',
    '^microsoft\s+copilot$', '^microsoft\s+family$', '^microsoft\s+news$', '^microsoft\s+photos$', '^microsoft\s+solitaire',
    '^microsoft\s+store$', '^microsoft\s+to\s+do$', '^microsoft\s+whiteboard$', '^mixed\s+reality\s+portal$',
    '^movies\s*&\s*tv$', '^msn\s+', '^notepad$', '^office$', '^onenote$', '^outlook\s*\(new\)$', '^paint$',
    '^people$', '^phone\s+link$', '^quick\s+assist$', '^remote\s+desktop$', '^screen\s+sketch$', '^snipping\s+tool$',
    '^sound\s+recorder$', '^sticky\s+notes$', '^terminal$', '^tips$', '^voice\s+recorder$', '^weather$',
    '^windows\s+(security|backup|camera|maps|media|notepad|photos|terminal|web\s+experience)', '^xbox', '^your\s+phone$'
)


$ExcludedMicrosoftPackagePrefixes = @(
    'MicrosoftWindows',
    'Microsoft.Sec',
    'Microsoft.Start',
    'Microsoft.Windows',
    'Microsoft.Win32'
)

$DefaultWindowsPackagePatterns = @(
    '^Clipchamp\.', '^Microsoft\.WindowsAlarms$', '^Microsoft\.WindowsCalculator$', '^Microsoft\.WindowsCamera$',
    '^Microsoft\.WindowsMaps$', '^Microsoft\.WindowsSoundRecorder$', '^Microsoft\.WindowsFeedbackHub$', '^Microsoft\.GetHelp$',
    '^Microsoft\.Getstarted$', '^Microsoft\.People$', '^Microsoft\.Todos$', '^Microsoft\.MicrosoftSolitaireCollection$',
    '^Microsoft\.MicrosoftStickyNotes$', '^Microsoft\.WindowsNotepad$', '^Microsoft\.Paint$', '^Microsoft\.MSPaint$',
    '^Microsoft\.ScreenSketch$', '^Microsoft\.SnippingTool$', '^Microsoft\.Windows\.Photos$', '^Microsoft\.Photos$',
    '^Microsoft\.Bing(News|Weather|Finance|Sports)', '^Microsoft\.Zune(Music|Video)$', '^Microsoft\.Xbox', '^Microsoft\.YourPhone$',
    '^Microsoft\.CorporationII\.QuickAssist$', '^Microsoft\.549981C3F5F10$', '^Microsoft\.Windows\.DevHome$',
    '^Microsoft\.WindowsTerminal$', '^Microsoft\.WindowsStore$', '^Microsoft\.StorePurchaseApp$', '^Microsoft\.DesktopAppInstaller$',
    '^Microsoft\.WebMediaExtensions$', '^Microsoft\.WebpImageExtension$', '^Microsoft\.HEIFImageExtension$',
    '^Microsoft\.HEVCVideoExtension$', '^Microsoft\.VP9VideoExtensions$', '^Microsoft\.VCLibs', '^Microsoft\.UI\.Xaml',
    '^Microsoft\.NET\.Native', '^Microsoft\.Services\.Store'
)

$FrameworkPatterns = @(
    '^microsoft\s+\.net', '^\.net\s+', 'asp\.net', 'windows\s+desktop\s+runtime',
    'microsoft\s+edge\s+webview2\s+runtime', 'webview2\s+runtime',
    '^microsoft\s+visual\s+c\+\+.*redistributable', '^microsoft\s+visual\s+c\+\+.*runtime',
    'visual\s+c\+\+\s+redistributable', '\bvc\+\+\b', '\bvc_redist\b', '^microsoft\s+vclibs',
    '^microsoft\s+ui\s+xaml', '^microsoft\s+windows\s+app\s+runtime', '^windows\s+app\s+runtime',
    '^microsoft\s+update\s+health\s+tools$', '^update\s+for\s+', '^security\s+update\s+for\s+', '^hotfix\s+for\s+',
    '^servicing\s+stack', '^kb\d+', 'language\s+pack$', 'click-to-run\s+extensibility\s+component',
    'office\s+shared\s+.*\s+component', '^microsoft\s+edge\s+update$', 'mozilla\s+maintenance\s+service',
    'google\s+update\s+helper', '^intel\(r\).*driver', '^intel\s+.*driver', '^amd\s+.*driver',
    '^nvidia\s+(graphics|hd\s+audio|frameview).*driver', '^realtek\s+.*driver'
)


function Test-ExcludedMicrosoftPackagePrefix {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    $value = $Text.Trim()
    foreach ($prefix in $ExcludedMicrosoftPackagePrefixes) {
        if ($value.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Test-ExplicitKeep {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    $n = $Name.Trim().ToLowerInvariant()
    foreach ($p in $ExplicitKeepNamePatterns) { if ($n -match $p) { return $true } }
    return $false
}

function Test-DefaultWindowsOrFramework {
    param($App)
    if (-not $FilterWindowsDefaultsAndFrameworks) { return $false }
    if ($null -eq $App -or [string]::IsNullOrWhiteSpace($App.Name)) { return $true }
    if (Test-ExplicitKeep $App.Name) { return $false }

    $name = $App.Name.Trim()
    $lower = $name.ToLowerInvariant()
    $id = if ($App.Id) { $App.Id.ToString() } else { "" }
    if ((Test-ExcludedMicrosoftPackagePrefix $name) -or (Test-ExcludedMicrosoftPackagePrefix $id)) { return $true }
    foreach ($p in $DefaultWindowsAppNamePatterns) { if ($lower -match $p) { return $true } }
    foreach ($p in $DefaultWindowsPackagePatterns) { if ($name -match $p -or $id -match $p) { return $true } }
    foreach ($p in $FrameworkPatterns) { if ($lower -match $p) { return $true } }
    return $false
}

function Test-VisibleRegistryInstalledApp {
    param($Item)
    if ($null -eq $Item.DisplayName -or [string]::IsNullOrWhiteSpace($Item.DisplayName.ToString())) { return $false }
    try { if ($Item.SystemComponent -eq 1) { return $false } } catch {}
    try { if ($Item.NoDisplay -eq 1) { return $false } } catch {}
    try { if ($Item.ReleaseType -match 'Update|Hotfix|Security|Service Pack') { return $false } } catch {}
    try { if ($Item.ParentKeyName) { return $false } } catch {}
    return $true
}

function Get-InstalledAppsFromRegistry {
    $apps = @()
    $paths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    foreach ($path in $paths) {
        try {
            foreach ($item in @(Get-ItemProperty -Path $path -ErrorAction SilentlyContinue)) {
                if (-not (Test-VisibleRegistryInstalledApp $item)) { continue }
                $app = New-AppEntry -Name ($item.DisplayName) -Version (Get-RegistryDisplayVersion $item) -Publisher ($item.Publisher) -Source 'Installed Apps - Registry' -Id ($item.PSChildName)
                if ($app -and -not (Test-DefaultWindowsOrFramework $app)) { $apps += $app }
            }
        } catch {}
    }
    return $apps
}

function Get-AppxFriendlyName {
    param($Package)
    $names = @()
    try {
        $manifest = Get-AppxPackageManifest -Package $Package.PackageFullName -ErrorAction SilentlyContinue
        if ($manifest.Package.Properties.DisplayName) { $names += $manifest.Package.Properties.DisplayName.ToString() }
    } catch {}
    try { $names += $Package.Name } catch {}
    foreach ($n in $names) {
        if ([string]::IsNullOrWhiteSpace($n)) { continue }
        $clean = $n.Trim()
        if ($clean -match '^ms-resource:' -or $clean -match '^@\{') { continue }
        return $clean
    }
    try { return $Package.Name } catch { return "" }
}

function Get-InstalledAppsFromAppx {
    $apps = @()
    try { $packages = @(Get-AppxPackage -PackageTypeFilter Main -ErrorAction SilentlyContinue) }
    catch { try { $packages = @(Get-AppxPackage -ErrorAction SilentlyContinue | Where-Object { -not $_.IsFramework }) } catch { $packages = @() } }

    foreach ($pkg in $packages) {
        try { if ($pkg.IsFramework) { continue } } catch {}
        $name = Get-AppxFriendlyName $pkg
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $publisher = ""; try { $publisher = $pkg.Publisher } catch {}
        $id = ""; try { $id = $pkg.Name } catch {}
        $app = New-AppEntry -Name $name -Version ($pkg.Version) -Publisher $publisher -Source 'Installed Apps - Store/AppX' -Id $id
        if ($app -and -not (Test-DefaultWindowsOrFramework $app)) { $apps += $app }
    }
    return $apps
}

function Get-InstalledAppsFromProgramFiles {
    $apps = @()
    $roots = @($env:ProgramFiles, ${env:ProgramFiles(x86)}) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path $_) } | Select-Object -Unique
    foreach ($root in $roots) {
        try {
            foreach ($vendorDir in @(Get-ChildItem -Path $root -Directory -ErrorAction SilentlyContinue)) {
                $exeFiles = @()
                try { $exeFiles += @(Get-ChildItem -Path $vendorDir.FullName -Filter '*.exe' -File -ErrorAction SilentlyContinue | Select-Object -First 3) } catch {}
                try {
                    foreach ($childDir in @(Get-ChildItem -Path $vendorDir.FullName -Directory -ErrorAction SilentlyContinue | Select-Object -First 12)) {
                        $exeFiles += @(Get-ChildItem -Path $childDir.FullName -Filter '*.exe' -File -ErrorAction SilentlyContinue | Select-Object -First 2)
                    }
                } catch {}
                foreach ($exe in @($exeFiles | Select-Object -Unique FullName)) {
                    try {
                        $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($exe.FullName)
                        $name = $versionInfo.ProductName
                        if ([string]::IsNullOrWhiteSpace($name)) { $name = $versionInfo.FileDescription }
                        if ([string]::IsNullOrWhiteSpace($name)) { continue }
                        $version = $versionInfo.ProductVersion
                        if ([string]::IsNullOrWhiteSpace($version)) { $version = $versionInfo.FileVersion }
                        $app = New-AppEntry -Name $name -Version $version -Publisher ($versionInfo.CompanyName) -Source 'Program Files executable metadata' -Id ($exe.FullName)
                        if ($app -and -not (Test-DefaultWindowsOrFramework $app)) { $apps += $app }
                    } catch {}
                }
            }
        } catch {}
    }
    return $apps
}

function Get-InstalledAppsFromStartMenu {
    $apps = @()
    $paths = @(
        [Environment]::GetFolderPath('StartMenu'),
        [Environment]::GetFolderPath('CommonStartMenu')
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path $_) } | Select-Object -Unique
    $shell = $null
    try { $shell = New-Object -ComObject WScript.Shell } catch { return $apps }
    foreach ($path in $paths) {
        foreach ($link in @(Get-ChildItem -Path $path -Filter '*.lnk' -File -Recurse -ErrorAction SilentlyContinue)) {
            try {
                $shortcut = $shell.CreateShortcut($link.FullName)
                $target = $shortcut.TargetPath
                if ([string]::IsNullOrWhiteSpace($target) -or -not (Test-Path $target) -or $target -notmatch '\.exe$') { continue }
                $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($target)
                $name = $versionInfo.ProductName
                if ([string]::IsNullOrWhiteSpace($name)) { $name = $versionInfo.FileDescription }
                if ([string]::IsNullOrWhiteSpace($name)) { $name = [System.IO.Path]::GetFileNameWithoutExtension($link.Name) }
                $version = $versionInfo.ProductVersion
                if ([string]::IsNullOrWhiteSpace($version)) { $version = $versionInfo.FileVersion }
                $app = New-AppEntry -Name $name -Version $version -Publisher ($versionInfo.CompanyName) -Source 'Start Menu shortcut target metadata' -Id $target
                if ($app -and -not (Test-DefaultWindowsOrFramework $app)) { $apps += $app }
            } catch {}
        }
    }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null } catch {}
    return $apps
}

function Get-AppVersionScore { param($App) if (Test-NonEmptyVersion $App.Version) { return 20 } return 0 }
function Get-AppSourceScore {
    param($App)
    if ($App.Source -like '*Registry*') { return 30 }
    if ($App.Source -like '*Store*' -or $App.Source -like '*AppX*') { return 20 }
    if ($App.Source -like '*Program Files*') { return 10 }
    if ($App.Source -like '*Start Menu*') { return 8 }
    return 1
}

function Merge-InstalledApps {
    param([array]$Apps)
    return @($Apps |
        Where-Object { $_ -and $_.Normalized -ne "" } |
        Group-Object Normalized |
        ForEach-Object {
            $_.Group | Sort-Object `
                @{Expression={Get-AppVersionScore $_};Descending=$true}, `
                @{Expression={Get-AppSourceScore $_};Descending=$true}, `
                @{Expression={if ($_.Name) { $_.Name.Length } else { 9999 }};Ascending=$true}, `
                Name | Select-Object -First 1
        } |
        Sort-Object Name)
}

function Get-InstalledApps {
    Write-TroubleshootingLog -Category 'Software' -Message 'Starting installed software collection.'
    $registryApps = @(Get-InstalledAppsFromRegistry)
    Write-TroubleshootingLog -Category 'Software' -Message ("Registry entries collected: {0}" -f $registryApps.Count)
    $appxApps = @(Get-InstalledAppsFromAppx)
    Write-TroubleshootingLog -Category 'Software' -Message ("AppX/MSIX entries collected: {0}" -f $appxApps.Count)
    $programFilesApps = @(Get-InstalledAppsFromProgramFiles)
    Write-TroubleshootingLog -Category 'Software' -Message ("Program Files metadata entries collected: {0}" -f $programFilesApps.Count)
    $startMenuApps = @(Get-InstalledAppsFromStartMenu)
    Write-TroubleshootingLog -Category 'Software' -Message ("Start Menu metadata entries collected: {0}" -f $startMenuApps.Count)
    $mergedApps = @(Merge-InstalledApps -Apps @(($registryApps) + ($appxApps) + ($programFilesApps) + ($startMenuApps)))
    Write-TroubleshootingLog -Category 'Software' -Message ("Installed software collection complete; merged entries: {0}" -f $mergedApps.Count)
    return $mergedApps
}

function Get-RequiredAliases {
    param([string]$Name)
    $req = Normalize-Name $Name
    $aliases = @($req)
    $map = @{
        'adobepremierpro' = @('adobepremierepro')
        'adobepremierpro2025' = @('adobepremierepro2025')
        'adobeacrobat64bit' = @('adobeacrobat')
        'microsoftvisualstudiocode' = @('visualstudiocode')
        'vscode' = @('visualstudiocode')
        'adobedigitaledition' = @('adobedigitaleditions')
        'adobedigitaledition45' = @('adobedigitaleditions45')
        'ibmspssstatics' = @('ibmspssstatistics')
    }
    if ($map.ContainsKey($req)) { $aliases += $map[$req] }
    return @($aliases | Where-Object { $_ } | Select-Object -Unique)
}

function Get-MatchScore {
    param([string]$Key, $App)
    if ([string]::IsNullOrWhiteSpace($Key) -or $null -eq $App -or [string]::IsNullOrWhiteSpace($App.Normalized)) { return 0 }
    $i = $App.Normalized
    if ($i -eq $Key) { return 100 }
    if ($i.StartsWith($Key)) { return 85 }
    if ($Key.StartsWith($i) -and $i.Length -ge 6) { return 80 }
    if ($i.Contains($Key) -and $Key.Length -ge 6) { return 70 }
    if ($Key.Contains($i) -and $i.Length -ge 8) { return 65 }
    return 0
}

function Find-AppMatch {
    param([string]$RequiredName, [array]$InstalledApps)
    $hits = @()
    foreach ($app in $InstalledApps) {
        foreach ($key in (Get-RequiredAliases $RequiredName)) {
            $score = Get-MatchScore -Key $key -App $app
            if ($score -gt 0) {
                $hits += [PSCustomObject]@{ App=$app; Score=$score; VersionScore=(Get-AppVersionScore $app); SourceScore=(Get-AppSourceScore $app); Len=$app.Name.Length }
            }
        }
    }
    if ($hits.Count -eq 0) { return $null }
    return ($hits | Sort-Object @{Expression={$_.Score};Descending=$true}, @{Expression={$_.VersionScore};Descending=$true}, @{Expression={$_.SourceScore};Descending=$true}, @{Expression={$_.Len};Ascending=$true} | Select-Object -First 1).App
}

function Set-FillColor {
    param($RangeOrCell, [int]$Color)
    try {
        $RangeOrCell.Interior.Pattern = 1
        $RangeOrCell.Interior.Color = $Color
    } catch {}
}

function Clear-CellComment {
    param($Cell)
    try { if ($Cell.Comment()) { $Cell.Comment().Delete() | Out-Null } } catch {}
}

function Set-CellComment {
    param($Cell, [string]$Text)
    Clear-CellComment -Cell $Cell
    if (-not [string]::IsNullOrWhiteSpace($Text)) { try { $Cell.AddComment($Text) | Out-Null } catch {} }
}

# Audited asset cells are always explicitly marked as one of three outcomes: green, red, or yellow.
function Set-AuditSame {
    param($Cell)
    try { $Cell.ClearContents() | Out-Null } catch {}
    Clear-CellComment -Cell $Cell
    Set-FillColor -RangeOrCell $Cell -Color $OriginalGreenColor
}

function Set-AuditMissing {
    param($Cell, [string]$Comment)
    try { $Cell.ClearContents() | Out-Null } catch {}
    Set-FillColor -RangeOrCell $Cell -Color $OriginalRedColor
    Set-CellComment -Cell $Cell -Text $Comment
}

function Set-AuditDifferent {
    param($Cell, [string]$Version, [string]$Comment)
    $value = if (Test-NonEmptyVersion $Version) { $Version } else { 'Installed - version unavailable' }
    try { $Cell.Value2 = [string]$value } catch {}
    Set-FillColor -RangeOrCell $Cell -Color $OriginalYellowColor
    Set-CellComment -Cell $Cell -Text $Comment
}

function Apply-OriginalDefaultFormatting {
    param($Sheet, [int]$FirstAppRow, [int]$LastAppRow, [int]$LastCol)
    if ($LastAppRow -lt $FirstAppRow) { return }
    try {
        $blueEnd = [Math]::Min(2, $LastCol)
        if ($blueEnd -ge 1) { Set-FillColor -RangeOrCell ($Sheet.Range($Sheet.Cells.Item($FirstAppRow,1), $Sheet.Cells.Item($LastAppRow,$blueEnd))) -Color $OriginalBlueColor }
        if ($LastCol -ge 3) { Set-FillColor -RangeOrCell ($Sheet.Range($Sheet.Cells.Item($FirstAppRow,3), $Sheet.Cells.Item($LastAppRow,$LastCol))) -Color $OriginalGreenColor }
    } catch {}
}

function Get-AutoAssetTag {
    $bad = @('', 'default string', 'none', 'no asset tag', 'no asset information', 'to be filled by o.e.m.', 'system serial number', 'unknown')
    try {
        $tag = (Get-CimInstance Win32_SystemEnclosure -ErrorAction Stop | Select-Object -First 1 -ExpandProperty SMBIOSAssetTag)
        if ($tag) { $tag = $tag.ToString().Trim(); if ($bad -notcontains $tag.ToLowerInvariant()) { return $tag } }
    } catch {}
    try {
        $serial = (Get-CimInstance Win32_BIOS -ErrorAction Stop | Select-Object -First 1 -ExpandProperty SerialNumber)
        if ($serial) { $serial = $serial.ToString().Trim(); if ($bad -notcontains $serial.ToLowerInvariant()) { return $serial } }
    } catch {}
    return $null
}

function Get-AssetTagChoice {
    $auto = Get-AutoAssetTag
    if ($auto) {
        Write-Host "Detected asset tag / serial: $auto" -ForegroundColor Green
        $accept = Read-Host "Use this value? [Y/n]"
        if ([string]::IsNullOrWhiteSpace($accept) -or $accept.Trim().ToLowerInvariant().StartsWith('y')) { return $auto }
    } else {
        Write-Host "No usable BIOS asset tag was found. Switching to manual entry." -ForegroundColor Yellow
    }

    do {
        $manual = (Read-Host "Enter asset tag for this device").Trim()
        if ([string]::IsNullOrWhiteSpace($manual)) { Write-Host "Asset tag cannot be blank." -ForegroundColor Yellow }
    } until (-not [string]::IsNullOrWhiteSpace($manual))
    return $manual
}

function Get-MainMenuChoice {
    Write-Host "Company Endpoint Audit Utility" -ForegroundColor Cyan
    Write-Host "  1. Run full room-based audit workbook"
    Write-Host "  2. Export filtered Installed Apps inventory to CSV"
    do { $choice = Read-Host "Choose 1 or 2" } until ($choice -in @('1','2'))
    return $choice
}

function Get-WorkbookPath {
    $dir = Split-Path -Parent $MyInvocation.ScriptName
    $files = @(Get-ChildItem -Path $dir -Filter '*.xlsx' -File | Where-Object { $_.Name -notlike '~$*' -and $_.Name -notlike '*_backup_*' -and $_.FullName -notmatch '\\Backups\\' -and $_.FullName -notmatch '\\Exports\\' })
    if ($files.Count -eq 1) { return $files[0].FullName }
    if ($files.Count -gt 1) {
        Write-Host "Multiple spreadsheets found:" -ForegroundColor Cyan
        for ($i=0; $i -lt $files.Count; $i++) { Write-Host ("  {0}. {1}" -f ($i+1), $files[$i].Name) }
        do { $raw = Read-Host "Select spreadsheet number"; $n=0; $ok=[int]::TryParse($raw,[ref]$n) -and $n -ge 1 -and $n -le $files.Count } until ($ok)
        return $files[$n-1].FullName
    }
    $manual = Read-Host "Workbook path"
    if (-not (Test-Path $manual)) { throw "Workbook not found: $manual" }
    return $manual
}

function Get-WorksheetSelection {
    param($Workbook)
    $sheets = @(); for ($i=1; $i -le $Workbook.Worksheets.Count; $i++) { $sheets += $Workbook.Worksheets.Item($i) }
    if ($sheets.Count -eq 1) { return @($sheets[0]) }
    Write-Host "Worksheets found:" -ForegroundColor Cyan
    for ($i=0; $i -lt $sheets.Count; $i++) { Write-Host ("  {0}. {1}" -f ($i+1), $sheets[$i].Name) }
    Write-Host "  A. All worksheets"
    do {
        $raw = Read-Host "Select worksheet number or A"
        if ($raw -match '^[Aa]$') { return @($sheets) }
        $n=0; $ok=[int]::TryParse($raw,[ref]$n) -and $n -ge 1 -and $n -le $sheets.Count
    } until ($ok)
    return @($sheets[$n-1])
}

function Test-HeaderRow {
    param($Sheet)
    $a = Normalize-Name ($Sheet.Cells.Item(1,1).Text)
    $b = Normalize-Name ($Sheet.Cells.Item(1,2).Text)
    return (($a -in @('software','requiredsoftware','softwarename','audititem')) -and ($b -in @('version','expectedversion','expectedbaseline')))
}

function Get-UsedLastRow { param($Sheet) try { $u=$Sheet.UsedRange; return ($u.Row + $u.Rows.Count - 1) } catch { return 1 } }
function Get-UsedLastCol { param($Sheet) try { $u=$Sheet.UsedRange; return ($u.Column + $u.Columns.Count - 1) } catch { return 2 } }

function Get-AppLastRow {
    param($Sheet, [int]$FirstAppRow)
    $last = Get-UsedLastRow $Sheet
    for ($r=$FirstAppRow; $r -le $last; $r++) {
        $v = $Sheet.Cells.Item($r,1).Text.Trim()
        if ([string]::IsNullOrWhiteSpace($v) -or $v -match '^notes?:\s*$') { return ($r-1) }
    }
    return $last
}

function Get-SheetLayout {
    param($Sheet)
    $hasHeader = Test-HeaderRow $Sheet
    $first = if ($hasHeader) { 2 } else { 1 }
    $last = Get-AppLastRow -Sheet $Sheet -FirstAppRow $first
    [PSCustomObject]@{ HasHeader=$hasHeader; FirstAppRow=$first; LastAppRow=$last; LastColumn=[Math]::Max(2,(Get-UsedLastCol $Sheet)) }
}

function Invoke-SafeValue {
    param([scriptblock]$Script, $Default = "")
    try {
        $value = & $Script
        if ($null -eq $value) { return $Default }
        return $value
    } catch {
        return $Default
    }
}

function Format-AuditDate {
    param($Value)
    if ($null -eq $Value) { return "" }
    try {
        if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm:ss') }
        if ($Value.ToString() -match '^\d{14}\.') {
            return ([System.Management.ManagementDateTimeConverter]::ToDateTime($Value)).ToString('yyyy-MM-dd HH:mm:ss')
        }
        return $Value.ToString()
    } catch {
        return $Value.ToString()
    }
}

function Format-GB {
    param($Bytes)
    try {
        if ($null -eq $Bytes) { return "" }
        return ("{0:N1} GB" -f ([double]$Bytes / 1GB))
    } catch {
        return ""
    }
}

function Format-List {
    param($Values)
    $items = @($Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.ToString().Trim() })
    if ($items.Count -eq 0) { return "" }
    return ($items -join '; ')
}

function New-AuditRow {
    param(
        [string]$Name,
        [string]$Expected = "",
        [string]$Value = "",
        [string]$Status = "Info",
        [string]$Comment = ""
    )
    [PSCustomObject]@{
        Name     = $Name
        Expected = $Expected
        Value    = $Value
        Status   = $Status
        Comment  = $Comment
    }
}

function Get-ChassisTypeText {
    param($Codes)
    $names = @()
    foreach ($code in @($Codes)) {
        switch ([int]$code) {
            3  { $names += 'Desktop' }
            4  { $names += 'Low Profile Desktop' }
            5  { $names += 'Pizza Box' }
            6  { $names += 'Mini Tower' }
            7  { $names += 'Tower' }
            8  { $names += 'Portable' }
            9  { $names += 'Laptop' }
            10 { $names += 'Notebook' }
            11 { $names += 'Hand Held' }
            12 { $names += 'Docking Station' }
            13 { $names += 'All-in-One' }
            14 { $names += 'Sub Notebook' }
            30 { $names += 'Tablet' }
            31 { $names += 'Convertible' }
            32 { $names += 'Detachable' }
            default { $names += "Chassis code $code" }
        }
    }
    return (Format-List $names)
}

function Test-PendingReboot {
    $paths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired',
        'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'
    )
    try { if (Test-Path $paths[0]) { return $true } } catch {}
    try { if (Test-Path $paths[1]) { return $true } } catch {}
    try {
        $session = Get-ItemProperty -Path $paths[2] -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
        if ($session.PendingFileRenameOperations) { return $true }
    } catch {}
    return $false
}

function Get-DeviceAuditRows {
    $rows = @()
    $cs = Invoke-SafeValue { Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop } $null
    $os = Invoke-SafeValue { Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop } $null
    $bios = Invoke-SafeValue { Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop } $null
    $enclosure = Invoke-SafeValue { Get-CimInstance -ClassName Win32_SystemEnclosure -ErrorAction Stop | Select-Object -First 1 } $null
    $cpu = Invoke-SafeValue { Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop | Select-Object -First 1 } $null
    $gpu = @(Invoke-SafeValue { Get-CimInstance -ClassName Win32_VideoController -ErrorAction Stop } @())
    $battery = Invoke-SafeValue { Get-CimInstance -ClassName Win32_Battery -ErrorAction Stop | Select-Object -First 1 } $null

    $rows += New-AuditRow -Name 'Computer name' -Value (Invoke-SafeValue { $env:COMPUTERNAME }) -Status 'Info'
    $rows += New-AuditRow -Name 'Domain or workgroup' -Value (Invoke-SafeValue { if ($cs.PartOfDomain) { $cs.Domain } else { $cs.Workgroup } }) -Status 'Info'
    $rows += New-AuditRow -Name 'Domain joined' -Expected 'Company policy dependent' -Value (Invoke-SafeValue { $cs.PartOfDomain.ToString() }) -Status 'Info'
    $rows += New-AuditRow -Name 'Manufacturer' -Value (Invoke-SafeValue { $cs.Manufacturer }) -Status 'Info'
    $rows += New-AuditRow -Name 'Model' -Value (Invoke-SafeValue { $cs.Model }) -Status 'Info'
    $rows += New-AuditRow -Name 'Serial number' -Value (Invoke-SafeValue { $bios.SerialNumber }) -Status 'Info'
    $rows += New-AuditRow -Name 'BIOS asset tag' -Value (Invoke-SafeValue { $enclosure.SMBIOSAssetTag }) -Status 'Info'
    $rows += New-AuditRow -Name 'Chassis type' -Value (Invoke-SafeValue { Get-ChassisTypeText $enclosure.ChassisTypes }) -Status 'Info'
    $rows += New-AuditRow -Name 'Operating system' -Value (Invoke-SafeValue { $os.Caption }) -Status 'Info'
    $rows += New-AuditRow -Name 'OS version' -Value (Invoke-SafeValue { $os.Version }) -Status 'Info'
    $rows += New-AuditRow -Name 'OS build number' -Value (Invoke-SafeValue { $os.BuildNumber }) -Status 'Info'
    $rows += New-AuditRow -Name 'OS architecture' -Value (Invoke-SafeValue { $os.OSArchitecture }) -Status 'Info'
    $rows += New-AuditRow -Name 'OS install date' -Value (Invoke-SafeValue { Format-AuditDate $os.InstallDate }) -Status 'Info'
    $lastBoot = Invoke-SafeValue { Format-AuditDate $os.LastBootUpTime }
    $rows += New-AuditRow -Name 'Last boot time' -Value $lastBoot -Status 'Info'
    $uptime = Invoke-SafeValue {
        $boot = if ($os.LastBootUpTime -is [datetime]) { $os.LastBootUpTime } else { [System.Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime) }
        $span = New-TimeSpan -Start $boot -End (Get-Date)
        "{0} days {1} hours" -f [int]$span.TotalDays, $span.Hours
    }
    $rows += New-AuditRow -Name 'Uptime' -Value $uptime -Status 'Info'
    $rows += New-AuditRow -Name 'Time zone' -Value (Invoke-SafeValue { (Get-TimeZone).Id }) -Status 'Info'
    $rows += New-AuditRow -Name 'CPU' -Value (Invoke-SafeValue { $cpu.Name }) -Status 'Info'
    $rows += New-AuditRow -Name 'CPU cores / logical processors' -Value (Invoke-SafeValue { "$($cpu.NumberOfCores) Cores / $($cpu.NumberOfLogicalProcessors) Threads" }) -Status 'Info'
    $rows += New-AuditRow -Name 'Total RAM' -Value (Invoke-SafeValue { Format-GB $cs.TotalPhysicalMemory }) -Status 'Info'

    $diskValues = @()
    $diskPercentages = @()
    foreach ($disk in @(Invoke-SafeValue { Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop } @())) {
        $size = [double]$disk.Size
        $free = [double]$disk.FreeSpace
        $pct = if ($size -gt 0) { [Math]::Round(($free / $size) * 100, 1) } else { 0 }
        $diskPercentages += $pct
        $diskValues += ("{0} {1} free of {2} ({3}%)" -f $disk.DeviceID, (Format-GB $free), (Format-GB $size), $pct)
    }
    $minimumAvailable = if ($diskPercentages.Count -gt 0) { [double]($diskPercentages | Measure-Object -Minimum).Minimum } else { 0 }
    $rows += New-AuditRow -Name '% Space Available' -Value ("{0}%" -f $minimumAvailable) -Status 'Info' -Comment ("Minimum available percentage across fixed disks.`n{0}" -f (Format-List $diskValues))

    $rows += New-AuditRow -Name 'Battery' -Value (Invoke-SafeValue { if ($battery) { "Status $($battery.BatteryStatus), $($battery.EstimatedChargeRemaining)% remaining" } else { 'No battery detected' } }) -Status 'Info'
    $rows += New-AuditRow -Name 'BIOS version' -Value (Invoke-SafeValue { Format-List $bios.SMBIOSBIOSVersion }) -Status 'Info'
    $rows += New-AuditRow -Name 'BIOS release date' -Value (Invoke-SafeValue { Format-AuditDate $bios.ReleaseDate }) -Status 'Info'
    $rows += New-AuditRow -Name 'Graphics adapter' -Value (Format-List (@($gpu | Select-Object -ExpandProperty Name -ErrorAction SilentlyContinue))) -Status 'Info'
    $rows += New-AuditRow -Name 'Physical disks' -Value (Invoke-SafeValue { Format-List (@(Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction Stop | ForEach-Object { "$($_.Model), $(Format-GB $_.Size), status $($_.Status)" })) }) -Status 'Info'
    $rows += New-AuditRow -Name 'Monitors' -Value (Invoke-SafeValue { Format-List (@(Get-CimInstance -ClassName Win32_DesktopMonitor -ErrorAction Stop | Where-Object { $_.Name } | ForEach-Object { "$($_.Name) $($_.ScreenWidth)x$($_.ScreenHeight)" })) }) -Status 'Info'
    $rows += New-AuditRow -Name 'Printers' -Value (Invoke-SafeValue { Format-List (@(Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop | ForEach-Object { if ($_.Default) { "$($_.Name) [default]" } else { $_.Name } })) }) -Status 'Info'
    $rows += New-AuditRow -Name 'Pending reboot' -Expected 'False' -Value ((Test-PendingReboot).ToString()) -Status $(if (Test-PendingReboot) { 'Warning' } else { 'Pass' })
    return $rows
}

function Get-TpmAuditValue {
    try {
        if (Get-Command Get-Tpm -ErrorAction SilentlyContinue) {
            $tpm = Get-Tpm -ErrorAction Stop
            if ($tpm.TpmPresent) { return [PSCustomObject]@{ Value=("Present, ready: {0}" -f $tpm.TpmReady); Status=$(if ($tpm.TpmReady) { 'Pass' } else { 'Warning' }) } }
            return [PSCustomObject]@{ Value='Not present'; Status='Fail' }
        }
    } catch {}
    return [PSCustomObject]@{ Value='Unavailable'; Status='Warning' }
}

function Get-SecureBootAuditValue {
    try {
        if (Get-Command Confirm-SecureBootUEFI -ErrorAction SilentlyContinue) {
            $enabled = Confirm-SecureBootUEFI -ErrorAction Stop
            return [PSCustomObject]@{ Value=($enabled.ToString()); Status=$(if ($enabled) { 'Pass' } else { 'Fail' }) }
        }
    } catch {}
    return [PSCustomObject]@{ Value='Unavailable or legacy BIOS'; Status='Warning' }
}

function Get-BitLockerAuditValue {
    try {
        if (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue) {
            $volume = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
            $value = "Volume $($volume.MountPoint), protection $($volume.ProtectionStatus), encryption $($volume.VolumeStatus)"
            $status = if ($volume.ProtectionStatus.ToString() -eq 'On') { 'Pass' } else { 'Fail' }
            return [PSCustomObject]@{ Value=$value; Status=$status }
        }
    } catch {}
    return [PSCustomObject]@{ Value='Unavailable'; Status='Warning' }
}

function Get-DefenderAuditRows {
    $rows = @()
    try {
        if (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue) {
            $mp = Get-MpComputerStatus -ErrorAction Stop
            $rows += New-AuditRow -Name 'Defender antivirus enabled' -Expected 'True' -Value ($mp.AntivirusEnabled.ToString()) -Status $(if ($mp.AntivirusEnabled) { 'Pass' } else { 'Fail' })
            $rows += New-AuditRow -Name 'Defender real-time protection' -Expected 'True' -Value ($mp.RealTimeProtectionEnabled.ToString()) -Status $(if ($mp.RealTimeProtectionEnabled) { 'Pass' } else { 'Fail' })
            $rows += New-AuditRow -Name 'Defender signatures' -Expected 'Current' -Value ("AV {0}, AS {1}" -f $mp.AntivirusSignatureLastUpdated, $mp.AntispywareSignatureLastUpdated) -Status 'Info'
            return $rows
        }
    } catch {}
    $rows += New-AuditRow -Name 'Defender status' -Expected 'Available' -Value 'Unavailable' -Status 'Warning'
    return $rows
}

function Get-SecurityAuditRows {
    $rows = @()
    $tpm = Get-TpmAuditValue
    $secureBoot = Get-SecureBootAuditValue
    $bitLocker = Get-BitLockerAuditValue
    $rows += New-AuditRow -Name 'TPM' -Expected 'Present and ready' -Value ($tpm.Value) -Status ($tpm.Status)
    $rows += New-AuditRow -Name 'Secure Boot' -Expected 'True' -Value ($secureBoot.Value) -Status ($secureBoot.Status)
    $rows += New-AuditRow -Name 'BitLocker system drive' -Expected 'Protection On' -Value ($bitLocker.Value) -Status ($bitLocker.Status)
    $rows += Get-DefenderAuditRows

    try {
        if (Get-Command Get-NetFirewallProfile -ErrorAction SilentlyContinue) {
            $profiles = @(Get-NetFirewallProfile -ErrorAction Stop)
            $disabled = @($profiles | Where-Object { -not $_.Enabled })
            $rows += New-AuditRow -Name 'Windows Firewall profiles' -Expected 'All enabled' -Value (Format-List (@($profiles | ForEach-Object { "$($_.Name): $($_.Enabled)" }))) -Status $(if ($disabled.Count -eq 0) { 'Pass' } else { 'Fail' })
        } else {
            $rows += New-AuditRow -Name 'Windows Firewall profiles' -Expected 'All enabled' -Value 'Unavailable' -Status 'Warning'
        }
    } catch {
        $rows += New-AuditRow -Name 'Windows Firewall profiles' -Expected 'All enabled' -Value 'Unavailable' -Status 'Warning'
    }

    try {
        $uac = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -ErrorAction Stop
        $enabled = ($uac.EnableLUA -eq 1)
        $rows += New-AuditRow -Name 'UAC enabled' -Expected 'True' -Value ($enabled.ToString()) -Status $(if ($enabled) { 'Pass' } else { 'Fail' })
    } catch {
        $rows += New-AuditRow -Name 'UAC enabled' -Expected 'True' -Value 'Unavailable' -Status 'Warning'
    }

    try {
        $rdp = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections -ErrorAction Stop
        $enabled = ($rdp.fDenyTSConnections -eq 0)
        $rows += New-AuditRow -Name 'Remote Desktop enabled' -Expected 'Company policy dependent' -Value ($enabled.ToString()) -Status $(if ($enabled) { 'Warning' } else { 'Pass' })
    } catch {
        $rows += New-AuditRow -Name 'Remote Desktop enabled' -Expected 'Company policy dependent' -Value 'Unavailable' -Status 'Warning'
    }

    try {
        $smb1Enabled = $false
        if (Get-Command Get-SmbServerConfiguration -ErrorAction SilentlyContinue) {
            $smb = Get-SmbServerConfiguration -ErrorAction Stop
            $smb1Enabled = [bool]$smb.EnableSMB1Protocol
            $rows += New-AuditRow -Name 'SMBv1 enabled' -Expected 'False' -Value ($smb1Enabled.ToString()) -Status $(if ($smb1Enabled) { 'Fail' } else { 'Pass' })
        } else {
            $rows += New-AuditRow -Name 'SMBv1 enabled' -Expected 'False' -Value 'Unavailable' -Status 'Warning'
        }
    } catch {
        $rows += New-AuditRow -Name 'SMBv1 enabled' -Expected 'False' -Value 'Unavailable' -Status 'Warning'
    }

    try {
        $av = @(Get-CimInstance -Namespace 'root\SecurityCenter2' -ClassName AntiVirusProduct -ErrorAction Stop)
        $rows += New-AuditRow -Name 'Security Center antivirus products' -Expected 'At least one active product' -Value (Format-List (@($av | Select-Object -ExpandProperty displayName -ErrorAction SilentlyContinue))) -Status $(if ($av.Count -gt 0) { 'Pass' } else { 'Warning' })
    } catch {
        $rows += New-AuditRow -Name 'Security Center antivirus products' -Expected 'At least one active product' -Value 'Unavailable' -Status 'Warning'
    }
    try {
        $soon = (Get-Date).AddDays(45)
        $certs = @(Get-ChildItem -Path Cert:\LocalMachine\My -ErrorAction Stop | Where-Object { $_.NotAfter -le $soon } | ForEach-Object { "$($_.Subject), expires $($_.NotAfter.ToString('yyyy-MM-dd'))" })
        $rows += New-AuditRow -Name 'Machine certificates expiring within 45 days' -Expected 'None' -Value (Format-List $certs) -Status $(if ($certs.Count -gt 0) { 'Warning' } else { 'Pass' })
    } catch {
        $rows += New-AuditRow -Name 'Machine certificates expiring within 45 days' -Expected 'None' -Value 'Unavailable' -Status 'Warning'
    }
    return $rows
}

function Get-NetworkAuditRows {
    $rows = @()
    $rows += New-AuditRow -Name 'DNS host entry' -Value (Invoke-SafeValue { ([System.Net.Dns]::GetHostEntry([System.Net.Dns]::GetHostName()).HostName) }) -Status 'Info'

    $configs = @(Invoke-SafeValue { Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" -ErrorAction Stop } @())
    $rows += New-AuditRow -Name 'Active IP addresses' -Value (Format-List (@($configs | ForEach-Object { $_.IPAddress }))) -Status 'Info'
    $rows += New-AuditRow -Name 'MAC addresses' -Value (Format-List (@($configs | ForEach-Object { "$($_.Description): $($_.MACAddress)" }))) -Status 'Info'
    $rows += New-AuditRow -Name 'Default gateways' -Value (Format-List (@($configs | ForEach-Object { $_.DefaultIPGateway }))) -Status 'Info'
    $rows += New-AuditRow -Name 'DNS servers' -Value (Format-List (@($configs | ForEach-Object { $_.DNSServerSearchOrder }))) -Status 'Info'
    $rows += New-AuditRow -Name 'DHCP enabled adapters' -Value (Format-List (@($configs | Where-Object { $_.DHCPEnabled } | ForEach-Object { $_.Description }))) -Status 'Info'
    $rows += New-AuditRow -Name 'Mapped network drives' -Value (Invoke-SafeValue { Format-List (@(Get-CimInstance -ClassName Win32_MappedLogicalDisk -ErrorAction Stop | ForEach-Object { "$($_.DeviceID) -> $($_.ProviderName)" })) }) -Status 'Info'

    try {
        if (Get-Command Get-NetConnectionProfile -ErrorAction SilentlyContinue) {
            $profiles = @(Get-NetConnectionProfile -ErrorAction Stop)
            $rows += New-AuditRow -Name 'Network profiles' -Expected 'Domain where appropriate' -Value (Format-List (@($profiles | ForEach-Object { "$($_.Name): $($_.NetworkCategory)" }))) -Status 'Info'
        }
    } catch {}

    $wifi = Invoke-SafeValue {
        $ssid = ""
        $netsh = netsh wlan show interfaces 2>$null
        foreach ($line in $netsh) {
            if ($line -match '^\s*SSID\s*:\s*(.+)$' -and $line -notmatch 'BSSID') { $ssid = $matches[1].Trim(); break }
        }
        $ssid
    }
    $rows += New-AuditRow -Name 'Wi-Fi SSID' -Expected 'Company policy dependent' -Value $wifi -Status 'Info'
    return $rows
}

function Get-UserAuditRows {
    $rows = @()
    $rows += New-AuditRow -Name 'Current user' -Value (Invoke-SafeValue { [System.Security.Principal.WindowsIdentity]::GetCurrent().Name }) -Status 'Info'
    $rows += New-AuditRow -Name 'Last logged-on user' -Value (Invoke-SafeValue { (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI' -Name LastLoggedOnUser -ErrorAction Stop).LastLoggedOnUser }) -Status 'Info'

    try {
        if (Get-Command Get-LocalUser -ErrorAction SilentlyContinue) {
            $users = @(Get-LocalUser -ErrorAction Stop)
            $rows += New-AuditRow -Name 'Enabled local users' -Expected 'Review exceptions' -Value (Format-List (@($users | Where-Object { $_.Enabled } | Select-Object -ExpandProperty Name))) -Status 'Warning'
            $rows += New-AuditRow -Name 'Disabled local users' -Value (Format-List (@($users | Where-Object { -not $_.Enabled } | Select-Object -ExpandProperty Name))) -Status 'Info'
        } else {
            $rows += New-AuditRow -Name 'Local users' -Value 'Unavailable' -Status 'Warning'
        }
    } catch {
        $rows += New-AuditRow -Name 'Local users' -Value 'Unavailable' -Status 'Warning'
    }

    try {
        if (Get-Command Get-LocalGroupMember -ErrorAction SilentlyContinue) {
            $admins = @(Get-LocalGroupMember -Group 'Administrators' -ErrorAction Stop)
            $rows += New-AuditRow -Name 'Local Administrators members' -Expected 'Review exceptions' -Value (Format-List (@($admins | ForEach-Object { "$($_.Name) [$($_.ObjectClass)]" }))) -Status 'Warning'
        } else {
            $rows += New-AuditRow -Name 'Local Administrators members' -Expected 'Review exceptions' -Value 'Unavailable' -Status 'Warning'
        }
    } catch {
        $rows += New-AuditRow -Name 'Local Administrators members' -Expected 'Review exceptions' -Value 'Unavailable' -Status 'Warning'
    }

    try {
        $profiles = @(Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop | Where-Object { -not $_.Special })
        $profileText = @($profiles | Sort-Object LocalPath | ForEach-Object { "$($_.LocalPath), last use $(Format-AuditDate $_.LastUseTime)" })
        $rows += New-AuditRow -Name 'User profiles' -Expected 'Review stale profiles' -Value (Format-List $profileText) -Status 'Info'
    } catch {
        $rows += New-AuditRow -Name 'User profiles' -Expected 'Review stale profiles' -Value 'Unavailable' -Status 'Warning'
    }
    return $rows
}

function Get-ServiceAuditRows {
    $rows = @()
    $checks = @(
        @{ Name='WinDefend'; Display='Microsoft Defender Antivirus'; Required=$true },
        @{ Name='MpsSvc'; Display='Windows Firewall'; Required=$true },
        @{ Name='wuauserv'; Display='Windows Update'; Required=$true },
        @{ Name='BITS'; Display='Background Intelligent Transfer Service'; Required=$true },
        @{ Name='IntuneManagementExtension'; Display='Microsoft Intune Management Extension'; Required=$false },
        @{ Name='CcmExec'; Display='Microsoft Configuration Manager Client'; Required=$false },
        @{ Name='Sense'; Display='Microsoft Defender for Endpoint Sensor'; Required=$false }
    ) + $RequiredCompanyServices

    foreach ($check in $checks) {
        $svc = $null
        try { $svc = Get-Service -Name ($check.Name) -ErrorAction SilentlyContinue } catch {}
        $expected = if ($check.Required) { 'Installed and running' } else { 'Company policy dependent' }
        if ($svc) {
            $status = if ($svc.Status -eq 'Running') { 'Pass' } elseif ($check.Required) { 'Fail' } else { 'Warning' }
            $rows += New-AuditRow -Name ($check.Display) -Expected $expected -Value ("{0} ({1})" -f $svc.Status, $check.Name) -Status $status
        } else {
            $status = if ($check.Required) { 'Fail' } else { 'Info' }
            $rows += New-AuditRow -Name ($check.Display) -Expected $expected -Value ("Not installed ({0})" -f $check.Name) -Status $status
        }
    }
    return $rows
}

function Get-AuditRowByName {
    param([array]$Rows, [string]$Name)
    $matches = @($Rows | Where-Object { $_.Name -eq $Name } | Select-Object -First 1)
    if ($matches.Count -gt 0) { return $matches[0] }
    return $null
}

function Get-ComplianceAuditRows {
    param([array]$DeviceRows, [array]$SecurityRows, [array]$ServiceRows, [array]$UserRows = @())
    $rows = @()
    foreach ($name in @('Pending reboot', '% Space Available')) {
        $row = Get-AuditRowByName -Rows $DeviceRows -Name $name
        if ($row) { $rows += New-AuditRow -Name $name -Expected ($row.Expected) -Value ($row.Value) -Status ($row.Status) -Comment ($row.Comment) }
    }
    foreach ($name in @('TPM', 'Secure Boot', 'BitLocker system drive', 'Defender antivirus enabled', 'Defender real-time protection', 'Windows Firewall profiles', 'UAC enabled', 'SMBv1 enabled', 'Security Center antivirus products')) {
        $row = Get-AuditRowByName -Rows $SecurityRows -Name $name
        if ($row) { $rows += New-AuditRow -Name $name -Expected ($row.Expected) -Value ($row.Value) -Status ($row.Status) -Comment ($row.Comment) }
    }
    foreach ($row in @($ServiceRows | Where-Object { $_.Expected -eq 'Installed and running' -or $_.Status -in @('Fail','Warning') })) {
        $rows += New-AuditRow -Name ($row.Name) -Expected ($row.Expected) -Value ($row.Value) -Status ($row.Status) -Comment ($row.Comment)
    }
    foreach ($name in @('Enabled local users', 'Disabled local users', 'Local Administrators members')) {
        $row = Get-AuditRowByName -Rows $UserRows -Name $name
        if ($row) { $rows += New-AuditRow -Name $name -Expected ($row.Expected) -Value ($row.Value) -Status ($row.Status) -Comment ($row.Comment) }
    }
    $failCount = @($rows | Where-Object { $_.Status -eq 'Fail' }).Count
    $warningCount = @($rows | Where-Object { $_.Status -eq 'Warning' }).Count
    $overallStatus = if ($failCount -gt 0) { 'Fail' } elseif ($warningCount -gt 0) { 'Warning' } else { 'Pass' }
    $rows = @(New-AuditRow -Name 'Overall compliance summary' -Expected 'No failures or warnings' -Value ("Failures: {0}; Warnings: {1}" -f $failCount, $warningCount) -Status $overallStatus) + $rows
    return $rows
}

function Get-RoomChoice {
    do {
        $room = (Read-Host "Enter room / site / department name for these audit sheets").Trim()
        if ([string]::IsNullOrWhiteSpace($room)) { Write-Host "Room name cannot be blank." -ForegroundColor Yellow }
    } until (-not [string]::IsNullOrWhiteSpace($room))
    return $room
}

function Get-YesNoChoice {
    param([string]$Prompt, [bool]$Default = $false)
    $suffix = if ($Default) { '[Y/n]' } else { '[y/N]' }
    do {
        $raw = Read-Host "$Prompt $suffix"
        if ([string]::IsNullOrWhiteSpace($raw)) { return $Default }
        $value = $raw.Trim().ToLowerInvariant()
        if ($value.StartsWith('y')) { return $true }
        if ($value.StartsWith('n')) { return $false }
    } until ($false)
}

function Get-SoftwareAuditModeChoice {
    Write-Host "Software audit detail:" -ForegroundColor Cyan
    Write-Host "  1. Simple - required software comparison only"
    Write-Host "  2. In-depth - required software comparison plus full installed software inventory"
    do { $choice = Read-Host "Choose 1 or 2" } until ($choice -in @('1','2'))
    if ($choice -eq '2') { return 'InDepth' }
    return 'Simple'
}

function Get-SafeWorksheetName {
    param([string]$Room, [string]$AuditType)
    $safeRoom = ($Room -replace '[\\/\?\*\[\]:]', ' ').Trim()
    $safeType = ($AuditType -replace '[\\/\?\*\[\]:]', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($safeRoom)) { $safeRoom = 'Room' }
    $maxRoomLength = [Math]::Max(1, 31 - $safeType.Length - 1)
    if ($safeRoom.Length -gt $maxRoomLength) { $safeRoom = $safeRoom.Substring(0, $maxRoomLength).Trim() }
    $name = "$safeRoom $safeType".Trim()
    if ($name.Length -gt 31) { $name = $name.Substring(0,31).Trim() }
    return $name
}

function Get-UniqueWorksheetName {
    param($Workbook, [string]$BaseName)
    $safeBase = ($BaseName -replace '[\\/\?\*\[\]:]', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($safeBase)) { $safeBase = 'Audit' }
    if ($safeBase.Length -gt 31) { $safeBase = $safeBase.Substring(0,31).Trim() }
    if (-not (Get-WorksheetByName -Workbook $Workbook -Name $safeBase)) { return $safeBase }

    for ($i=2; $i -le 99; $i++) {
        $suffix = " $i"
        $baseLength = 31 - $suffix.Length
        $candidate = ($safeBase.Substring(0, [Math]::Min($safeBase.Length, $baseLength)).Trim() + $suffix).Trim()
        if (-not (Get-WorksheetByName -Workbook $Workbook -Name $candidate)) { return $candidate }
    }
    throw "Could not create a unique worksheet name for '$BaseName'."
}

function Get-AuditWorksheetNameMap {
    param(
        $Workbook,
        [string]$Room,
        [array]$AuditTypes,
        [bool]$StartNewSheets,
        [string]$RunLabel
    )
    $map = @{}
    $reservedNames = @{}
    foreach ($type in $AuditTypes) {
        $baseType = if ($StartNewSheets) { "$type $RunLabel" } else { $type }
        $baseName = Get-SafeWorksheetName -Room $Room -AuditType $baseType
        if ($StartNewSheets) {
            $candidate = Get-UniqueWorksheetName -Workbook $Workbook -BaseName $baseName
            $counter = 2
            while ($reservedNames.ContainsKey($candidate)) {
                $suffix = " $counter"
                $baseLength = 31 - $suffix.Length
                $candidate = ($baseName.Substring(0, [Math]::Min($baseName.Length, $baseLength)).Trim() + $suffix).Trim()
                if (Get-WorksheetByName -Workbook $Workbook -Name $candidate) {
                    $candidate = Get-UniqueWorksheetName -Workbook $Workbook -BaseName $candidate
                }
                $counter++
            }
            $map[$type] = $candidate
        } else {
            $map[$type] = $baseName
        }
        $reservedNames[$map[$type]] = $true
    }
    return $map
}

function Get-WorksheetByName {
    param($Workbook, [string]$Name)
    for ($i=1; $i -le $Workbook.Worksheets.Count; $i++) {
        $sheet = $Workbook.Worksheets.Item($i)
        if ($sheet.Name -eq $Name) { return $sheet }
    }
    return $null
}

function Get-OrCreateWorksheet {
    param($Workbook, [string]$Name)
    $sheet = Get-WorksheetByName -Workbook $Workbook -Name $Name
    if ($sheet) { return $sheet }
    $sheet = $Workbook.Worksheets.Add()
    $sheet.Name = $Name
    return $sheet
}

function Ensure-AuditSheetHeader {
    param($Sheet, [string]$FirstHeader, [string]$SecondHeader)
    try {
        $Sheet.Cells.Item(1,1).Value2 = $FirstHeader
        $Sheet.Cells.Item(1,2).Value2 = $SecondHeader
        $header = $Sheet.Range($Sheet.Cells.Item(1,1), $Sheet.Cells.Item(1,2))
        $header.Font.Bold = $true
        Set-FillColor -RangeOrCell $header -Color $OriginalBlueColor
        $header.Borders.LineStyle = 1
    } catch {}
}

function Get-AuditDisplayValue {
    param([string]$Expected, [string]$Value, [string]$Status)
    if ([string]::IsNullOrWhiteSpace($Value)) {
        if ($Status -in @('Fail','Missing')) { return '(not detected)' }
        return $Expected
    }
    return $Value
}

function Set-ColumnPixelWidth {
    param($Sheet, [int]$ColumnNumber, [double]$Pixels)
    try {
        $column = $Sheet.Columns.Item($ColumnNumber)
        $column.ColumnWidth = 10
        $targetPoints = $Pixels * 72.0 / 96.0
        for ($attempt = 0; $attempt -lt 4; $attempt++) {
            $currentPoints = [double]$column.Width
            if ($currentPoints -le 0) { break }
            $column.ColumnWidth = [Math]::Max(0.1, [double]$column.ColumnWidth * ($targetPoints / $currentPoints))
        }
    } catch {}
}

function Set-MaximumColumnPixelWidth {
    param($Sheet, [int]$ColumnNumber, [double]$MaximumPixels = 45)
    try {
        $column = $Sheet.Columns.Item($ColumnNumber)
        $column.AutoFit() | Out-Null
        $maximumPoints = $MaximumPixels * 72.0 / 96.0
        for ($attempt = 0; $attempt -lt 3; $attempt++) {
            $currentPoints = [double]$column.Width
            if ($currentPoints -le $maximumPoints) { break }
            $currentColumnWidth = [double]$column.ColumnWidth
            $column.ColumnWidth = [Math]::Max(0.1, $currentColumnWidth * ($maximumPoints / $currentPoints))
        }
    } catch {}
}

function Set-BaselineValidation {
    param($Sheet)
    $lastRow = [Math]::Max(1, (Get-UsedLastRow $Sheet))
    $lastCol = [Math]::Max(2, (Get-UsedLastCol $Sheet))
    for ($r = 2; $r -le $lastRow; $r++) {
        $cell = $Sheet.Cells.Item($r,2)
        try { $cell.Validation.Delete() } catch {}
        if ($Sheet.Cells.Item($r,1).Text.Trim() -eq 'Additional Software') { continue }
        $values = @($cell.Text.Trim())
        for ($c = 3; $c -le $lastCol; $c++) { $values += $Sheet.Cells.Item($r,$c).Text.Trim() }
        $normalized = @($values | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.ToLowerInvariant() })
        $options = ''
        if (@($normalized | Where-Object { $_ -in @('true','false') }).Count -gt 0) { $options = 'True,False' }
        elseif (@($normalized | Where-Object { $_ -in @('yes','no') }).Count -gt 0) { $options = 'Yes,No' }
        if ([string]::IsNullOrWhiteSpace($options)) { continue }
        try {
            $cell.Validation.Add(3, 1, 1, $options) | Out-Null
            $cell.Validation.IgnoreBlank = $true
            $cell.Validation.InCellDropdown = $true
            $cell.Validation.ShowError = $true
            $cell.Validation.ErrorTitle = 'Choose a valid baseline value'
            $cell.Validation.ErrorMessage = 'Select one of the values in the dropdown.'
        } catch {}
    }
}

function Set-DynamicAuditConditionalFormatting {
    param($Sheet)
    $lastRow = [Math]::Max(1, (Get-UsedLastRow $Sheet))
    $lastCol = [Math]::Max(2, (Get-UsedLastCol $Sheet))
    if ($lastRow -lt 2 -or $lastCol -lt 3) { return }
    try {
        $range = $Sheet.Range($Sheet.Cells.Item(2,3), $Sheet.Cells.Item($lastRow,$lastCol))
        $range.FormatConditions.Delete()
        $redFormula = '=OR(LOWER(TRIM(C2))="(not detected)",LOWER(TRIM(C2))="unavailable",LOWER(TRIM(C2))="missing")'
        $greenFormula = '=AND(C2<>"",IF($A2="% Space Available",IFERROR(VALUE(SUBSTITUTE(C2,"%",""))>=VALUE(SUBSTITUTE($B2,"%","")),FALSE),LOWER(TRIM(C2))=LOWER(TRIM($B2))))'
        $yellowFormula = '=AND(C2<>"",NOT(OR(LOWER(TRIM(C2))="(not detected)",LOWER(TRIM(C2))="unavailable",LOWER(TRIM(C2))="missing")),IF($A2="% Space Available",IFERROR(VALUE(SUBSTITUTE(C2,"%",""))<VALUE(SUBSTITUTE($B2,"%","")),TRUE),LOWER(TRIM(C2))<>LOWER(TRIM($B2))))'
        $redRule = $range.FormatConditions.Add(2, $null, $redFormula)
        $redRule.Interior.Color = $OriginalRedColor
        $redRule.StopIfTrue = $true
        $greenRule = $range.FormatConditions.Add(2, $null, $greenFormula)
        $greenRule.Interior.Color = $OriginalGreenColor
        $greenRule.StopIfTrue = $true
        $yellowRule = $range.FormatConditions.Add(2, $null, $yellowFormula)
        $yellowRule.Interior.Color = $OriginalYellowColor
        $yellowRule.StopIfTrue = $true
    } catch {
        Write-TroubleshootingLog -Level 'WARN' -Category 'Excel' -Message ("Could not apply dynamic conditional formatting to '{0}': {1}" -f $Sheet.Name, $_.Exception.Message)
    }
}

function Restore-MatchingGreenValues {
    param($Sheet)
    $lastRow = [Math]::Max(1, (Get-UsedLastRow $Sheet))
    $lastCol = [Math]::Max(2, (Get-UsedLastCol $Sheet))
    for ($r = 2; $r -le $lastRow; $r++) {
        $baseline = $Sheet.Cells.Item($r,2).Text.Trim()
        if ([string]::IsNullOrWhiteSpace($baseline) -or $Sheet.Cells.Item($r,1).Text.Trim() -eq 'Additional Software') { continue }
        for ($c = 3; $c -le $lastCol; $c++) {
            $cell = $Sheet.Cells.Item($r,$c)
            if (-not [string]::IsNullOrWhiteSpace($cell.Text)) { continue }
            try {
                if ([int]$cell.Interior.Color -eq $OriginalGreenColor) {
                    $cell.NumberFormat = '@'
                    $cell.Value2 = [string]$baseline
                }
            } catch {}
        }
    }
}

function Remove-DuplicateAuditRows {
    param($Sheet)
    $firstRowByName = @{}
    $duplicates = @()
    $lastRow = [Math]::Max(1, (Get-UsedLastRow $Sheet))
    $lastCol = [Math]::Max(2, (Get-UsedLastCol $Sheet))

    for ($r = 2; $r -le $lastRow; $r++) {
        $name = $Sheet.Cells.Item($r,1).Text.Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $key = ($name -replace '\s+', ' ').Trim()
        if ($firstRowByName.ContainsKey($key)) {
            $duplicates += [PSCustomObject]@{ Row = $r; FirstRow = $firstRowByName[$key] }
        } else {
            $firstRowByName[$key] = $r
        }
    }

    foreach ($duplicate in @($duplicates | Sort-Object Row -Descending)) {
        for ($c = 2; $c -le $lastCol; $c++) {
            $source = $Sheet.Cells.Item($duplicate.Row,$c)
            $destination = $Sheet.Cells.Item($duplicate.FirstRow,$c)
            if ([string]::IsNullOrWhiteSpace($destination.Text) -and -not [string]::IsNullOrWhiteSpace($source.Text)) {
                try { $source.Copy($destination) | Out-Null } catch {}
            }
        }
        try { $Sheet.Rows.Item($duplicate.Row).Delete() | Out-Null } catch {}
    }
    return $duplicates.Count
}

function Remove-AuditRowsByName {
    param($Sheet, [string[]]$Names)
    if ($null -eq $Sheet -or $Names.Count -eq 0) { return 0 }
    $removed = 0
    $lastRow = [Math]::Max(1, (Get-UsedLastRow $Sheet))
    for ($r = $lastRow; $r -ge 2; $r--) {
        $name = $Sheet.Cells.Item($r,1).Text.Trim()
        if ($Names -contains $name) {
            try {
                $Sheet.Rows.Item($r).Delete() | Out-Null
                $removed++
            } catch {}
        }
    }
    return $removed
}

function Remove-RedundantWorkbookAuditRows {
    param($Workbook)
    $removed = 0
    foreach ($sheetName in @('Software','Device','Security','Network','Services','Compliance')) {
        $sheet = Get-WorksheetByName -Workbook $Workbook -Name $sheetName
        if ($sheet) { $removed += Remove-DuplicateAuditRows -Sheet $sheet }
    }
    $removed += Remove-AuditRowsByName -Sheet (Get-WorksheetByName -Workbook $Workbook -Name 'Device') -Names @('Current user')
    $removed += Remove-AuditRowsByName -Sheet (Get-WorksheetByName -Workbook $Workbook -Name 'Network') -Names @('Hostname')
    return $removed
}

function Remove-RedundantMatchingValues {
    param($Sheet)
    # Matching values intentionally remain visible. Conditional formatting now
    # reacts to baseline edits without deleting the audited evidence.
}

function Format-AuditSheetColumns {
    param($Sheet)
    try {
        $Sheet.Columns.Item(1).AutoFit() | Out-Null
        Set-ColumnPixelWidth -Sheet $Sheet -ColumnNumber 2 -Pixels $BaselineColumnPixels
        $lastCol = [Math]::Max(2, (Get-UsedLastCol $Sheet))
        for ($c = 3; $c -le $lastCol; $c++) {
            Set-ColumnPixelWidth -Sheet $Sheet -ColumnNumber $c -Pixels $AuditColumnPixels
        }
        $used = $Sheet.UsedRange
        $used.HorizontalAlignment = -4131
        $used.VerticalAlignment = -4160
        # Keep every audit entry visually on one line. Longer values remain
        # available in the formula bar and supporting details remain in comments.
        $used.WrapText = $false
        $used.Rows.RowHeight = $AuditRowHeightPoints
        try {
            $used.Borders.LineStyle = 1
            $used.Borders.Color = $GridBorderColor
            $used.Borders.Weight = 2
        } catch {}
        Restore-MatchingGreenValues -Sheet $Sheet
        Set-BaselineValidation -Sheet $Sheet
        Set-DynamicAuditConditionalFormatting -Sheet $Sheet
    } catch {}
}

function Get-AuditAssetColumn {
    param($Sheet, [string]$Asset)
    $lastCol = [Math]::Max(2, (Get-UsedLastCol $Sheet))
    for ($c=3; $c -le $lastCol; $c++) {
        if ($Sheet.Cells.Item(1,$c).Text.Trim() -eq $Asset) { return $c }
    }
    $assetCol = $lastCol + 1
    $Sheet.Cells.Item(1,$assetCol).Value2 = [string]$Asset
    $Sheet.Cells.Item(1,$assetCol).Font.Bold = $true
    Set-ColumnPixelWidth -Sheet $Sheet -ColumnNumber $assetCol -Pixels $AuditColumnPixels
    return $assetCol
}

function Write-AuditRowsToSheet {
    param(
        $Sheet,
        [array]$Rows,
        [string]$Asset,
        [string]$FirstHeader = 'Audit Item',
        [string]$SecondHeader = 'Expected / Baseline'
    )
    Ensure-AuditSheetHeader -Sheet $Sheet -FirstHeader $FirstHeader -SecondHeader $SecondHeader
    $assetCol = Get-AuditAssetColumn -Sheet $Sheet -Asset $Asset
    Remove-DuplicateAuditRows -Sheet $Sheet | Out-Null
    $rowMap = @{}
    $lastRow = [Math]::Max(1, (Get-UsedLastRow $Sheet))
    for ($r=2; $r -le $lastRow; $r++) {
        $name = $Sheet.Cells.Item($r,1).Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($name) -and -not $rowMap.ContainsKey($name)) { $rowMap[$name] = $r }
    }
    foreach ($row in $Rows) {
        if ([string]::IsNullOrWhiteSpace($row.Name)) { continue }
        if ($rowMap.ContainsKey($row.Name)) {
            $r = $rowMap[$row.Name]
        } else {
            $lastRow++
            $r = $lastRow
            $rowMap[$row.Name] = $r
            $Sheet.Cells.Item($r,1).Value2 = [string]$row.Name
        }
        $Sheet.Cells.Item($r,2).NumberFormat = '@'
        $Sheet.Cells.Item($r,2).Value2 = [string]$row.Expected
        Set-FillColor -RangeOrCell ($Sheet.Range($Sheet.Cells.Item($r,1), $Sheet.Cells.Item($r,2))) -Color $OriginalBlueColor
        $displayValue = Get-AuditDisplayValue -Expected $row.Expected -Value $row.Value -Status $row.Status
        Set-AuditCellResult -Cell ($Sheet.Cells.Item($r,$assetCol)) -Value $displayValue -Status $row.Status -Comment $row.Comment
    }
    Remove-RedundantMatchingValues -Sheet $Sheet
    Format-AuditSheetColumns -Sheet $Sheet
}

function Set-AuditCellResult {
    param($Cell, [string]$Value, [string]$Status, [string]$Comment)
    try {
        $Cell.NumberFormat = '@'
        $Cell.Value2 = [string]$Value
        $Cell.HorizontalAlignment = -4131
    } catch {}
    $color = switch ($Status) {
        'Fail' { $OriginalRedColor; break }
        'Missing' { $OriginalRedColor; break }
        'Warning' { $OriginalYellowColor; break }
        'Different' { $OriginalYellowColor; break }
        default { $OriginalGreenColor }
    }
    Set-FillColor -RangeOrCell $Cell -Color $color
    Set-CellComment -Cell $Cell -Text $Comment
}

function Get-SoftwareReferenceRows {
    param([array]$Sheets)
    $rows = @()
    foreach ($sheet in $Sheets) {
        $layout = Get-SheetLayout $sheet
        if ($layout.LastAppRow -lt $layout.FirstAppRow) { continue }
        for ($r=$layout.FirstAppRow; $r -le $layout.LastAppRow; $r++) {
            $name = $sheet.Cells.Item($r,1).Text.Trim()
            $expected = $sheet.Cells.Item($r,2).Text.Trim()
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            if (($r -eq 1) -and ($name -match '^(Audit Item|Software)$')) { continue }
            $rows += [PSCustomObject]@{ Name=$name; Expected=$expected; Normalized=(Normalize-Name $name) }
        }
    }
    $uniqueRows = @()
    $seen = @{}
    foreach ($row in @($rows | Where-Object { $_.Normalized })) {
        if ($seen.ContainsKey($row.Normalized)) { continue }
        $seen[$row.Normalized] = $true
        $uniqueRows += $row
    }
    return $uniqueRows
}

function Get-SoftwareAuditRows {
    param([array]$ReferenceRows, [array]$InstalledApps, [string]$Timestamp)
    $rows = @()
    foreach ($ref in $ReferenceRows) {
        $match = Find-AppMatch -RequiredName $ref.Name -InstalledApps $InstalledApps
        if ($match) {
            $version = Normalize-VersionText $match.Version
            if ([string]::IsNullOrWhiteSpace($version)) { $version = 'Installed' }
            if (Test-VersionMatch -Expected $ref.Expected -Installed $version) {
                $rows += New-AuditRow -Name ($ref.Name) -Expected ($ref.Expected) -Value '' -Status 'Pass' -Comment "Detected as: $($match.Name)`nInstalled version: $version`nSource: $($match.Source)`nAudit time: $Timestamp"
            } else {
                $expectedLabel = if ([string]::IsNullOrWhiteSpace($ref.Expected)) { '(blank)' } else { $ref.Expected }
                $rows += New-AuditRow -Name ($ref.Name) -Expected ($ref.Expected) -Value $version -Status 'Warning' -Comment "Detected as: $($match.Name)`nInstalled version: $version`nExpected version: $expectedLabel`nSource: $($match.Source)`nAudit time: $Timestamp"
            }
        } else {
            $rows += New-AuditRow -Name ($ref.Name) -Expected ($ref.Expected) -Value '' -Status 'Fail' -Comment "No matching Installed Apps entry found for '$($ref.Name)'. Audit time: $Timestamp"
        }
    }
    return $rows
}

function Get-SoftwareInventoryRows {
    param([array]$InstalledApps, [string]$Timestamp)
    $rows = @()
    foreach ($app in @($InstalledApps | Sort-Object Name, Version)) {
        if ($null -eq $app -or [string]::IsNullOrWhiteSpace($app.Name)) { continue }
        $version = Normalize-VersionText $app.Version
        if ([string]::IsNullOrWhiteSpace($version)) { $version = 'Installed - version unavailable' }
        $publisher = if ([string]::IsNullOrWhiteSpace($app.Publisher)) { '(publisher unavailable)' } else { $app.Publisher }
        $source = if ([string]::IsNullOrWhiteSpace($app.Source)) { '(source unavailable)' } else { $app.Source }
        $id = if ([string]::IsNullOrWhiteSpace($app.Id)) { '(id unavailable)' } else { $app.Id }
        $rows += New-AuditRow `
            -Name ($app.Name) `
            -Expected ("{0} / {1}" -f $publisher, $source) `
            -Value $version `
            -Status 'Info' `
            -Comment "Publisher: $publisher`nSource: $source`nId: $id`nAudit time: $Timestamp"
    }
    return $rows
}

function Get-SimpleInstalledSoftwareRows {
    param([array]$InstalledApps, [string]$Timestamp)
    $rows = @(
        New-AuditRow `
            -Name 'Required software baseline' `
            -Expected 'Software names in column A, expected versions in column B' `
            -Value 'No baseline rows found; generated from installed software inventory.' `
            -Status 'Warning' `
            -Comment "The selected reference worksheet was blank. Add required software rows later to enable pass/missing/version comparison. Audit time: $Timestamp"
    )
    foreach ($app in @($InstalledApps | Sort-Object Name, Version)) {
        if ($null -eq $app -or [string]::IsNullOrWhiteSpace($app.Name)) { continue }
        $version = Normalize-VersionText $app.Version
        if ([string]::IsNullOrWhiteSpace($version)) { $version = 'Installed - version unavailable' }
        $rows += New-AuditRow `
            -Name ($app.Name) `
            -Expected 'Installed software inventory' `
            -Value $version `
            -Status 'Info' `
            -Comment "Detected installed app.`nSource: $($app.Source)`nAudit time: $Timestamp"
    }
    return $rows
}

function Ensure-AuditHeader {
    param($Sheet)
    $layout = Get-SheetLayout $Sheet
    if ($layout.HasHeader) { return (Get-SheetLayout $Sheet) }
    Write-Host "Worksheet '$($Sheet.Name)' has no header. Inserting one so the asset column can be labeled." -ForegroundColor Yellow
    $Sheet.Rows.Item(1).Insert() | Out-Null
    $Sheet.Cells.Item(1,1).Value2 = 'Software '
    $Sheet.Cells.Item(1,2).Value2 = 'Version'
    try { $Sheet.Range($Sheet.Cells.Item(1,1),$Sheet.Cells.Item(1,2)).Font.Bold = $true; $Sheet.Range($Sheet.Cells.Item(1,1),$Sheet.Cells.Item(1,2)).Borders.LineStyle = 1 } catch {}
    return (Get-SheetLayout $Sheet)
}

function Copy-FormatRangeSafe {
    param($SourceRange, $DestinationRange, $Excel)
    try { $SourceRange.Copy() | Out-Null; $DestinationRange.PasteSpecial(-4122) | Out-Null; $Excel.CutCopyMode=$false; return $true }
    catch { try { $Excel.CutCopyMode=$false } catch {}; return $false }
}

function Get-SafeFileName {
    param([string]$Name)
    $safe = ($Name -replace '[\\/:*?"<>|]', ' ').Trim()
    $safe = ($safe -replace '\s+', ' ')
    if ([string]::IsNullOrWhiteSpace($safe)) { return 'Audit' }
    return $safe
}

function Get-AuditOutputDirectory {
    $dir = Join-Path (Split-Path -Parent $MyInvocation.ScriptName) 'Audits'
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    return $dir
}

function New-RoomAuditWorkbookPath {
    param([string]$Room)
    $dir = Get-AuditOutputDirectory
    $baseName = Get-SafeFileName $Room
    $path = Join-Path $dir ($baseName + '.xlsx')
    if (-not (Test-Path $path)) { return $path }
    return (Join-Path $dir ($baseName + '_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.xlsx'))
}

function Get-StreamlinedAuditTypes {
    return @('Software','Device','Security','Network','Services','Compliance')
}

function Remove-WorksheetByName {
    param($Workbook, [string]$Name)
    $sheet = Get-WorksheetByName -Workbook $Workbook -Name $Name
    if (-not $sheet) { return $false }
    try {
        $sheet.Delete()
        return $true
    } catch {
        Write-TroubleshootingLog -Level 'WARN' -Category 'Excel' -Message ("Could not delete worksheet '{0}': {1}" -f $Name, $_.Exception.Message)
        return $false
    }
}

function Remove-ObsoleteWorkbookSheets {
    param($Workbook)
    foreach ($name in @('Audit Index','Users')) { Remove-WorksheetByName -Workbook $Workbook -Name $name | Out-Null }
}

function Remove-ExtraWorkbookSheets {
    param($Workbook)
    try {
        while ($Workbook.Worksheets.Count -gt 1) {
            $Workbook.Worksheets.Item($Workbook.Worksheets.Count).Delete()
        }
    } catch {}
}

function New-AuditWorkbook {
    param([string]$Path)
    Write-TroubleshootingLog -Category 'Excel' -Message ("Creating audit workbook: {0}" -f $Path)
    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()
        Remove-ExtraWorkbookSheets -Workbook $workbook
        $types = @(Get-StreamlinedAuditTypes)
        $workbook.Worksheets.Item(1).Name = $types[0]
        foreach ($name in @($types | Select-Object -Skip 1)) {
            $sheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $workbook.Worksheets.Item($workbook.Worksheets.Count))
            $sheet.Name = $name
        }
        $workbook.SaveAs($Path, 51)
        Write-TroubleshootingLog -Category 'Excel' -Message ("New audit workbook saved: {0}" -f $Path)
    } catch {
        Write-TroubleshootingException -ErrorRecord $_ -Context ("Failed to create audit workbook '{0}'" -f $Path)
        throw
    } finally {
        if ($workbook) {
            try { $workbook.Close($true) | Out-Null }
            catch { Write-TroubleshootingException -ErrorRecord $_ -Context 'Could not close the newly created workbook cleanly' }
        }
        if ($excel) {
            try { $excel.Quit() | Out-Null }
            catch { Write-TroubleshootingException -ErrorRecord $_ -Context 'Could not quit Excel after creating a workbook' }
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
        }
    }
}

function Get-GlobalAssetColumn {
    param($Workbook, [string]$Asset)
    $largestColumn = 2
    foreach ($sheetName in (Get-StreamlinedAuditTypes)) {
        $sheet = Get-WorksheetByName -Workbook $Workbook -Name $sheetName
        if (-not $sheet) { continue }
        $lastCol = [Math]::Max(2, (Get-UsedLastCol $sheet))
        $largestColumn = [Math]::Max($largestColumn, $lastCol)
        for ($c=3; $c -le $lastCol; $c++) {
            if ($sheet.Cells.Item(1,$c).Text.Trim() -eq $Asset) { return $c }
        }
    }
    return ($largestColumn + 1)
}

function Get-SheetBaselineMap {
    param($Sheet, [switch]$StopAtAdditionalSoftware)
    $map = @{}
    $lastRow = [Math]::Max(1, (Get-UsedLastRow $Sheet))
    for ($r=2; $r -le $lastRow; $r++) {
        $name = $Sheet.Cells.Item($r,1).Text.Trim()
        if ($StopAtAdditionalSoftware -and $name -eq 'Additional Software') { break }
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if (-not $map.ContainsKey($name)) {
            $map[$name] = $Sheet.Cells.Item($r,2).Text.Trim()
        }
    }
    return $map
}

function Get-AdditionalSoftwareNames {
    param($Sheet)
    $names = @()
    $inAdditionalSection = $false
    $lastRow = [Math]::Max(1, (Get-UsedLastRow $Sheet))
    for ($r = 2; $r -le $lastRow; $r++) {
        $name = $Sheet.Cells.Item($r,1).Text.Trim()
        if ($name -eq 'Additional Software') { $inAdditionalSection = $true; continue }
        if ($inAdditionalSection -and -not [string]::IsNullOrWhiteSpace($name)) { $names += $name }
    }
    return @($names | Select-Object -Unique)
}

function Ensure-AdditionalSoftwareSection {
    param($Sheet)
    $lastRow = [Math]::Max(1, (Get-UsedLastRow $Sheet))
    $sectionRow = 0
    for ($r = 2; $r -le $lastRow; $r++) {
        if ($Sheet.Cells.Item($r,1).Text.Trim() -eq 'Additional Software') { $sectionRow = $r; break }
    }
    if ($sectionRow -eq 0) {
        $sectionRow = $lastRow + 1
        $Sheet.Cells.Item($sectionRow,1).Value2 = 'Additional Software'
    }
    $lastCol = [Math]::Max(2, (Get-UsedLastCol $Sheet))
    $sectionRange = $Sheet.Range($Sheet.Cells.Item($sectionRow,1), $Sheet.Cells.Item($sectionRow,$lastCol))
    Set-FillColor -RangeOrCell $sectionRange -Color $SectionGrayColor
    $sectionRange.Font.Bold = $true
    $sectionRange.HorizontalAlignment = -4131
    try { $sectionRange.Borders.LineStyle = 1 } catch {}
}

function Get-ExcelColumnName {
    param([int]$ColumnNumber)
    $name = ''
    $number = $ColumnNumber
    while ($number -gt 0) {
        $number--
        $name = ([char](65 + ($number % 26))) + $name
        $number = [Math]::Floor($number / 26)
    }
    return $name
}

function Sync-ComplianceSpaceAvailableRow {
    param($Workbook)
    $device = Get-WorksheetByName -Workbook $Workbook -Name 'Device'
    $compliance = Get-WorksheetByName -Workbook $Workbook -Name 'Compliance'
    if (-not $device -or -not $compliance) { return }
    $deviceRow = 0
    $complianceRow = 0
    for ($r = 2; $r -le [Math]::Max(1, (Get-UsedLastRow $device)); $r++) {
        if ($device.Cells.Item($r,1).Text.Trim() -eq '% Space Available') { $deviceRow = $r; break }
    }
    for ($r = 2; $r -le [Math]::Max(1, (Get-UsedLastRow $compliance)); $r++) {
        if ($compliance.Cells.Item($r,1).Text.Trim() -eq '% Space Available') { $complianceRow = $r; break }
    }
    if ($deviceRow -eq 0 -or $complianceRow -eq 0) { return }
    $deviceLastCol = [Math]::Max(2, (Get-UsedLastCol $device))
    $complianceLastCol = [Math]::Max(2, (Get-UsedLastCol $compliance))
    for ($c = 2; $c -le $complianceLastCol; $c++) {
        $sourceColumn = 0
        if ($c -eq 2) {
            $sourceColumn = 2
        } else {
            $asset = $compliance.Cells.Item(1,$c).Text.Trim()
            if ([string]::IsNullOrWhiteSpace($asset)) { continue }
            for ($deviceColumn = 3; $deviceColumn -le $deviceLastCol; $deviceColumn++) {
                if ($device.Cells.Item(1,$deviceColumn).Text.Trim() -eq $asset) {
                    $sourceColumn = $deviceColumn
                    break
                }
            }
        }
        if ($sourceColumn -eq 0) {
            Write-TroubleshootingLog -Level 'WARN' -Category 'Excel' -Message ("Could not sync % Space Available for Compliance column {0}; no matching Device asset header was found." -f $c)
            continue
        }
        $cell = $compliance.Cells.Item($complianceRow,$c)
        # These cells were initially written as text. Restore General before
        # assigning a formula so Excel evaluates the link instead of displaying it.
        $cell.NumberFormat = 'General'
        $cell.FormulaR1C1 = "='Device'!R${deviceRow}C${sourceColumn}"
        $cell.HorizontalAlignment = -4131
    }
    Format-AuditSheetColumns -Sheet $compliance
}

function Set-DynamicComplianceSummary {
    param($Workbook)
    $sheet = Get-WorksheetByName -Workbook $Workbook -Name 'Compliance'
    if (-not $sheet) { return }
    $summaryRow = 0
    $lastRow = [Math]::Max(1, (Get-UsedLastRow $sheet))
    for ($r = 2; $r -le $lastRow; $r++) {
        if ($sheet.Cells.Item($r,1).Text.Trim() -eq 'Overall compliance summary') { $summaryRow = $r; break }
    }
    if ($summaryRow -ne 2 -or $lastRow -lt 3) { return }
    $sheet.Cells.Item($summaryRow,2).Value2 = 'Compliant'
    $lastCol = [Math]::Max(2, (Get-UsedLastCol $sheet))
    for ($c = 3; $c -le $lastCol; $c++) {
        if ([string]::IsNullOrWhiteSpace($sheet.Cells.Item(1,$c).Text)) { continue }
        $differenceChecks = @()
        for ($r = 3; $r -le $lastRow; $r++) {
            if ([string]::IsNullOrWhiteSpace($sheet.Cells.Item($r,1).Text)) { continue }
            if ($sheet.Cells.Item($r,1).Text.Trim() -eq '% Space Available') {
                $differenceChecks += ('AND(R{0}C{1}<>"",IFERROR(VALUE(SUBSTITUTE(R{0}C{1},"%",""))<VALUE(SUBSTITUTE(R{0}C2,"%","")),TRUE))' -f $r, $c)
            } else {
                $differenceChecks += ('AND(R{0}C{1}<>"",LOWER(TRIM(R{0}C{1}))<>LOWER(TRIM(R{0}C2)))' -f $r, $c)
            }
        }
        $summaryCell = $sheet.Cells.Item($summaryRow,$c)
        $summaryCell.NumberFormat = 'General'
        if ($differenceChecks.Count -eq 0) {
            $summaryCell.Value2 = 'Compliant'
        } else {
            # R1C1 avoids Excel interpreting an A1 reference such as C1 as an
            # invalid name, and avoids the array formula used by older versions.
            $summaryCell.FormulaR1C1 = '=IF(OR({0}),"Review differences","Compliant")' -f ($differenceChecks -join ',')
        }
    }
    Format-AuditSheetColumns -Sheet $sheet
}

function Get-ComparisonRowsFromBaseline {
    param([array]$CurrentRows, [hashtable]$BaselineMap)
    $rows = @()
    $seen = @{}
    foreach ($row in $CurrentRows) {
        if ([string]::IsNullOrWhiteSpace($row.Name)) { continue }
        $seen[$row.Name] = $true
        $baseline = if ($BaselineMap.ContainsKey($row.Name)) { $BaselineMap[$row.Name] } else { '' }
        $value = Normalize-VersionText $row.Value
        $expected = Normalize-VersionText $baseline
        $status = $row.Status
        $comment = $row.Comment
        if ([string]::IsNullOrWhiteSpace($baseline)) {
            $expected = '(not in baseline)'
            if ($status -notin @('Fail','Warning')) { $status = 'Warning' }
            $comment = "This item was not present in the baseline.`n$comment"
        } elseif ($row.Status -eq 'Fail') {
            $status = 'Fail'
        } elseif ($row.Name -eq '% Space Available') {
            $baselineNumber = 0.0
            $valueNumber = 0.0
            $baselineParsed = [double]::TryParse(($expected -replace '[^0-9.\-]', ''), [ref]$baselineNumber)
            $valueParsed = [double]::TryParse(($value -replace '[^0-9.\-]', ''), [ref]$valueNumber)
            if ($baselineParsed -and $valueParsed -and $valueNumber -ge $baselineNumber) { $status = 'Pass' }
            else {
                $status = 'Warning'
                $comment = "Baseline minimum: $baseline`nCurrent minimum: $($row.Value)`n$comment"
            }
        } elseif ($expected.Equals($value, [System.StringComparison]::OrdinalIgnoreCase)) {
            $status = 'Pass'
        } else {
            $status = 'Warning'
            $comment = "Baseline: $baseline`nCurrent: $($row.Value)`n$comment"
        }
        $rows += New-AuditRow -Name ($row.Name) -Expected $expected -Value ($row.Value) -Status $status -Comment $comment
    }
    foreach ($name in @($BaselineMap.Keys | Sort-Object)) {
        if ($seen.ContainsKey($name)) { continue }
        $rows += New-AuditRow -Name $name -Expected ($BaselineMap[$name]) -Value '(not detected)' -Status 'Fail' -Comment 'This baseline item was not detected on this device.'
    }
    return $rows
}

function Write-AlignedAuditRows {
    param(
        $Workbook,
        [string]$SheetName,
        [array]$Rows,
        [string]$Asset,
        [int]$AssetColumn,
        [bool]$BaselineMode,
        [string]$FirstHeader = 'Audit Item',
        [string]$SecondHeader = 'Baseline / Expected'
    )
    $sheet = Get-OrCreateWorksheet -Workbook $Workbook -Name $SheetName
    Ensure-AuditSheetHeader -Sheet $sheet -FirstHeader $FirstHeader -SecondHeader $SecondHeader
    $sheet.Cells.Item(1,$AssetColumn).Value2 = [string]$Asset
    $sheet.Cells.Item(1,$AssetColumn).Font.Bold = $true
    Remove-DuplicateAuditRows -Sheet $sheet | Out-Null

    $rowMap = @{}
    $lastRow = [Math]::Max(1, (Get-UsedLastRow $sheet))
    for ($r=2; $r -le $lastRow; $r++) {
        $name = $sheet.Cells.Item($r,1).Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($name) -and -not $rowMap.ContainsKey($name)) { $rowMap[$name] = $r }
    }

    foreach ($row in $Rows) {
        if ([string]::IsNullOrWhiteSpace($row.Name)) { continue }
        if ($rowMap.ContainsKey($row.Name)) {
            $r = $rowMap[$row.Name]
        } else {
            $lastRow++
            $r = $lastRow
            $rowMap[$row.Name] = $r
            $sheet.Cells.Item($r,1).Value2 = [string]$row.Name
        }
        if ($BaselineMode -or [string]::IsNullOrWhiteSpace($sheet.Cells.Item($r,2).Text)) {
            $sheet.Cells.Item($r,2).NumberFormat = '@'
            $sheet.Cells.Item($r,2).Value2 = [string]$row.Expected
        }
        Set-FillColor -RangeOrCell ($sheet.Range($sheet.Cells.Item($r,1), $sheet.Cells.Item($r,2))) -Color $OriginalBlueColor
        $displayValue = Get-AuditDisplayValue -Expected $row.Expected -Value $row.Value -Status $row.Status
        Set-AuditCellResult -Cell ($sheet.Cells.Item($r,$AssetColumn)) -Value $displayValue -Status $row.Status -Comment $row.Comment
    }
    Remove-RedundantMatchingValues -Sheet $sheet
    Format-AuditSheetColumns -Sheet $sheet
}

function Invoke-AuditWorkbook {
    param(
        [array]$InstalledApps,
        [string]$WorkbookPath = "",
        [string]$Room = "",
        [string]$Asset = "",
        [string[]]$ReferenceWorksheetNames = @(),
        [string]$SoftwareAuditMode = "Simple",
        [bool]$StartNewSheets = $false
    )
    $path = if (-not [string]::IsNullOrWhiteSpace($WorkbookPath)) { $WorkbookPath } else { Get-WorkbookPath }
    if (-not (Test-Path $path)) { throw "Workbook not found: $path" }
    $room = if (-not [string]::IsNullOrWhiteSpace($Room)) { $Room.Trim() } else { Get-RoomChoice }
    $asset = if (-not [string]::IsNullOrWhiteSpace($Asset)) { $Asset.Trim() } else { Get-AssetTagChoice }
    if ([string]::IsNullOrWhiteSpace($SoftwareAuditMode)) { $SoftwareAuditMode = 'Simple' }
    if ($SoftwareAuditMode -notin @('Simple','InDepth')) { throw "Software audit mode must be Simple or InDepth." }
    if ([string]::IsNullOrWhiteSpace($room)) { throw "Room / site / department cannot be blank." }
    if ([string]::IsNullOrWhiteSpace($asset)) { throw "Asset tag cannot be blank." }
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $workbook = $null
    try {
        $workbook = $excel.Workbooks.Open($path)
        $backupDir = Join-Path ([System.IO.Path]::GetDirectoryName($path)) 'Backups'
        if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }
        $backupPath = Join-Path $backupDir ([System.IO.Path]::GetFileNameWithoutExtension($path) + '_backup_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.xlsx')
        $workbook.SaveCopyAs($backupPath)

        $runLabel = Get-Date -Format 'MMddHHmm'
        Remove-ObsoleteWorkbookSheets -Workbook $workbook
        $auditTypes = @('Software','Device','Security','Network','Services','Compliance')
        if ($SoftwareAuditMode -eq 'InDepth') { $auditTypes = @('Software','Software Inventory','Device','Security','Network','Services','Compliance') }
        $worksheetNames = Get-AuditWorksheetNameMap -Workbook $workbook -Room $room -AuditTypes $auditTypes -StartNewSheets $StartNewSheets -RunLabel $runLabel
        $softwareSheetName = $worksheetNames['Software']
        $existingSoftwareSheet = Get-WorksheetByName -Workbook $workbook -Name $softwareSheetName

        if ($ReferenceWorksheetNames.Count -gt 0) {
            $referenceSheets = @()
            foreach ($sheetName in $ReferenceWorksheetNames) {
                $sheet = Get-WorksheetByName -Workbook $workbook -Name $sheetName
                if (-not $sheet) { throw "Reference worksheet not found: $sheetName" }
                $referenceSheets += $sheet
            }
            Write-Host "Using selected software reference worksheet(s): $($ReferenceWorksheetNames -join ', ')." -ForegroundColor Cyan
        } elseif ($existingSoftwareSheet) {
            Write-Host "Using existing software reference sheet '$softwareSheetName'." -ForegroundColor Cyan
            $referenceSheets = @($existingSoftwareSheet)
        } else {
            Write-Host "Select the worksheet that contains the required software list for room '$room'." -ForegroundColor Cyan
            $referenceSheets = @(Get-WorksheetSelection -Workbook $workbook)
        }

        $softwareReferenceRows = @(Get-SoftwareReferenceRows -Sheets $referenceSheets)
        $usingBlankReferenceSheet = ($softwareReferenceRows.Count -eq 0)
        if ($usingBlankReferenceSheet -and -not $StartNewSheets) {
            throw "No required software rows were found in the selected worksheet(s). Select 'Start new sheet set' to build a complete audit from a blank worksheet."
        }
        if ($usingBlankReferenceSheet) {
            Write-Host "No required software baseline rows found. Creating new audit sheets from live device data." -ForegroundColor Yellow
        }

        Write-Host "Collecting device, security, network, user, service, and compliance data..."
        $deviceRows = @(Get-DeviceAuditRows)
        $securityRows = @(Get-SecurityAuditRows)
        $networkRows = @(Get-NetworkAuditRows)
        $userRows = @(Get-UserAuditRows)
        $serviceRows = @(Get-ServiceAuditRows)
        $complianceRows = @(Get-ComplianceAuditRows -DeviceRows $deviceRows -SecurityRows $securityRows -ServiceRows $serviceRows -UserRows $userRows)
        if ($usingBlankReferenceSheet) {
            $softwareRows = @(Get-SimpleInstalledSoftwareRows -InstalledApps $InstalledApps -Timestamp $timestamp)
        } else {
            $softwareRows = @(Get-SoftwareAuditRows -ReferenceRows $softwareReferenceRows -InstalledApps $InstalledApps -Timestamp $timestamp)
        }
        $softwareInventoryRows = @()
        if ($SoftwareAuditMode -eq 'InDepth') {
            $softwareInventoryRows = @(Get-SoftwareInventoryRows -InstalledApps $InstalledApps -Timestamp $timestamp)
        }

        $writtenSoftwareSheet = Get-OrCreateWorksheet -Workbook $workbook -Name $softwareSheetName
        Write-AuditRowsToSheet -Sheet $writtenSoftwareSheet -Rows $softwareRows -Asset $asset -FirstHeader 'Software' -SecondHeader 'Expected Version'
        Ensure-AdditionalSoftwareSection -Sheet $writtenSoftwareSheet
        Format-AuditSheetColumns -Sheet $writtenSoftwareSheet
        if ($SoftwareAuditMode -eq 'InDepth') {
            Write-AuditRowsToSheet -Sheet (Get-OrCreateWorksheet -Workbook $workbook -Name $worksheetNames['Software Inventory']) -Rows $softwareInventoryRows -Asset $asset -FirstHeader 'Installed Software' -SecondHeader 'Publisher / Source'
        }
        Write-AuditRowsToSheet -Sheet (Get-OrCreateWorksheet -Workbook $workbook -Name $worksheetNames['Device']) -Rows $deviceRows -Asset $asset
        Write-AuditRowsToSheet -Sheet (Get-OrCreateWorksheet -Workbook $workbook -Name $worksheetNames['Security']) -Rows $securityRows -Asset $asset
        Write-AuditRowsToSheet -Sheet (Get-OrCreateWorksheet -Workbook $workbook -Name $worksheetNames['Network']) -Rows $networkRows -Asset $asset
        Write-AuditRowsToSheet -Sheet (Get-OrCreateWorksheet -Workbook $workbook -Name $worksheetNames['Services']) -Rows $serviceRows -Asset $asset
        Write-AuditRowsToSheet -Sheet (Get-OrCreateWorksheet -Workbook $workbook -Name $worksheetNames['Compliance']) -Rows $complianceRows -Asset $asset

        $softwareSame = @($softwareRows | Where-Object { $_.Status -eq 'Pass' }).Count
        $softwareMissing = @($softwareRows | Where-Object { $_.Status -eq 'Fail' }).Count
        $softwareDifferent = @($softwareRows | Where-Object { $_.Status -eq 'Warning' }).Count
        $complianceFailures = @($complianceRows | Where-Object { $_.Status -eq 'Fail' }).Count
        $complianceWarnings = @($complianceRows | Where-Object { $_.Status -eq 'Warning' }).Count

        $workbook.Save()
        Write-Host "Audit complete." -ForegroundColor Green
        Write-Host "Room: $room"
        Write-Host "Asset: $asset"
        Write-Host "Software rows checked: $($softwareRows.Count)"
        Write-Host "Software audit mode: $SoftwareAuditMode"
        Write-Host "Used blank reference sheet: $usingBlankReferenceSheet"
        if ($SoftwareAuditMode -eq 'InDepth') { Write-Host "Software inventory rows: $($softwareInventoryRows.Count)" }
        Write-Host "Started new sheet set: $StartNewSheets"
        Write-Host "Same version: $softwareSame"
        Write-Host "Missing software: $softwareMissing"
        Write-Host "Different / version unavailable: $softwareDifferent"
        Write-Host "Compliance failures: $complianceFailures"
        Write-Host "Compliance warnings: $complianceWarnings"
        Write-Host "Workbook updated: $path"
        Write-Host "Backup saved: $backupPath"
        return [PSCustomObject]@{
            WorkbookPath       = $path
            BackupPath         = $backupPath
            Room               = $room
            Asset              = $asset
            SoftwareRows       = $softwareRows.Count
            SoftwareInventoryRows = $softwareInventoryRows.Count
            SoftwareAuditMode  = $SoftwareAuditMode
            StartedNewSheetSet = $StartNewSheets
            UsedBlankReferenceSheet = $usingBlankReferenceSheet
            SoftwareSame       = $softwareSame
            SoftwareMissing    = $softwareMissing
            SoftwareDifferent  = $softwareDifferent
            ComplianceFailures = $complianceFailures
            ComplianceWarnings = $complianceWarnings
        }
    } finally {
        if ($workbook) { $workbook.Close($true) | Out-Null }
        if ($excel) { $excel.Quit() | Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    }
}

function Get-BaselineSoftwareRows {
    param([array]$InstalledApps, [string]$Timestamp)
    $rows = @()
    foreach ($app in @($InstalledApps | Sort-Object Name, Version)) {
        if ($null -eq $app -or [string]::IsNullOrWhiteSpace($app.Name)) { continue }
        $version = Normalize-VersionText $app.Version
        if ([string]::IsNullOrWhiteSpace($version)) { $version = 'Installed - version unavailable' }
        $publisher = if ([string]::IsNullOrWhiteSpace($app.Publisher)) { '(publisher unavailable)' } else { $app.Publisher }
        $source = if ([string]::IsNullOrWhiteSpace($app.Source)) { '(source unavailable)' } else { $app.Source }
        $rows += New-AuditRow `
            -Name ($app.Name) `
            -Expected $version `
            -Value $version `
            -Status 'Pass' `
            -Comment "Baseline app from first audited computer.`nPublisher: $publisher`nSource: $source`nId: $($app.Id)`nAudit time: $Timestamp"
    }
    return $rows
}

function Get-SoftwareRowsFromWorkbookBaseline {
    param($Workbook, [array]$InstalledApps, [string]$Timestamp)
    $sheet = Get-OrCreateWorksheet -Workbook $Workbook -Name 'Software'
    $baseline = Get-SheetBaselineMap -Sheet $sheet -StopAtAdditionalSoftware
    $knownAdditionalNames = @(Get-AdditionalSoftwareNames -Sheet $sheet)
    $currentRows = Get-BaselineSoftwareRows -InstalledApps $InstalledApps -Timestamp $Timestamp
    $rows = @(Get-ComparisonRowsFromBaseline -CurrentRows $currentRows -BaselineMap $baseline)
    $currentNames = @{}
    foreach ($row in $currentRows) { $currentNames[$row.Name] = $true }
    foreach ($name in $knownAdditionalNames) {
        if ($currentNames.ContainsKey($name)) { continue }
        $rows += New-AuditRow -Name $name -Expected '(not in baseline)' -Value '(not detected)' -Status 'Fail' -Comment 'Additional software detected on another audited device was not detected on this device.'
    }
    return $rows
}

function Get-CategoryBaselineRows {
    param([array]$Rows)
    return @($Rows | ForEach-Object {
        New-AuditRow -Name ($_.Name) -Expected ($_.Value) -Value ($_.Value) -Status ($_.Status) -Comment ($_.Comment)
    })
}

function Get-CategoryRowsFromWorkbookBaseline {
    param($Workbook, [string]$SheetName, [array]$CurrentRows)
    $sheet = Get-OrCreateWorksheet -Workbook $Workbook -Name $SheetName
    $baseline = Get-SheetBaselineMap -Sheet $sheet
    return @(Get-ComparisonRowsFromBaseline -CurrentRows $CurrentRows -BaselineMap $baseline)
}

function Invoke-StreamlinedDeviceAudit {
    param(
        [string]$Mode,
        [string]$Room = '',
        [string]$WorkbookPath = ''
    )
    Write-TroubleshootingLog -Category 'Audit' -Message ("Audit requested. Mode: {0}; Room: {1}; Workbook: {2}" -f $Mode, $Room, $WorkbookPath)
    if ($Mode -notin @('StartNew','UseExisting')) { throw "Unknown audit mode: $Mode" }
    if ($Mode -eq 'StartNew') {
        if ([string]::IsNullOrWhiteSpace($Room)) { throw "Room name is required to start a new audit file." }
        $WorkbookPath = New-RoomAuditWorkbookPath -Room $Room
        New-AuditWorkbook -Path $WorkbookPath
        $baselineMode = $true
    } else {
        if ([string]::IsNullOrWhiteSpace($WorkbookPath) -or -not (Test-Path $WorkbookPath)) { throw "Choose an existing audit workbook." }
        $Room = [System.IO.Path]::GetFileNameWithoutExtension($WorkbookPath)
        $baselineMode = $false
    }
    Write-TroubleshootingLog -Category 'Audit' -Message ("Audit target resolved. Workbook: {0}; Baseline mode: {1}" -f $WorkbookPath, $baselineMode)

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $asset = Get-AutoAssetTag
    if ([string]::IsNullOrWhiteSpace($asset)) {
        $asset = $env:COMPUTERNAME
        if ([string]::IsNullOrWhiteSpace($asset)) { $asset = "Device_$((Get-Date).ToString('yyyyMMddHHmmss'))" }
    }
    Write-TroubleshootingLog -Category 'Audit' -Message ("Device identifier resolved: {0}" -f $asset)

    Write-TroubleshootingLog -Category 'Collection' -Message 'Collecting installed software.'
    $installedApps = @(Get-InstalledApps)
    Write-TroubleshootingLog -Category 'Collection' -Message ("Installed software rows collected: {0}" -f $installedApps.Count)
    Write-TroubleshootingLog -Category 'Collection' -Message 'Collecting device data.'
    $deviceRowsRaw = @(Get-DeviceAuditRows)
    Write-TroubleshootingLog -Category 'Collection' -Message ("Device rows collected: {0}" -f $deviceRowsRaw.Count)
    Write-TroubleshootingLog -Category 'Collection' -Message 'Collecting security data.'
    $securityRowsRaw = @(Get-SecurityAuditRows)
    Write-TroubleshootingLog -Category 'Collection' -Message ("Security rows collected: {0}" -f $securityRowsRaw.Count)
    Write-TroubleshootingLog -Category 'Collection' -Message 'Collecting network data.'
    $networkRowsRaw = @(Get-NetworkAuditRows)
    Write-TroubleshootingLog -Category 'Collection' -Message ("Network rows collected: {0}" -f $networkRowsRaw.Count)
    Write-TroubleshootingLog -Category 'Collection' -Message 'Collecting user data.'
    $userRowsRaw = @(Get-UserAuditRows)
    Write-TroubleshootingLog -Category 'Collection' -Message ("User rows collected: {0}" -f $userRowsRaw.Count)
    Write-TroubleshootingLog -Category 'Collection' -Message 'Collecting service data.'
    $serviceRowsRaw = @(Get-ServiceAuditRows)
    Write-TroubleshootingLog -Category 'Collection' -Message ("Service rows collected: {0}" -f $serviceRowsRaw.Count)
    $complianceRowsRaw = @(Get-ComplianceAuditRows -DeviceRows $deviceRowsRaw -SecurityRows $securityRowsRaw -ServiceRows $serviceRowsRaw -UserRows $userRowsRaw)
    Write-TroubleshootingLog -Category 'Collection' -Message ("Compliance rows calculated: {0}" -f $complianceRowsRaw.Count)

    $excel = $null
    $workbook = $null
    $backupPath = ''
    try {
        Write-TroubleshootingLog -Category 'Excel' -Message 'Starting Microsoft Excel COM automation.'
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        Write-TroubleshootingLog -Category 'Excel' -Message ("Opening workbook: {0}" -f $WorkbookPath)
        $workbook = $excel.Workbooks.Open($WorkbookPath)
        Write-TroubleshootingLog -Category 'Excel' -Message ("Workbook opened; worksheets: {0}" -f $workbook.Worksheets.Count)
        if (-not $baselineMode) {
            $backupDir = Join-Path ([System.IO.Path]::GetDirectoryName($WorkbookPath)) 'Backups'
            if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }
            $backupPath = Join-Path $backupDir ([System.IO.Path]::GetFileNameWithoutExtension($WorkbookPath) + '_backup_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.xlsx')
            $workbook.SaveCopyAs($backupPath)
            Write-TroubleshootingLog -Category 'Excel' -Message ("Workbook backup saved: {0}" -f $backupPath)
        }

        Remove-ObsoleteWorkbookSheets -Workbook $workbook

        $removedRedundantRows = Remove-RedundantWorkbookAuditRows -Workbook $workbook
        Write-TroubleshootingLog -Category 'Excel' -Message ("Redundant cross-sheet rows removed: {0}" -f $removedRedundantRows)

        $assetColumn = Get-GlobalAssetColumn -Workbook $workbook -Asset $asset
        Write-TroubleshootingLog -Category 'Excel' -Message ("Global asset column selected: {0}" -f $assetColumn)

        if ($baselineMode) {
            $softwareRows = @(Get-BaselineSoftwareRows -InstalledApps $installedApps -Timestamp $timestamp)
            $deviceRows = @(Get-CategoryBaselineRows -Rows $deviceRowsRaw)
            $securityRows = @(Get-CategoryBaselineRows -Rows $securityRowsRaw)
            $networkRows = @(Get-CategoryBaselineRows -Rows $networkRowsRaw)
            $serviceRows = @(Get-CategoryBaselineRows -Rows $serviceRowsRaw)
            $complianceRows = @(Get-CategoryBaselineRows -Rows $complianceRowsRaw)
        } else {
            $softwareRows = @(Get-SoftwareRowsFromWorkbookBaseline -Workbook $workbook -InstalledApps $installedApps -Timestamp $timestamp)
            $deviceRows = @(Get-CategoryRowsFromWorkbookBaseline -Workbook $workbook -SheetName 'Device' -CurrentRows $deviceRowsRaw)
            $securityRows = @(Get-CategoryRowsFromWorkbookBaseline -Workbook $workbook -SheetName 'Security' -CurrentRows $securityRowsRaw)
            $networkRows = @(Get-CategoryRowsFromWorkbookBaseline -Workbook $workbook -SheetName 'Network' -CurrentRows $networkRowsRaw)
            $serviceRows = @(Get-CategoryRowsFromWorkbookBaseline -Workbook $workbook -SheetName 'Services' -CurrentRows $serviceRowsRaw)
            $complianceRows = @(Get-CategoryRowsFromWorkbookBaseline -Workbook $workbook -SheetName 'Compliance' -CurrentRows $complianceRowsRaw)
        }

        Write-AlignedAuditRows -Workbook $workbook -SheetName 'Software' -Rows $softwareRows -Asset $asset -AssetColumn $assetColumn -BaselineMode:$baselineMode -FirstHeader 'Software' -SecondHeader 'Baseline Version'
        $softwareSheet = Get-OrCreateWorksheet -Workbook $workbook -Name 'Software'
        Ensure-AdditionalSoftwareSection -Sheet $softwareSheet
        Format-AuditSheetColumns -Sheet $softwareSheet
        Write-AlignedAuditRows -Workbook $workbook -SheetName 'Device' -Rows $deviceRows -Asset $asset -AssetColumn $assetColumn -BaselineMode:$baselineMode
        Write-AlignedAuditRows -Workbook $workbook -SheetName 'Security' -Rows $securityRows -Asset $asset -AssetColumn $assetColumn -BaselineMode:$baselineMode
        Write-AlignedAuditRows -Workbook $workbook -SheetName 'Network' -Rows $networkRows -Asset $asset -AssetColumn $assetColumn -BaselineMode:$baselineMode
        Write-AlignedAuditRows -Workbook $workbook -SheetName 'Services' -Rows $serviceRows -Asset $asset -AssetColumn $assetColumn -BaselineMode:$baselineMode
        Write-AlignedAuditRows -Workbook $workbook -SheetName 'Compliance' -Rows $complianceRows -Asset $asset -AssetColumn $assetColumn -BaselineMode:$baselineMode
        Sync-ComplianceSpaceAvailableRow -Workbook $workbook
        Set-DynamicComplianceSummary -Workbook $workbook
        try { $workbook.Application.CalculateFull() }
        catch { Write-TroubleshootingLog -Level 'WARN' -Category 'Excel' -Message ("Could not force a final formula recalculation: {0}" -f $_.Exception.Message) }

        Write-TroubleshootingLog -Category 'Excel' -Message 'All audit worksheets updated; saving workbook.'
        $workbook.Save()
        $summary = [PSCustomObject]@{
            WorkbookPath = $WorkbookPath
            BackupPath = $backupPath
            Room = $Room
            Asset = $asset
            AssetColumn = $assetColumn
            Mode = $Mode
            BaselineDevice = $baselineMode
            SoftwareCount = $softwareRows.Count
            DeviceCount = $deviceRows.Count
            SecurityCount = $securityRows.Count
            ComplianceFailures = @($complianceRows | Where-Object { $_.Status -eq 'Fail' }).Count
            ComplianceWarnings = @($complianceRows | Where-Object { $_.Status -eq 'Warning' }).Count
        }
        Write-TroubleshootingLog -Category 'Audit' -Message ("Audit completed. Asset column: {0}; Software rows: {1}; Compliance failures: {2}; Compliance warnings: {3}" -f $summary.AssetColumn, $summary.SoftwareCount, $summary.ComplianceFailures, $summary.ComplianceWarnings)
        return $summary
    } catch {
        Write-TroubleshootingException -ErrorRecord $_ -Context ("Audit failed for workbook '{0}'" -f $WorkbookPath)
        throw
    } finally {
        Write-TroubleshootingLog -Category 'Excel' -Message 'Closing workbook and Excel.'
        if ($workbook) {
            try { $workbook.Close($true) | Out-Null }
            catch { Write-TroubleshootingException -ErrorRecord $_ -Context 'Could not close the audit workbook cleanly' }
        }
        if ($excel) {
            try { $excel.Quit() | Out-Null }
            catch { Write-TroubleshootingException -ErrorRecord $_ -Context 'Could not quit Excel cleanly' }
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
        }
        Write-TroubleshootingLog -Category 'Excel' -Message 'Excel cleanup complete.'
    }
}

function Export-InstalledAppsCsv {
    param([array]$InstalledApps)
    $dir = Split-Path -Parent $MyInvocation.ScriptName
    $exportDir = Join-Path $dir 'Exports'
    if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir | Out-Null }
    $machine = $env:COMPUTERNAME -replace '[^a-zA-Z0-9_-]','_'
    if ([string]::IsNullOrWhiteSpace($machine)) { $machine = 'Computer' }
    $path = Join-Path $exportDir ('InstalledApps_Filtered_' + $machine + '_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.csv')
    $InstalledApps | Select-Object Name, Version, Publisher, Source, Id | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
    Write-TroubleshootingLog -Category 'Export' -Message ("Installed software CSV exported: {0}; rows: {1}" -f $path, $InstalledApps.Count)
    Write-Host "Filtered Installed Apps CSV exported:" -ForegroundColor Green
    Write-Host $path
    return $path
}

function Get-DefaultWorkbookPath {
    $dir = Split-Path -Parent $MyInvocation.ScriptName
    $files = @(Get-ChildItem -Path $dir -Filter '*.xlsx' -File | Where-Object { $_.Name -notlike '~$*' -and $_.Name -notlike '*_backup_*' -and $_.FullName -notmatch '\\Backups\\' -and $_.FullName -notmatch '\\Exports\\' })
    if ($files.Count -eq 1) { return $files[0].FullName }
    return ""
}

function Get-WorkbookWorksheetNames {
    param([string]$WorkbookPath)
    if ([string]::IsNullOrWhiteSpace($WorkbookPath) -or -not (Test-Path $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $null
    try {
        $workbook = $excel.Workbooks.Open($WorkbookPath)
        $names = @()
        for ($i=1; $i -le $workbook.Worksheets.Count; $i++) {
            $names += $workbook.Worksheets.Item($i).Name
        }
        return $names
    } finally {
        if ($workbook) { $workbook.Close($false) | Out-Null }
        if ($excel) { $excel.Quit() | Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    }
}

function Invoke-ConsoleMain {
    $choice = Get-MainMenuChoice
    Write-Host "Scanning Windows Installed Apps inventory..."
    $installedApps = @(Get-InstalledApps)
    Write-Host "Detected $($installedApps.Count) filtered Installed Apps entries."
    if ($choice -eq '1') {
        $startNewSheets = Get-YesNoChoice -Prompt "Start a new sheet set instead of updating existing room sheets?" -Default $false
        $softwareAuditMode = Get-SoftwareAuditModeChoice
        Invoke-AuditWorkbook -InstalledApps $installedApps -SoftwareAuditMode $softwareAuditMode -StartNewSheets $startNewSheets | Out-Null
    } else {
        Export-InstalledAppsCsv -InstalledApps $installedApps | Out-Null
    }
}

function Show-AuditGui {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    function New-GuiLabel {
        param([string]$Text, [int]$X, [int]$Y, [int]$Width = 150)
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $Text
        $label.Location = New-Object System.Drawing.Point($X, $Y)
        $label.Size = New-Object System.Drawing.Size($Width, 22)
        return $label
    }

    function New-GuiButton {
        param([string]$Text, [int]$X, [int]$Y, [int]$Width = 120)
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $Text
        $button.Location = New-Object System.Drawing.Point($X, $Y)
        $button.Size = New-Object System.Drawing.Size($Width, 30)
        return $button
    }

    function Add-GuiLog {
        param($TextBox, [string]$Message)
        $TextBox.AppendText(("[{0}] {1}{2}" -f (Get-Date -Format 'HH:mm:ss'), $Message, [Environment]::NewLine))
        $TextBox.SelectionStart = $TextBox.TextLength
        $TextBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Company Endpoint Audit Utility"
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(820, 640)
    $form.MinimumSize = New-Object System.Drawing.Size(760, 600)

    $workbookLabel = New-GuiLabel -Text "Workbook" -X 18 -Y 22
    $workbookText = New-Object System.Windows.Forms.TextBox
    $workbookText.Location = New-Object System.Drawing.Point(150, 18)
    $workbookText.Size = New-Object System.Drawing.Size(500, 24)
    $workbookText.Anchor = "Top,Left,Right"

    $browseButton = New-GuiButton -Text "Browse" -X 660 -Y 16 -Width 120
    $browseButton.Anchor = "Top,Right"

    $roomLabel = New-GuiLabel -Text "Room / Site" -X 18 -Y 62
    $roomText = New-Object System.Windows.Forms.TextBox
    $roomText.Location = New-Object System.Drawing.Point(150, 58)
    $roomText.Size = New-Object System.Drawing.Size(220, 24)

    $assetLabel = New-GuiLabel -Text "Asset Tag" -X 400 -Y 62 -Width 90
    $assetText = New-Object System.Windows.Forms.TextBox
    $assetText.Location = New-Object System.Drawing.Point(490, 58)
    $assetText.Size = New-Object System.Drawing.Size(160, 24)
    $detectButton = New-GuiButton -Text "Detect" -X 660 -Y 55 -Width 120
    $detectButton.Anchor = "Top,Right"

    $sheetLabel = New-GuiLabel -Text "Source Sheet" -X 18 -Y 102
    $sheetCombo = New-Object System.Windows.Forms.ComboBox
    $sheetCombo.Location = New-Object System.Drawing.Point(150, 98)
    $sheetCombo.Size = New-Object System.Drawing.Size(500, 24)
    $sheetCombo.DropDownStyle = "DropDownList"
    $sheetCombo.Anchor = "Top,Left,Right"
    $refreshSheetsButton = New-GuiButton -Text "Load Sheets" -X 660 -Y 95 -Width 120
    $refreshSheetsButton.Anchor = "Top,Right"

    $newSheetsCheck = New-Object System.Windows.Forms.CheckBox
    $newSheetsCheck.Text = "Start new sheet set"
    $newSheetsCheck.Location = New-Object System.Drawing.Point(150, 135)
    $newSheetsCheck.Size = New-Object System.Drawing.Size(180, 24)

    $softwareModeLabel = New-GuiLabel -Text "Software Audit" -X 400 -Y 138 -Width 90
    $softwareModeCombo = New-Object System.Windows.Forms.ComboBox
    $softwareModeCombo.Location = New-Object System.Drawing.Point(490, 134)
    $softwareModeCombo.Size = New-Object System.Drawing.Size(160, 24)
    $softwareModeCombo.DropDownStyle = "DropDownList"
    [void]$softwareModeCombo.Items.Add('Simple')
    [void]$softwareModeCombo.Items.Add('In-depth')
    $softwareModeCombo.SelectedIndex = 0

    $runAuditButton = New-GuiButton -Text "Run Full Audit" -X 150 -Y 178 -Width 160
    $exportButton = New-GuiButton -Text "Export Apps CSV" -X 322 -Y 178 -Width 160
    $closeButton = New-GuiButton -Text "Close" -X 494 -Y 178 -Width 110

    $statusLabel = New-GuiLabel -Text "Ready" -X 18 -Y 223 -Width 760
    $statusLabel.Anchor = "Top,Left,Right"

    $logText = New-Object System.Windows.Forms.TextBox
    $logText.Location = New-Object System.Drawing.Point(18, 252)
    $logText.Size = New-Object System.Drawing.Size(762, 324)
    $logText.Multiline = $true
    $logText.ReadOnly = $true
    $logText.ScrollBars = "Vertical"
    $logText.Anchor = "Top,Bottom,Left,Right"

    $form.Controls.AddRange(@(
        $workbookLabel, $workbookText, $browseButton,
        $roomLabel, $roomText, $assetLabel, $assetText, $detectButton,
        $sheetLabel, $sheetCombo, $refreshSheetsButton,
        $newSheetsCheck, $softwareModeLabel, $softwareModeCombo,
        $runAuditButton, $exportButton, $closeButton,
        $statusLabel, $logText
    ))

    $defaultWorkbook = Get-DefaultWorkbookPath
    if (-not [string]::IsNullOrWhiteSpace($defaultWorkbook)) {
        $workbookText.Text = $defaultWorkbook
    }
    $autoAsset = Get-AutoAssetTag
    if (-not [string]::IsNullOrWhiteSpace($autoAsset)) {
        $assetText.Text = $autoAsset
    }

    $loadSheets = {
        try {
            $sheetCombo.Items.Clear()
            $path = $workbookText.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($path)) { throw "Choose a workbook first." }
            Add-GuiLog -TextBox $logText -Message "Loading workbook sheets..."
            foreach ($name in (Get-WorkbookWorksheetNames -WorkbookPath $path)) {
                [void]$sheetCombo.Items.Add($name)
            }
            if ($sheetCombo.Items.Count -gt 0) { $sheetCombo.SelectedIndex = 0 }
            Add-GuiLog -TextBox $logText -Message "Loaded $($sheetCombo.Items.Count) worksheet(s)."
        } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Load Sheets Failed", "OK", "Error") | Out-Null
            Add-GuiLog -TextBox $logText -Message "Sheet load failed: $($_.Exception.Message)"
        }
    }

    $setBusy = {
        param([bool]$Busy, [string]$Text)
        $runAuditButton.Enabled = -not $Busy
        $exportButton.Enabled = -not $Busy
        $browseButton.Enabled = -not $Busy
        $refreshSheetsButton.Enabled = -not $Busy
        $detectButton.Enabled = -not $Busy
        $newSheetsCheck.Enabled = -not $Busy
        $softwareModeCombo.Enabled = -not $Busy
        $statusLabel.Text = $Text
        [System.Windows.Forms.Application]::DoEvents()
    }

    $browseButton.Add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = "Excel workbooks (*.xlsx)|*.xlsx|All files (*.*)|*.*"
        $dialog.Title = "Choose audit workbook"
        if (-not [string]::IsNullOrWhiteSpace($workbookText.Text)) {
            try { $dialog.InitialDirectory = Split-Path -Parent $workbookText.Text } catch {}
        }
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $workbookText.Text = $dialog.FileName
            & $loadSheets
        }
    })

    $refreshSheetsButton.Add_Click({ & $loadSheets })

    $detectButton.Add_Click({
        $detected = Get-AutoAssetTag
        if ([string]::IsNullOrWhiteSpace($detected)) {
            [System.Windows.Forms.MessageBox]::Show("No usable BIOS asset tag or serial number was found.", "Asset Detection", "OK", "Information") | Out-Null
            Add-GuiLog -TextBox $logText -Message "No usable asset tag detected."
        } else {
            $assetText.Text = $detected
            Add-GuiLog -TextBox $logText -Message "Detected asset tag / serial: $detected"
        }
    })

    $runAuditButton.Add_Click({
        try {
            $path = $workbookText.Text.Trim()
            $room = $roomText.Text.Trim()
            $asset = $assetText.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path $path)) { throw "Choose a valid workbook." }
            if ([string]::IsNullOrWhiteSpace($room)) { throw "Enter a room, site, or department name." }
            if ([string]::IsNullOrWhiteSpace($asset)) { throw "Enter or detect an asset tag." }
            if ($sheetCombo.SelectedItem -eq $null) { throw "Load and select a source worksheet. It may be blank when starting a new sheet set." }
            $softwareModeDisplay = $softwareModeCombo.SelectedItem.ToString()
            $softwareMode = if ($softwareModeDisplay -eq 'In-depth') { 'InDepth' } else { 'Simple' }
            $startNewSheets = [bool]$newSheetsCheck.Checked

            & $setBusy $true "Running full audit..."
            Add-GuiLog -TextBox $logText -Message "Scanning Windows Installed Apps inventory..."
            $installedApps = @(Get-InstalledApps)
            Add-GuiLog -TextBox $logText -Message "Detected $($installedApps.Count) filtered Installed Apps entries."
            Add-GuiLog -TextBox $logText -Message "Software audit mode: $softwareModeDisplay. Start new sheet set: $startNewSheets."
            Add-GuiLog -TextBox $logText -Message "Updating workbook. Excel may stay hidden while this runs."
            $summary = Invoke-AuditWorkbook -InstalledApps $installedApps -WorkbookPath $path -Room $room -Asset $asset -ReferenceWorksheetNames @($sheetCombo.SelectedItem.ToString()) -SoftwareAuditMode $softwareMode -StartNewSheets $startNewSheets
            Add-GuiLog -TextBox $logText -Message "Audit complete for $($summary.Asset) in $($summary.Room)."
            if ($summary.UsedBlankReferenceSheet) { Add-GuiLog -TextBox $logText -Message "Blank source sheet was accepted because a new sheet set was created." }
            Add-GuiLog -TextBox $logText -Message "Software rows: $($summary.SoftwareRows), same: $($summary.SoftwareSame), missing: $($summary.SoftwareMissing), different: $($summary.SoftwareDifferent)."
            if ($summary.SoftwareAuditMode -eq 'InDepth') { Add-GuiLog -TextBox $logText -Message "Software inventory rows: $($summary.SoftwareInventoryRows)." }
            Add-GuiLog -TextBox $logText -Message "Compliance failures: $($summary.ComplianceFailures), warnings: $($summary.ComplianceWarnings)."
            Add-GuiLog -TextBox $logText -Message "Backup saved: $($summary.BackupPath)"
            [System.Windows.Forms.MessageBox]::Show("Audit complete.`n`nWorkbook updated:`n$($summary.WorkbookPath)`n`nBackup saved:`n$($summary.BackupPath)", "Audit Complete", "OK", "Information") | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Audit Failed", "OK", "Error") | Out-Null
            Add-GuiLog -TextBox $logText -Message "Audit failed: $($_.Exception.Message)"
        } finally {
            & $setBusy $false "Ready"
        }
    })

    $exportButton.Add_Click({
        try {
            & $setBusy $true "Exporting installed apps CSV..."
            Add-GuiLog -TextBox $logText -Message "Scanning Windows Installed Apps inventory..."
            $installedApps = @(Get-InstalledApps)
            Add-GuiLog -TextBox $logText -Message "Detected $($installedApps.Count) filtered Installed Apps entries."
            $path = Export-InstalledAppsCsv -InstalledApps $installedApps
            Add-GuiLog -TextBox $logText -Message "CSV exported: $path"
            [System.Windows.Forms.MessageBox]::Show("CSV exported:`n$path", "Export Complete", "OK", "Information") | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Export Failed", "OK", "Error") | Out-Null
            Add-GuiLog -TextBox $logText -Message "Export failed: $($_.Exception.Message)"
        } finally {
            & $setBusy $false "Ready"
        }
    })

    $closeButton.Add_Click({ $form.Close() })

    $form.Add_Shown({
        Add-GuiLog -TextBox $logText -Message "GUI ready."
        if (-not [string]::IsNullOrWhiteSpace($workbookText.Text)) { & $loadSheets }
    })

    [void]$form.ShowDialog()
}

function Read-GuiTextInput {
    param([string]$Title, [string]$Prompt)
    Add-Type -AssemblyName Microsoft.VisualBasic
    return [Microsoft.VisualBasic.Interaction]::InputBox($Prompt, $Title, '')
}

function Invoke-StreamlinedConsoleMain {
    Write-Host "Company Endpoint Audit Utility" -ForegroundColor Cyan
    Write-Host "  1. Start new room audit file"
    Write-Host "  2. Use existing audit file"
    do { $choice = Read-Host "Choose 1 or 2" } until ($choice -in @('1','2'))
    if ($choice -eq '1') {
        do {
            $room = (Read-Host "Room name").Trim()
            if ([string]::IsNullOrWhiteSpace($room)) { Write-Host "Room name cannot be blank." -ForegroundColor Yellow }
        } until (-not [string]::IsNullOrWhiteSpace($room))
        $summary = Invoke-StreamlinedDeviceAudit -Mode 'StartNew' -Room $room
    } else {
        $path = Read-Host "Existing audit workbook path"
        $summary = Invoke-StreamlinedDeviceAudit -Mode 'UseExisting' -WorkbookPath $path
    }
    Write-Host "Audit complete." -ForegroundColor Green
    Write-Host "Workbook: $($summary.WorkbookPath)"
    Write-Host "Asset: $($summary.Asset)"
    Write-Host "Asset column: $($summary.AssetColumn)"
    Write-Host "Compliance failures: $($summary.ComplianceFailures)"
    Write-Host "Compliance warnings: $($summary.ComplianceWarnings)"
}

function Show-StreamlinedAuditGui {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Company Endpoint Audit Utility'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(720, 480)
    $form.MinimumSize = New-Object System.Drawing.Size(680, 440)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'Device Audit'
    $title.Font = New-Object System.Drawing.Font('Segoe UI', 18, [System.Drawing.FontStyle]::Bold)
    $title.Location = New-Object System.Drawing.Point(24, 22)
    $title.Size = New-Object System.Drawing.Size(640, 42)

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Text = 'Choose one action. The app will detect this device, build the workbook, and keep asset columns aligned across every audit page.'
    $subtitle.Location = New-Object System.Drawing.Point(28, 70)
    $subtitle.Size = New-Object System.Drawing.Size(650, 42)

    $startButton = New-Object System.Windows.Forms.Button
    $startButton.Text = 'Start New Room Audit'
    $startButton.Location = New-Object System.Drawing.Point(32, 124)
    $startButton.Size = New-Object System.Drawing.Size(300, 54)

    $existingButton = New-Object System.Windows.Forms.Button
    $existingButton.Text = 'Use Existing Audit File'
    $existingButton.Location = New-Object System.Drawing.Point(360, 124)
    $existingButton.Size = New-Object System.Drawing.Size(300, 54)

    $status = New-Object System.Windows.Forms.Label
    $status.Text = 'Ready'
    $status.Location = New-Object System.Drawing.Point(32, 198)
    $status.Size = New-Object System.Drawing.Size(640, 24)

    $log = New-Object System.Windows.Forms.TextBox
    $log.Location = New-Object System.Drawing.Point(32, 230)
    $log.Size = New-Object System.Drawing.Size(628, 170)
    $log.Multiline = $true
    $log.ReadOnly = $true
    $log.ScrollBars = 'Vertical'
    $log.Anchor = 'Top,Bottom,Left,Right'

    $form.Controls.AddRange(@($title, $subtitle, $startButton, $existingButton, $status, $log))

    $writeLog = {
        param([string]$Message)
        $log.AppendText(("[{0}] {1}{2}" -f (Get-Date -Format 'HH:mm:ss'), $Message, [Environment]::NewLine))
        $log.SelectionStart = $log.TextLength
        $log.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }

    $setBusy = {
        param([bool]$Busy, [string]$Message)
        $startButton.Enabled = -not $Busy
        $existingButton.Enabled = -not $Busy
        $status.Text = $Message
        [System.Windows.Forms.Application]::DoEvents()
    }

    $finish = {
        param($Summary)
        & $writeLog "Audit complete."
        & $writeLog "Workbook: $($Summary.WorkbookPath)"
        & $writeLog "Asset: $($Summary.Asset), column $($Summary.AssetColumn)"
        & $writeLog "Software rows: $($Summary.SoftwareCount); compliance failures: $($Summary.ComplianceFailures); warnings: $($Summary.ComplianceWarnings)"
        & $writeLog "Troubleshooting log: $script:TroubleshootingLogPath"
        $backupText = if ([string]::IsNullOrWhiteSpace($Summary.BackupPath)) { '' } else { "`n`nBackup:`n$($Summary.BackupPath)" }
        [System.Windows.Forms.MessageBox]::Show("Audit complete.`n`nWorkbook:`n$($Summary.WorkbookPath)`n`nAsset: $($Summary.Asset)$backupText", 'Audit Complete', 'OK', 'Information') | Out-Null
    }

    $startButton.Add_Click({
        try {
            $room = (Read-GuiTextInput -Title 'Start New Room Audit' -Prompt 'Enter the room name. This will be used as the Excel file name.').Trim()
            if ([string]::IsNullOrWhiteSpace($room)) { return }
            & $setBusy $true 'Creating new room audit workbook...'
            & $writeLog "Starting new audit file for room: $room"
            & $writeLog 'Collecting software, device, security, network, users, services, and compliance data...'
            $summary = Invoke-StreamlinedDeviceAudit -Mode 'StartNew' -Room $room
            & $finish $summary
        } catch {
            Write-TroubleshootingException -ErrorRecord $_ -Context 'Start New Room Audit failed'
            [System.Windows.Forms.MessageBox]::Show(("{0}`n`nTroubleshooting log:`n{1}" -f $_.Exception.Message, $script:TroubleshootingLogPath), 'Audit Failed', 'OK', 'Error') | Out-Null
            & $writeLog "Audit failed: $($_.Exception.Message)"
            & $writeLog "Troubleshooting log: $script:TroubleshootingLogPath"
        } finally {
            & $setBusy $false 'Ready'
        }
    })

    $existingButton.Add_Click({
        try {
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Filter = 'Excel workbooks (*.xlsx)|*.xlsx|All files (*.*)|*.*'
            $dialog.Title = 'Open existing room audit workbook'
            if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
            & $setBusy $true 'Adding this device to existing audit workbook...'
            & $writeLog "Opening existing workbook: $($dialog.FileName)"
            & $writeLog 'Collecting software, device, security, network, users, services, and compliance data...'
            $summary = Invoke-StreamlinedDeviceAudit -Mode 'UseExisting' -WorkbookPath $dialog.FileName
            & $finish $summary
        } catch {
            Write-TroubleshootingException -ErrorRecord $_ -Context 'Use Existing Audit File failed'
            [System.Windows.Forms.MessageBox]::Show(("{0}`n`nTroubleshooting log:`n{1}" -f $_.Exception.Message, $script:TroubleshootingLogPath), 'Audit Failed', 'OK', 'Error') | Out-Null
            & $writeLog "Audit failed: $($_.Exception.Message)"
            & $writeLog "Troubleshooting log: $script:TroubleshootingLogPath"
        } finally {
            & $setBusy $false 'Ready'
        }
    })

    & $writeLog 'Ready. Start a new room audit or open an existing audit file.'
    & $writeLog "Troubleshooting log: $script:TroubleshootingLogPath"
    [void]$form.ShowDialog()
}

Initialize-TroubleshootingLog
Write-TroubleshootingLog -Category 'Startup' -Message ("Application mode: {0}" -f $(if ($Console) { 'Console' } else { 'GUI' }))
try {
    if ($Console) {
        Write-Host "Troubleshooting log: $script:TroubleshootingLogPath"
        Invoke-StreamlinedConsoleMain
    } else {
        Show-StreamlinedAuditGui
    }
    Write-TroubleshootingLog -Category 'Shutdown' -Message 'Application exited normally.'
} catch {
    Write-TroubleshootingException -ErrorRecord $_ -Context 'Fatal application error'
    if ($Console) {
        Write-Error ("Application failed: {0}`nTroubleshooting log: {1}" -f $_.Exception.Message, $script:TroubleshootingLogPath)
    } else {
        try {
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show(("The application could not continue.`n`n{0}`n`nTroubleshooting log:`n{1}" -f $_.Exception.Message, $script:TroubleshootingLogPath), 'Application Failed', 'OK', 'Error') | Out-Null
        } catch {}
    }
    exit 1
}
